VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   450
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   30
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   30
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'++----------------------------------------------------------------------------
'//   FRM_MAIN
'//     author: Stavros Sirigos
'++----------------------------------------------------------------------------
'//   - Sample usage of the Douglas-Peucker Geometry Generalization Algorithm
'//   - Examples of GDI32 drawing, Passing UDT to a Class, Timing w/out a Timer
'++----------------------------------------------------------------------------

Option Explicit

'Width of rectangle defining the drawing area
Private Const lRecWidth As Long = 10&

'Polyline constants
Private Enum ePolyType
    PL_INIT = &H0         'Original polyline
    PL_AFTER = &H1        'Simplified polyline
    
    #If False Then
        Private PL_INIT, PL_AFTER
    #End If
End Enum

'Polyline constants
Private Enum eTimerConstants
    GET_TIME = &H0
    SET_TIME = &H1
    
    #If False Then
        Private GET_TIME, SET_TIME
    #End If
End Enum

'GDI32 Text constants
Private Enum eTextConstants
    TA_BASELINE = &H18
    TA_BOTTOM = &H8
    TA_CENTER = &H6
    TA_LEFT = &H0
    TA_NOUPDATECP = &H0
    TA_RIGHT = &H2
    TA_TOP = &H0
    TA_UPDATECP = &H1
    TA_MASK = (TA_BASELINE + TA_CENTER + TA_UPDATECP)
    
    #If False Then
        Private TA_BASELINE, TA_BOTTOM, TA_CENTER, TA_LEFT, TA_NOUPDATECP, TA_RIGHT, TA_TOP, TA_UPDATECP, TA_MASK
    #End If
End Enum

'GDI32 Point
Private Type POINTAPI
    X           As Long
    Y           As Long
End Type

'A more general Point type (used in clsGeneralize)
Private Type Type_Point
    X           As Double
    Y           As Double
    Z           As Double
End Type

'GDI32 Polyline properties
Private Type Type_Polyline
    Show        As Boolean
    Color       As Long
    Style       As Long
    Width       As Long
    Poly()      As POINTAPI
End Type

Private m_cGeneralize   As clsGeneralize    'The Douglas-Peucker Algorithm object

Private m_blPolygonMode As Boolean          'True->Draws Polygon | False->Draws Polyline
Private m_lTxtPolyPos   As Long             'Positions for m_strOut updates
Private m_lTxtPos       As Long             '           "
Private m_lTxtTmrPos(1) As Long             '           "
Private m_lTwipsTol     As Long             'Tolerance for simplification in Twips
Private m_bl3D          As Boolean          'True->Use 3D version of algorithm | False-> Use 2D
Private m_IsRunning     As Boolean          'Simplification is running
Private m_uPolylines(1) As Type_Polyline    'The original(PL_INIT=0) and simplified (PL_AFTER=1) polylines
Private m_cuStart       As Currency         'Timer start time
Private m_cuStop        As Currency         'Timer stop time
Private m_cuFreq        As Currency         'Timer frequency (hardware dependent)
Private m_strOut()      As String           'Strings to be printed on screen

'GDI32 functions
Private Declare Function SetTextAlign Lib "gdi32" (ByVal hdc As Long, ByVal wFlags As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
'kernel32 - Timing related functions
Private Declare Function QueryPerformanceCounter Lib "kernel32" (cuPerfCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (cuFrequency As Currency) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long ' For compatibility

'++----------------------------------------------------------------------------
'//   FORM_KEYPRESS
'++----------------------------------------------------------------------------
Private Sub Form_KeyPress(KeyAscii As Integer)
    DoKeyPress KeyAscii 'Process keystroke
End Sub

'++----------------------------------------------------------------------------
'//   FORM_LOAD
'++----------------------------------------------------------------------------
Private Sub Form_Load()
    
    ShowAbout ' Shows the about messagebox
    
    'Form properties
    Top = &H0
    Left = &H0
    FontSize = &H9
    BackColor = &HE0E0E0
    ForeColor = &H800000
    Width = Screen.Width
    Height = Screen.Height
    
    'GDI32 Text allign behaviour
    SetTextAlign hdc, TA_LEFT Or TA_TOP Or TA_NOUPDATECP
    
    'Default simplification tolerance in twips
    m_lTwipsTol = Screen.TwipsPerPixelX * 6
    
    'Initialize the strings to be displayed on screen
    ReDim m_strOut(1 To 8)
    m_strOut(1) = "Keyboard Controls: ""A"" decrease tolerance | ""D"" increase tolerance | ""S"" default tolerance | ""Q"" del. vertex | ""C"" Line<>Polygon | ""X"" Exit"
    m_strOut(2) = "   Mouse Controls: LeftButton[+MouseMove] = add new vertex[ices] | RightButton = clear all vertices | MiddleButton = hide original polyline"
    m_strOut(3) = String$(Len(m_strOut(2)), "-")
    m_strOut(4) = "            Current Tolerance = " + CStr(m_lTwipsTol / Screen.TwipsPerPixelX) + " Pixels" + Space$(16)
    m_strOut(5) = "   Original Polyline Vertices = 0" + Space$(32)
    m_strOut(6) = "Generalized Polyline Vertices = 0" + Space$(32)
    m_strOut(7) = "Total Iterations: 0 | Maximum Stack Depth: 0"
    m_strOut(8) = " Generalize Time: " + FormatNumber(0, 4) + " sec." + _
                   " | Drawing Time: " + FormatNumber(0, 4) + " sec." + _
                   " - (" + IIf(InIDE, "interpreted", "compiled") + " code)"
    
    'Initialize the positions for the m_strOut updates
    m_lTxtPolyPos = InStr(m_strOut(2), "hide")
    m_lTxtPos = InStrRev(m_strOut(4), "=") + 2
    m_lTxtTmrPos(0) = InStr(1, m_strOut(8), ":") + 2
    m_lTxtTmrPos(1) = InStr(m_lTxtTmrPos(0), m_strOut(8), ":") + 2
    
    DrawTxt m_strOut            'Draw text on screen
    DrawRect                    'Draw rectangle defining the (user) drawing area
    
    With m_uPolylines(PL_INIT)  'Original polyline
        .Show = True
        .Style = vbSolid
        .Color = vbCyan
        .Width = 3
        ReDim .Poly(0)
    End With
    With m_uPolylines(PL_AFTER) 'Simplified polyline
        .Show = True
        .Style = vbSolid
        .Color = vbBlue
        .Width = 1
        ReDim .Poly(0)
    End With
    
    'Initialize the Douglas-Peucker Algorithm object
    Set m_cGeneralize = New clsGeneralize
    
End Sub

'++----------------------------------------------------------------------------
'//   GENERALIZE (SIMPLIFY) POLYLINE
'++----------------------------------------------------------------------------
'//   - Prepares an array of Type_Point to be passed in m_cGeneralize
'//   - Calls main m_cGeneralize.Generalize Routine
'//   - Gets and prints/draws results
'++----------------------------------------------------------------------------
Private Sub DoGeneralize()
  
  Dim uTmpPts() As Type_Point   'points to be passed in m_cGeneralize
  Dim lUB       As Long
  Dim dTimeDPL  As Double       ' ~ generalize CPU time
  Dim dTimeDraw As Double       ' ~ drawing CPU time
  Dim blUse()   As Boolean      'Array that shows for each vertex if it is used in the simplified polyline
  Dim lRet      As Long
  Dim i         As Long
  Dim k         As Long
       
    If Not m_IsRunning Then
        m_IsRunning = True
        
        With m_uPolylines(PL_INIT)
            
            lUB = UBound(.Poly)
            
            If lUB Then
                'Initialize the polyline to be simplified
                ReDim m_uPolylines(PL_AFTER).Poly(lUB)
                m_uPolylines(PL_AFTER).Poly = .Poly
                
                ReDim blUse(1 To lUB)
                                
                'Initialize the Type_Point points
                ReDim uTmpPts(1 To lUB)
                For i = 1 To lUB
                    uTmpPts(i).X = CDbl(.Poly(i).X)
                    uTmpPts(i).Y = CDbl(.Poly(i).Y)
                    uTmpPts(i).Z = CDbl((.Poly(i).X + .Poly(i).Y)) / 2 ' We will not draw in 3D anyway so put any other value
                Next i
                                
                '//Generalize
                RTime SET_TIME
                       'We can't pass a UDT from/to a private object directly.
                       'For simple UDTs, we can pass the starting memory address of the data and the ByteLength instead.
                       m_cGeneralize.LetPolyline VarPtr(uTmpPts(1)), LenB(uTmpPts(1)) * lUB
                       'Generalize
                lRet = m_cGeneralize.Generalize(CDbl(m_lTwipsTol / Screen.TwipsPerPixelX), m_bl3D)
                       'Get the results (which points stay)
                       m_cGeneralize.GetResult blUse
                dTimeDPL = RTime(GET_TIME)
                '//
                                        
                If lRet Then
                    'Rebuild simplified polyline
                    With m_uPolylines(PL_AFTER)
                        k = 0
                        For i = 1 To lUB
                            If blUse(i) Then
                                k = k + 1
                                .Poly(k) = .Poly(i)
                            End If
                        Next i
                        ReDim Preserve .Poly(k)
                    End With
                Else
                    k = lUB
                End If
            End If
        End With
        
        'Draw results, update text etc.
        Cls
        DrawRect
        
        RTime SET_TIME
        DrawPoly PL_INIT
        DrawPoly PL_AFTER
        dTimeDraw = RTime(GET_TIME)
        
        UpdateTxt lUB, k, m_cGeneralize.Iterations, m_cGeneralize.MaxDepth, dTimeDPL, dTimeDraw
        DrawTxt m_strOut
        
        Refresh
        DoEvents
        m_IsRunning = False
        
    End If
    
End Sub

'++----------------------------------------------------------------------------
'//   DRAW POLYLINE or POLYGON
'++----------------------------------------------------------------------------
'//   - Draws a Polyline or Polygon with the specified polyline
'//   - Polylines with Width > 1 are drawn in chunks -too slow otherwise (?)
'++----------------------------------------------------------------------------
Private Sub DrawPoly(ByVal eWhich As ePolyType)

  Const lSize   As Long = 32&   'Max (thick) polyline segments to draw in one Polyline call
  Dim lSteps    As Long         'Number of Polyline calls
  Dim lChunk    As Long
  Dim lStart    As Long
  Dim lEnd      As Long
  Dim lPen      As Long
  Dim hBrush    As Long
  Dim lUB       As Long
    
    With m_uPolylines(eWhich)
        If .Show Then
            lUB = UBound(.Poly)
            If lUB > 1 Then
            
                If Not m_blPolygonMode Then
                    'Draw as polyline
                    lPen = CreatePen(.Style, .Width, .Color)
                    DeleteObject (SelectObject(hdc, lPen))
                
                    If .Width > 1 Then
                        'Break the _thick_ Polyline into smaller parts to speed up drawing.
                        'Works correctly only for DrawMode=CopyPen. For other draw modes,
                        'a workaround could be to draw into a temp DC (with CopyPen)
                        'and BitBlt the result to the original DC using a mask.
                        lSteps = lUB \ lSize + Not CBool(lUB Mod lSize)
                        lEnd = 1
                        For lChunk = 0 To lSteps - 1
                            lStart = lEnd
                            lEnd = lSize * (lChunk + 1)
                            Polyline hdc, .Poly(lStart), lEnd - lStart + 1
                        Next lChunk
                        Polyline hdc, .Poly(lEnd), lUB - lEnd + 1
                    Else
                        Polyline hdc, .Poly(1), lUB
                    End If
                Else
                    'Draw as polygon
                    lPen = CreatePen(vbSolid, 1, .Color)
                    DeleteObject (SelectObject(hdc, lPen))
                    hBrush = CreateSolidBrush(.Color)
                    DeleteObject SelectObject(hdc, hBrush)
                    
                    Polygon hdc, .Poly(1), lUB
                    
                    If eWhich = PL_AFTER Then
                        With m_uPolylines(PL_INIT)
                            If .Show Then
                                'Draw "trace" for overlapping (initial) polygon
                                lPen = CreatePen(vbSolid, 1, .Color)
                                DeleteObject (SelectObject(hdc, lPen))
                                Polyline hdc, .Poly(1), UBound(.Poly)
                            End If
                        End With
                    End If
                    
                    DeleteObject hBrush
                End If
                
                DeleteObject (lPen)
            End If
        End If
    End With

End Sub

'++----------------------------------------------------------------------------
'//   DRAWS RECTANGLE FOR DEFINING THE USER'S DRAWING AREA
'++----------------------------------------------------------------------------
Private Sub DrawRect()
    
  Dim lPen      As Long
  Dim hBrush    As Long
    
    lPen = CreatePen(vbSolid, lRecWidth, &HC0C0C0)
    hBrush = CreateSolidBrush(vbWhite)
    DeleteObject SelectObject(hdc, lPen)
    DeleteObject SelectObject(hdc, hBrush)
    
    Rectangle hdc, lRecWidth \ 2, (2 + UBound(m_strOut)) * (2 + FontSize), ScaleWidth - 4, ScaleHeight - 4
    
    DeleteObject hBrush
    DeleteObject lPen
    
End Sub

'++----------------------------------------------------------------------------
'//   UPDATE TEXT
'++----------------------------------------------------------------------------
'//   - Manipulating a constant length string with Mid$ is faster than rebuilding
'//   - Concatenating with "+" is little faster than with "&" (skips string TypeCast)
'++----------------------------------------------------------------------------
Private Sub UpdateTxt(Optional ByVal lVert0 As Long, _
                      Optional ByVal lVert1 As Long, _
                      Optional ByVal lIter As Long, _
                      Optional ByVal lDepth As Long, _
                      Optional ByVal dTimeDPL As Double, _
                      Optional ByVal dTimeDraw As Double)
    
    If m_uPolylines(PL_INIT).Show Then
        Mid$(m_strOut(2), m_lTxtPolyPos, 4) = "hide"
    Else
        Mid$(m_strOut(2), m_lTxtPolyPos, 4) = "show"
    End If
    Mid$(m_strOut(5), m_lTxtPos) = (CStr(lVert0) + Space$(10))
    Mid$(m_strOut(6), m_lTxtPos) = (CStr(lVert1) + Space$(10))
    m_strOut(7) = ("Total Iterations: " + CStr(lIter)) + (" | Maximum Stack Depth: " + CStr(lDepth))
    Mid$(m_strOut(8), m_lTxtTmrPos(0), 6) = FormatNumber(dTimeDPL, 4)
    Mid$(m_strOut(8), m_lTxtTmrPos(1), 6) = FormatNumber(dTimeDraw, 4)
    
End Sub

'++----------------------------------------------------------------------------
'//   DRAW TEXT ON SCREEN
'++----------------------------------------------------------------------------
Private Sub DrawTxt(ByRef strArr() As String, _
           Optional ByVal lIndex As Long = -1)
           
    If lIndex > -1 Then
        TextOut hdc, 5, lIndex * (2 + FontSize), strArr(lIndex), Len(strArr(lIndex))
    Else
        For lIndex = LBound(strArr) To UBound(strArr)
            TextOut hdc, 5, lIndex * (2 + FontSize), strArr(lIndex), Len(strArr(lIndex))
        Next lIndex
    End If
    
End Sub

'++----------------------------------------------------------------------------
'//   CHECK IF MOUSE POINTER IS INSIDE USER'S DRAWING AREA
'++----------------------------------------------------------------------------
Private Function AllowVertex(ByVal X As Single, ByVal Y As Single) As Boolean
    
    If X > lRecWidth Then
        If X < ScaleWidth - lRecWidth - 1 Then
            If Y < ScaleHeight - lRecWidth - 1 Then
                If Y > lRecWidth \ 2 + (2 + UBound(m_strOut)) * (2 + FontSize) Then
                    'Vertex is inside the drawing area
                    'We could also use any complex Region as canvas and the PtInRegion API
                    AllowVertex = True
                End If
            End If
        End If
    End If
    
End Function

'++----------------------------------------------------------------------------
'//   CLEAR POLYLINES AND REDRAW SCREEN
'++----------------------------------------------------------------------------
Private Sub ClearAll()
    ReDim m_uPolylines(PL_INIT).Poly(0)
    ReDim m_uPolylines(PL_AFTER).Poly(0)
    Cls
    UpdateTxt
    DrawTxt m_strOut
    DrawRect
    Refresh
End Sub

'++----------------------------------------------------------------------------
'//   COUNT AND REPORT TIME INTERVALS
'++----------------------------------------------------------------------------
'//   - SET_RTIME 'resets' the timer
'//   - GET_RTIME reads ellapsed time since SET_RTIME in seconds
'//   - GetTickCount is used for compatibility purposes (very low precision)
'++----------------------------------------------------------------------------
Private Function RTime(ByVal eMode As eTimerConstants) As Double
    
    If eMode = SET_TIME Then
        QueryPerformanceFrequency m_cuFreq
        If m_cuFreq Then
            QueryPerformanceCounter m_cuStart
        Else
            m_cuStart = GetTickCount
        End If
        
    ElseIf eMode = GET_TIME Then
        QueryPerformanceCounter m_cuStop
        If m_cuFreq Then
            RTime = (m_cuStop - m_cuStart) / m_cuFreq
        Else
            RTime = (GetTickCount - m_cuStart) / 1000
        End If
        
    End If
    
End Function

'++----------------------------------------------------------------------------
'//   FORM_MOUSEMOVE
'++----------------------------------------------------------------------------
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    MousePointer = vbArrow - AllowVertex(X, Y) '(vbArrow=1, vbCross=2)
    
    If Not m_IsRunning Then
        If Button = vbLeftButton Then
            If AllowVertex(X, Y) Then
                With m_uPolylines(PL_INIT)
                    ReDim Preserve .Poly(UBound(.Poly) + 1)
                    .Poly(UBound(.Poly)).X = CLng(X)
                    .Poly(UBound(.Poly)).Y = CLng(Y)
                    
                    'Limit maximum calls per second (currently 100)
                    If RTime(GET_TIME) > 0.01 Or RTime(GET_TIME) = 0 Then
                        DoGeneralize
                        RTime (SET_TIME)
                    End If
                End With
            End If
        End If
    End If
    
End Sub

'++----------------------------------------------------------------------------
'//   FORM_MOUSEDOWN
'++----------------------------------------------------------------------------
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Not m_IsRunning Then
    
        If Button = vbLeftButton Then
            'Add new vertex
            Form_MouseMove Button, Shift, X, Y
            
        ElseIf Button = vbRightButton Then
            'Clear all vertices
            ClearAll
            
        ElseIf Button = vbMiddleButton Then
            'Toggle original polyline
            m_uPolylines(PL_INIT).Show = Not m_uPolylines(PL_INIT).Show
            DoGeneralize
            
        End If
    End If
    
End Sub

'++----------------------------------------------------------------------------
'//   PROCESS A KEYSTROKE
'++----------------------------------------------------------------------------
Private Sub DoKeyPress(ByVal KeyAscii As Integer)
    
  Dim lTwipsPPX As Long
    
    lTwipsPPX = Screen.TwipsPerPixelX
    
    If Not m_IsRunning Then
        Select Case UCase$(Chr$(KeyAscii))
        
            Case "D" 'Increase Tolerance
                If m_lTwipsTol >= lTwipsPPX Then
                    m_lTwipsTol = m_lTwipsTol + lTwipsPPX
                    m_lTwipsTol = m_lTwipsTol - m_lTwipsTol Mod lTwipsPPX
                Else
                    m_lTwipsTol = m_lTwipsTol + 1
                End If
                
            Case "A" 'Decrease Tolerance
                If m_lTwipsTol > lTwipsPPX Then
                    If m_lTwipsTol > lTwipsPPX Then
                        m_lTwipsTol = m_lTwipsTol - lTwipsPPX
                        m_lTwipsTol = m_lTwipsTol - m_lTwipsTol Mod lTwipsPPX
                    Else
                        m_lTwipsTol = m_lTwipsTol - 1
                    End If
                Else
                    If m_lTwipsTol > 1 Then
                        m_lTwipsTol = m_lTwipsTol - 1
                    End If
                End If
                
            Case "S" 'Default Tolerance
                m_lTwipsTol = Screen.TwipsPerPixelX * 6
                
            Case "Q" 'Delete last vertex
                With m_uPolylines(PL_INIT)
                    If UBound(.Poly) > 1 Then
                        ReDim Preserve .Poly(UBound(.Poly) - 1)
                    Else
                        ClearAll
                        Exit Sub
                    End If
                End With
            
            Case "C" 'Toggle draw PolyLine or Polygon
                m_blPolygonMode = Not m_blPolygonMode
            
            Case "Z" 'Toggle 3D version of algorithm
                     '3D version is 2-3 times slower than 2D version
                     'In this demo application we do not draw 3D points anyway
                m_bl3D = Not m_bl3D
                
            Case "X" 'Exit
                Unload Me
                Exit Sub
                
            Case Else
                If KeyAscii = vbKeyEscape Then 'Exit
                    Unload Me
                End If
                Exit Sub
                
        End Select
        
        'Update tolerance text
        If m_lTwipsTol > lTwipsPPX Then
            Mid$(m_strOut(4), m_lTxtPos) = (CStr(m_lTwipsTol \ lTwipsPPX) + " Pixels   ")
        Else
            If m_lTwipsTol > 1 Then
                Mid$(m_strOut(4), m_lTxtPos) = (CStr(m_lTwipsTol) + " Twips    ")
            Else
                Mid$(m_strOut(4), m_lTxtPos) = "1 Twip     "
            End If
        End If
        
        DoGeneralize
    End If
    
End Sub

'++----------------------------------------------------------------------------
'//   DETERMINES IF APPLICATION IS COMPILED (thanks to Ulli :)
'++----------------------------------------------------------------------------
Private Function InIDE(Optional ByRef blArg As Boolean) As Boolean

  Static blRet As Boolean

    blRet = blArg
    If Not blRet Then
        Debug.Assert InIDE(True)
    End If
    InIDE = blRet 'Are we in the IDE?

End Function

'++----------------------------------------------------------------------------
'//   SHOWS ABOUT MESSAGEBOX
'++----------------------------------------------------------------------------
Private Sub ShowAbout()
    MsgBox "'Douglas-Peucker' geometry generalization algorithm" + vbLf + vbLf _
           + "Implementation by Stavros Sirigos." + vbLf + "<ssirig@uth.gr>" + vbLf + vbLf _
           + "Quick instructions:" + vbLf + vbLf _
           + "- Left Mouse Button+Mouse Move: add vertices" + vbLf + vbLf _
           + "- Right Mouse Button: clear all vertices" + vbLf + vbLf _
           + "- ""A"" decrease tolerance" + vbLf + vbLf _
           + "- ""D"" increase tolerance" + vbLf + vbLf _
           + "- ""C"" choose Line or Polygon" + vbLf + vbLf _
           + "- ""X"" exit" + vbLf + vbLf _
           + "Have fun!", vbInformation, "Douglas-Peucker"
End Sub

'++----------------------------------------------------------------------------
'//   FORM_TERMINATE
'++----------------------------------------------------------------------------
Private Sub Form_Terminate()
    Set m_cGeneralize = Nothing
End Sub
