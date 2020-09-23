VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000C&
   Caption         =   "Form1"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11010
   LinkTopic       =   "Form1"
   ScaleHeight     =   554
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   734
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CD 
      Left            =   5670
      Top             =   3780
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Select a Color"
   End
   Begin VB.PictureBox pJ 
      Align           =   4  'Align Right
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   8310
      Left            =   6765
      ScaleHeight     =   550
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   279
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   4245
   End
   Begin VB.PictureBox pM 
      Align           =   3  'Align Left
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   8310
      Left            =   0
      ScaleHeight     =   550
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   348
      TabIndex        =   0
      Top             =   0
      Width           =   5280
      Begin VB.Image imgCur 
         Height          =   480
         Left            =   -225
         Picture         =   "Form1.frx":0000
         Top             =   -225
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.Menu mnuGenMandel 
      Caption         =   "&Generate Mandelbrot"
   End
   Begin VB.Menu mnusettings 
      Caption         =   "&Settings"
      Begin VB.Menu mnuSquareCanvas 
         Caption         =   "S&quare Canvas"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuCol1 
         Caption         =   "Color 1"
      End
      Begin VB.Menu mnuCol2 
         Caption         =   "Color2"
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuXmin 
         Caption         =   "XMin"
      End
      Begin VB.Menu mnuXmax 
         Caption         =   "Xmax"
      End
      Begin VB.Menu mnuYmin 
         Caption         =   "YMin"
      End
      Begin VB.Menu MnuYMax 
         Caption         =   "Ymax"
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuJuliaColor 
         Caption         =   "Julia Color"
      End
      Begin VB.Menu mnujulbgcol 
         Caption         =   "Julia Background Color"
      End
      Begin VB.Menu mnujuliasetiteration 
         Caption         =   "Julia set Iteration"
      End
   End
   Begin VB.Menu mnuabt 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long) As Long

Private Declare Function SetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal crColor As Long) As Long

Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long

Private Const Pi = 3.14159265358979

Dim XMin As Double, XMax As Double
Dim YMin As Double, YMax As Double
Dim MaxIter As Double, XRes As Double, YRes As Double
Dim m_Color1 As Long, m_Color2 As Long, m_ColorJul As Long
Dim m_JulIter As Long
Dim RandLookUp(1000) As Long


Private Sub Form_Load()
    Dim i As Long
    Randomize
    For i = 0 To 1000
        RandLookUp(i) = Rnd() * 100
    Next i
'    mnuSquareCanvas.Checked = False
    XMin = -2.1
    XMax = 0.6
    YMin = -1.2
    YMax = 1.2
    MaxIter = 30
    m_Color1 = 0&
    m_Color2 = RGB(0, 255, 0)
    m_ColorJul = RGB(0, 0, 255)
    m_JulIter = 10000
    
    mnuabt_Click
End Sub

Private Sub Form_Resize()
    Dim wid As Long
    Dim xpos As Double, Ypos As Double
    xpos = imgCur.Left / pM.ScaleWidth
    Ypos = imgCur.Top / pM.ScaleHeight
    
    If mnuSquareCanvas.Checked = True Then
        pM.Align = vbAlignNone
        pJ.Align = vbAlignNone
        wid = (Me.ScaleWidth / 2)
        pM.Move 0, (Me.ScaleHeight - wid) / 2, wid, wid
        pJ.Move Me.ScaleWidth - wid, pM.Top, wid, wid
    Else
        pM.Align = vbAlignLeft
        pJ.Align = vbAlignRight
        pM.Width = (Me.ScaleWidth / 2)
        pJ.Width = pM.Width
    End If
    XRes = pM.ScaleWidth
    YRes = pM.ScaleHeight
    
    imgCur.Left = pM.ScaleWidth * xpos
    imgCur.Top = pM.ScaleHeight * Ypos
    
    Me.Refresh
    pM.Refresh
    pJ.Refresh
    
    mnuGenMandel_Click
End Sub

Private Sub imgCur_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    X = imgCur.Left + (X / Screen.TwipsPerPixelX)
    y = imgCur.Top + (y / Screen.TwipsPerPixelY)
    pM_MouseDown Button, Shift, X, y
End Sub

Private Sub imgCur_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    imgCur_MouseDown Button, Shift, X, y
End Sub

Private Sub imgCur_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    imgCur_MouseDown Button, Shift, X, y
End Sub

Private Sub mnuabt_Click()
    Form2.Show vbModal, Me
End Sub

Private Sub mnuCol1_Click()
    On Error GoTo errExt:
    CD.Flags = &HFF&
    CD.Color = m_Color1
    CD.ShowColor
    m_Color1 = CD.Color
    mnuGenMandel_Click
errExt:
End Sub

Private Sub mnuCol2_Click()
    On Error GoTo errExt:
    CD.Flags = &HFF&
    CD.Color = m_Color2
    CD.ShowColor
    m_Color2 = CD.Color
    mnuGenMandel_Click
errExt:
End Sub

Private Sub mnuGenMandel_Click()
    imgCur.Visible = False
    pM.Cls
    GenMandelBrot
    pM.Refresh
    pM_MouseDown vbLeftButton, 0, 0, 0
    imgCur.Visible = True
End Sub

Private Sub mnujulbgcol_Click()
    On Error GoTo errExt:
    CD.Flags = &HFF&
    CD.Color = pJ.BackColor
    CD.ShowColor
    pJ.BackColor = CD.Color
    mnuGenMandel_Click
errExt:
End Sub

Private Sub mnuJuliaColor_Click()
    On Error GoTo errExt:
    CD.Flags = &HFF&
    CD.Color = m_ColorJul
    CD.ShowColor
    m_ColorJul = CD.Color
    pM_MouseDown vbLeftButton, 0, 0, 0
errExt:
End Sub

Private Sub mnujuliasetiteration_Click()
    On Error GoTo errExt:
    Dim X As Long
    X = Val(InputBox("Enter how many points to approximate Julia set ?" & vbCrLf & "(default is 1000)", "Value of Julia Iteration", m_JulIter))
    m_JulIter = IIf(X >= 10, X, m_JulIter)
    pM_MouseDown vbLeftButton, 0, 0, 0
errExt:
End Sub

Private Sub mnuSquareCanvas_Click()
    Dim wid As Long
    Dim xpos As Double, Ypos As Double
    xpos = imgCur.Left / pM.ScaleWidth
    Ypos = imgCur.Top / pM.ScaleHeight
    
    If mnuSquareCanvas.Checked = False Then
        pM.Align = vbAlignNone
        pJ.Align = vbAlignNone
        wid = (Me.ScaleWidth / 2)
        pM.Move 0, (Me.ScaleHeight - wid) / 2, wid, wid
        pJ.Move Me.ScaleWidth - wid, pM.Top, wid, wid
        mnuSquareCanvas.Checked = True
    Else
        pM.Align = vbAlignLeft
        pJ.Align = vbAlignRight
        pM.Width = (Me.ScaleWidth / 2)
        pJ.Width = pM.Width
        mnuSquareCanvas.Checked = False
    End If
    XRes = pM.ScaleWidth
    YRes = pM.ScaleHeight
    
    imgCur.Left = pM.ScaleWidth * xpos
    imgCur.Top = pM.ScaleHeight * Ypos
    
    Me.Refresh
    pM.Refresh
    pJ.Refresh
    mnuGenMandel_Click
End Sub

Private Sub mnuXmax_Click()
    On Error GoTo errExt:
    XMax = Val(InputBox("Enter Xmax ?" & vbCrLf & "(default is 0.6)", "Value of XMAX", XMax))
    mnuGenMandel_Click
errExt:
End Sub

Private Sub mnuXmin_Click()
    On Error GoTo errExt:
    XMin = Val(InputBox("Enter Xmin ?" & vbCrLf & "(default is -2.1)", "Value of XMin", XMin))
    mnuGenMandel_Click
errExt:
End Sub

Private Sub MnuYMax_Click()
    On Error GoTo errExt:
    YMax = Val(InputBox("Enter YMax ?" & vbCrLf & "(default is 1.2)", "Value of YMax", YMax))
    mnuGenMandel_Click
errExt:
End Sub

Private Sub mnuYmin_Click()
    On Error GoTo errExt:
    YMin = Val(InputBox("Enter YMin ?" & vbCrLf & "(default is -1.2)", "Value of YMin", YMin))
    mnuGenMandel_Click
errExt:
End Sub

Private Sub pM_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Step As Long
    
    Step = IIf((Shift And vbCtrlMask) = vbCtrlMask, 5, 1)
    Select Case KeyCode
        Case vbKeyLeft
                imgCur.Left = IIf(imgCur.Left - Step < 0, 0, imgCur.Left - Step)
        Case vbKeyRight
                imgCur.Left = IIf(imgCur.Left + Step > (pM.ScaleWidth - imgCur.Width), (pM.ScaleWidth - imgCur.Width), imgCur.Left + Step)
        Case vbKeyUp
                imgCur.Top = IIf(imgCur.Top - Step < 0, 0, imgCur.Top - Step)
        Case vbKeyDown
                imgCur.Top = IIf(imgCur.Top + Step > (pM.ScaleHeight - imgCur.Height), (pM.ScaleHeight - imgCur.Height), imgCur.Top + Step)
    End Select
    pM_MouseDown vbLeftButton, Shift, imgCur.Left + 15, imgCur.Top + 15
End Sub

Private Sub pM_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    On Error GoTo errExt
    Dim cx As Double, cy As Double, dx As Double, dy As Double
    If Button = vbLeftButton Then
        dx = (XMax - XMin) / (XRes - 1)
        dy = (YMax - YMin) / (YRes - 1)
        cx = XMin + dx * (imgCur.Left + imgCur.Width / 2 + 1)
        cy = YMin + dy * (imgCur.Top + imgCur.Height / 2 + 1)
        CalcJulSet cx, cy
        imgCur.Move X - imgCur.Width / 2 + 1, y - imgCur.Height / 2 + 1
    End If
errExt:
End Sub

Private Sub pM_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    pM_MouseDown Button, Shift, X, y
End Sub

Private Sub pM_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    pM_MouseDown Button, Shift, X, y
End Sub



'*****************************************************************************
'       MANDELBROT Routines
'*****************************************************************************

' This is the iterative routine to calculate the equation Z(n) = Z(n-1)^2 + c
' Where Zn , Z(n-1) and C are all complex numbers
' MIterate() iterates until maxIteration is reached
' or the function blows up beyond certain limit!
Private Function MIterate(ByVal cx As Double, ByVal cy As Double) As Long
    On Error GoTo errExt
    Dim iters As Long, X As Double, y As Double, x2 As Double, y2 As Double
    Dim temp As Double
    X = cx
    x2 = X * X
    y = cy
    y2 = y * y
    iters = 0
    While (iters < MaxIter) And (x2 + y2 < 4)
        temp = cx + x2 - y2
        y = cy + 2 * X * y
        y2 = y * y
        X = temp
        x2 = X * X
        iters = iters + 1
    Wend
    MIterate = iters
errExt:
End Function

' Draws a Mandelbrotset using the above MIterate function
' Levels of colors designate the residual value of Iters after
' the function had blown up. Center color (0-default) indicates the
' region where equation holds stable upto Maxiter.
Public Function GenMandelBrot()
    On Error GoTo errExt
    Dim iX As Long, iY As Long, iters As Long
    Dim cx As Double, cy As Double, dx As Double, dy As Double
    Dim Level() As Long
    
    Me.Caption = "Mandel Explorer v-1.1    [ Calculating ]"
    
    PrepareLevels m_Color1, m_Color2, Level
    
    dx = (XMax - XMin) / (XRes - 1)
    dy = (YMax - YMin) / (YRes - 1)
    
    For iY = 0 To ((YRes - 1) / 2)
        cy = YMin + iY * dy
        For iX = 0 To (XRes - 0)
            cx = XMin + iX * dx
            iters = MIterate(cx, cy)
            If iters = MaxIter Then
                SetPixel pM.hdc, iX, iY, RGB(0, 0, 0)
                SetPixel pM.hdc, iX, YRes - iY - 1, RGB(0, 0, 0)
            Else
                SetPixel pM.hdc, iX, iY, Level(iters)
                SetPixel pM.hdc, iX, YRes - iY - 1, Level(iters)
            End If
        Next iX
        DoEvents
    Next iY
    pM.Refresh
    Me.Caption = "Mandel Explorer v-1.1    [ Done ]"
    Exit Function
errExt:
    pM.Cls
    pM.Print vbCrLf & " Error: "; Err.Number & vbCrLf & " Description : " & Err.Description
End Function

' This routine Calculates the Julia set for the Specified Mandelbrot value.
Private Function CalcJulSet(ByVal cx As Double, ByVal cy As Double)
    On Error GoTo errExt
    Dim Xp As Long, yP As Long
    Dim dx As Double, dy As Double
    Dim r As Double
    Dim theta As Double
    Dim X As Double, y As Double
    Dim i As Long
    
    X = 0
    y = 0
    pJ.Cls
    For i = 0& To m_JulIter
        dx = X - cx
        dy = y - cy
        If dx > 0 Then
            theta = Atn(dy / dx) * 0.5
        Else
            If dx < 0 Then
                theta = (Pi + Atn(dy / dx)) * 0.5
            Else
                theta = Pi * 0.25
            End If
        End If
        r = Sqr(Sqr(dx * dx + dy * dy))
        If vRandom() < 50 Then
            r = -r
        End If
        X = r * Cos(theta)
        y = r * Sin(theta)
        Xp = (XRes / 2) + CLng(X * (XRes / 3.5))
        yP = (YRes / 2) + CLng(y * (YRes / 3.5))
        SetPixel pJ.hdc, Xp, yP, m_ColorJul
    Next i
    pJ.Refresh
    Exit Function
errExt:
    pJ.Cls
    pJ.Print vbCrLf & " Error: "; Err.Number & vbCrLf & " Description : " & Err.Description
End Function

' Function to Blend between two colors
' These Colors can be customized using the Settings menu.
Private Function PrepareLevels(c1 As Long, c2 As Long, l() As Long, Optional ByVal nLevel As Long = 0)
    Dim r As Double, g As Double, b As Double
    Dim rs As Double, gs As Double, bs As Double
    Dim rr As Double, gg As Double, bb As Double
    Dim i As Long
    
    If nLevel <= 0 Then nLevel = MaxIter
    
    ReDim l(nLevel) As Long
    
    toRGB c1, r, g, b
    toRGB c2, rr, gg, bb
    rs = (rr - r) / (nLevel + 1)
    gs = (gg - g) / (nLevel + 1)
    bs = (bb - b) / (nLevel + 1)
    For i = 0 To nLevel
        l(i) = RGB(r, g, b)
        r = r + rs
        g = g + gs
        b = b + bs
    Next i
errExt:
End Function

' Function to parse a Long color Windows native value into
'  corresponding Red, Green, and Blue values.
Public Function toRGB(ByVal c, ByRef r, ByRef g, ByRef b)
    On Error GoTo errExt
    r = CLng(c And &HFF&)
    g = CLng((c And &HFF00&) / &H100&)
    b = CLng((c And &HFF0000) / &H10000)
errExt:
End Function

Private Function vRandom() As Long
    On Error GoTo errExt
    Static Index As Long
    Index = IIf(Index < 0, Index = 0, IIf(Index >= 1000, 0, Index + 1))
    vRandom = RandLookUp(Index)
errExt:
End Function
