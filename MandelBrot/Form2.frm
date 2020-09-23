VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About Mandel Explorer"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   6
      X2              =   222
      Y1              =   36
      Y2              =   36
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "harietr@yahoo.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   135
      TabIndex        =   4
      Top             =   2790
      Width           =   1650
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hari krishnan G."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   135
      TabIndex        =   3
      Top             =   2565
      Width           =   1395
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "eXeption"
      Height          =   240
      Left            =   2925
      TabIndex        =   2
      Top             =   1380
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "v - 1.1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2700
      TabIndex        =   1
      Top             =   270
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mandelbrot Explorer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   180
      TabIndex        =   0
      Top             =   225
      Width           =   2325
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hari krishnan G."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   150
      TabIndex        =   6
      Top             =   2580
      Width           =   1395
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "harietr@yahoo.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   150
      TabIndex        =   5
      Top             =   2805
      Width           =   1650
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long) As Long

Private Declare Function SetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal crColor As Long) As Long

Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long

Dim XMin As Double, XMax As Double
Dim YMin As Double, YMax As Double
Dim MaxIter As Double, XRes As Double, YRes As Double
Dim m_Color1 As Long, m_Color2 As Long


Private Sub Form_Load()
    XMin = -2.1
    XMax = 0.6
    YMin = -1.2
    YMax = 1.2
    MaxIter = 30
    m_Color1 = GetSysColor(&HF&)
    m_Color2 = GetSysColor(&H12&)
End Sub

Private Sub Form_Resize()
    XRes = Me.ScaleWidth
    YRes = Me.ScaleHeight
    
    Me.Refresh
    
    mnuGenMandel_Click
End Sub

Private Sub mnuGenMandel_Click()
    Me.Cls
    GenMandelBrot
    Me.Refresh
End Sub


'*****************************************************************************
'       MANDELBROT Routines ' ABOUT's Copy!!!!
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
    
    PrepareLevels m_Color1, m_Color2, Level
    
    dx = (XMax - XMin) / (XRes - 1)
    dy = (YMax - YMin) / (YRes - 1)
    
    For iY = 0 To ((YRes - 1) / 2)
        cy = YMin + iY * dy
        For iX = 0 To (XRes - 0)
            cx = XMin + iX * dx
            iters = MIterate(cx, cy)
            If iters <> MaxIter Then
                SetPixel Me.hdc, iX, iY, Level(iters)
                SetPixel Me.hdc, iX, YRes - iY - 1, Level(iters)
            End If
        Next iX
    Next iY
    Me.Refresh
    Exit Function
errExt:
    Me.Cls
    Me.Print vbCrLf & " Error: "; Err.Number & vbCrLf & " Description : " & Err.Description
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

