VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Screen Colour Getter"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4275
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   4275
   Begin VB.Timer Timer3 
      Interval        =   5000
      Left            =   1530
      Top             =   1050
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   1530
      Top             =   1065
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   270
      LargeChange     =   5
      Left            =   780
      Max             =   255
      Min             =   155
      TabIndex        =   16
      Top             =   2280
      Value           =   155
      Width           =   2310
   End
   Begin VB.CommandButton Command1 
      Caption         =   "停止取色(&S)"
      Height          =   585
      Left            =   3105
      TabIndex        =   14
      Top             =   1590
      Width           =   1155
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2280
      Top             =   2550
   End
   Begin VB.Frame Frame2 
      Caption         =   "颜色信息"
      Height          =   1515
      Left            =   75
      TabIndex        =   5
      Top             =   705
      Width           =   2985
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "#FFFFFF"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1185
         TabIndex        =   19
         Top             =   1155
         UseMnemonic     =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "PS颜色"
         Height          =   240
         Left            =   90
         TabIndex        =   18
         Top             =   1200
         Width           =   1065
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "网页颜色"
         Height          =   240
         Left            =   105
         TabIndex        =   11
         Top             =   870
         Width           =   1065
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "十六进制颜色"
         Height          =   240
         Left            =   105
         TabIndex        =   10
         Top             =   540
         Width           =   1260
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "RGB颜色"
         Height          =   240
         Left            =   90
         TabIndex        =   9
         Top             =   195
         Width           =   1065
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&HFFFFFFFF&"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1185
         TabIndex        =   8
         Top             =   825
         UseMnemonic     =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0xFFFFFFFF"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1185
         TabIndex        =   7
         Top             =   495
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "255,255,255"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1185
         TabIndex        =   6
         Top             =   165
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "鼠标位置"
      Height          =   630
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2985
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1815
         TabIndex        =   4
         Top             =   240
         Width           =   990
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   435
         TabIndex        =   3
         Top             =   240
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Y:"
         Height          =   270
         Index           =   1
         Left            =   1305
         TabIndex        =   2
         Top             =   300
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "X:"
         Height          =   270
         Index           =   0
         Left            =   90
         TabIndex        =   1
         Top             =   300
         Width           =   390
      End
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3135
      TabIndex        =   17
      Top             =   2295
      Width           =   1080
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "透明度"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2310
      Width           =   855
   End
   Begin VB.Label ColourView 
      BorderStyle     =   1  'Fixed Single
      Height          =   1245
      Left            =   3120
      TabIndex        =   13
      Top             =   255
      Width           =   1125
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "颜色:"
      Height          =   435
      Left            =   3135
      TabIndex        =   12
      Top             =   45
      Width           =   1515
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
X As Long
Y As Long
End Type
Dim P As POINTAPI
Dim DC
Dim pi1&
Dim red As Integer
Dim green As Integer
Dim blue As Integer
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOREDRAW = &H8
Const SWP_NOMOVE = &H2
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Const VK_CONTROL = &H11
Dim blncaption As Boolean
Private Sub Command1_Click()
Select Case Left(Command1.Caption, 1)
Case "停"
Timer1.Enabled = False
Command1.Caption = "开始取色(&S)"
Case "开"
Timer1.Enabled = True
Command1.Caption = "停止取色(&S)"
End Select
End Sub
Private Sub Command1_GotFocus()
On Error Resume Next
With Me
.Caption = "Screen Colour Getter"
End With
With Me.Timer1
.Enabled = True
.Interval = 100
End With
With Me.Timer2
.Enabled = True
.Interval = 100
End With
With Me.Timer3
.Enabled = True
.Interval = 2450
End With
End Sub
Private Sub Command1_LostFocus()
On Error Resume Next
With Me
.Caption = "按CTRL+S激活窗口"
End With
With Me.Timer1
.Enabled = True
.Interval = 100
End With
With Me.Timer2
.Enabled = True
.Interval = 100
End With
With Me.Timer3
.Enabled = True
.Interval = 2450
End With
End Sub
Private Sub Form_Activate()
Command1.SetFocus
On Error Resume Next
With Me
.Caption = "Screen Colour Getter"
End With
With Me.Timer1
.Enabled = True
.Interval = 100
End With
With Me.Timer2
.Enabled = True
.Interval = 100
End With
With Me.Timer3
.Enabled = True
.Interval = 2450
End With
End Sub
Private Sub Form_Deactivate()
On Error Resume Next
With Me
.Caption = "按CTRL+S激活窗口"
End With
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Unload Me
End If
End Sub
Private Sub Form_Load()
On Error Resume Next
Dim rtn     As Long
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 230, LWA_ALPHA
Me.HScroll1.Value = 230
Label12.Caption = Me.HScroll1.Value
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, Me.Width, Me.Height, SWP_NOMOVE Or SWP_NOSIZE
DC = GetDC(0)
GetCursorPos P
Me.ColourView.BackColor = GetPixel(DC, P.X, P.Y)
pi1& = Me.ColourView.BackColor
Label2.Caption = P.X
Label3.Caption = P.Y
red = pi1& Mod 256
green = ((pi1& And &HFF00) / 256&) Mod 256&
blue = (pi1& And &HFF0000) / 65536
Label4.Caption = red & "," & green & "," & blue
Label5.Caption = "0x" & Hex(pi1&)
Select Case Len(Hex(pi1&))
Case 1
Label6.Caption = "&H0000000" & Hex(pi1&) & "&"
Case 2
Label6.Caption = "&H000000" & Hex(pi1&) & "&"
Case 3
Label6.Caption = "&H00000" & Hex(pi1&) & "&"
Case 4
Label6.Caption = "&H0000" & Hex(pi1&) & "&"
Case 5
Label6.Caption = "&H000" & Hex(pi1&) & "&"
Case 6
Label6.Caption = "&H00" & Hex(pi1&) & "&"
Case 7
Label6.Caption = "&H0" & Hex(pi1&) & "&"
Case 8
Label6.Caption = "&H" & Hex(pi1&) & "&"
Case 0
Label6.Caption = "&H00000000" & Hex(pi1&) & "&"
End Select
With Me
.KeyPreview = True
.Height = 2985
.Width = 4365
.Left = 0
.Top = 0
End With
With Me
.Caption = "Screen Colour Getter"
End With
With Me.Timer1
.Enabled = True
.Interval = 100
End With
With Me.Timer2
.Enabled = True
.Interval = 100
End With
With Me.Timer3
.Enabled = True
.Interval = 2450
End With
blncaption = True
End Sub
Private Sub Form_LostFocus()
On Error Resume Next
With Me
.Caption = "按CTRL+S激活窗口"
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
Cancel = 0
Form1.Show
End Sub
Private Sub HScroll1_Change()
Dim rtn     As Long
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, Me.HScroll1.Value, LWA_ALPHA
Label12.Caption = Me.HScroll1.Value
End Sub
Private Sub HScroll1_GotFocus()
On Error Resume Next
With Me
.Caption = "Screen Colour Getter"
End With
With Me.Timer1
.Enabled = True
.Interval = 100
End With
With Me.Timer2
.Enabled = True
.Interval = 100
End With
With Me.Timer3
.Enabled = True
.Interval = 2450
End With
End Sub
Private Sub HScroll1_LostFocus()
On Error Resume Next
With Me
.Caption = "按CTRL+S激活窗口"
End With
With Me.Timer1
.Enabled = True
.Interval = 100
End With
With Me.Timer2
.Enabled = True
.Interval = 100
End With
With Me.Timer3
.Enabled = True
.Interval = 2450
End With
End Sub
Private Sub Label12_Click()
On Error Resume Next
Dim alp As Integer
Dim oldalp As Integer
Dim rtn As Long
oldalp = Label12.Caption
alp = Val(InputBox$("请输入透明度" & vbCrLf & "范围:155-255", "Alpha", 230))
If Val(alp) = 0 Then
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, oldalp, LWA_ALPHA
Me.HScroll1.Value = oldalp
Label12.Caption = Me.HScroll1.Value
Exit Sub
End If
If 155 <= alp And alp <= 255 Then
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, alp, LWA_ALPHA
Me.HScroll1.Value = alp
Label12.Caption = Me.HScroll1.Value
Else
MsgBox "无效透明度数值", vbCritical, "Error"
End If
End Sub
Private Sub Label14_Click()
On Error Resume Next
If Trim(Label14.Caption) = "" Then Exit Sub
Clipboard.clear
Clipboard.SetText Label14.Caption
MsgBox "已经将颜色值复制到剪切板", vbExclamation, "Info"
End Sub
Private Sub Label4_Click()
On Error Resume Next
If Trim(Label4.Caption) = "" Then Exit Sub
Clipboard.clear
Clipboard.SetText Label4.Caption
MsgBox "已经将颜色值复制到剪切板", vbExclamation, "Info"
End Sub
Private Sub Label5_Click()
On Error Resume Next
If Trim(Label5.Caption) = "" Then Exit Sub
Clipboard.clear
Clipboard.SetText Label5.Caption
MsgBox "已经将颜色值复制到剪切板", vbExclamation, "Info"
End Sub
Private Sub Label6_Click()
On Error Resume Next
If Trim(Label6.Caption) = "" Then Exit Sub
Clipboard.clear
Clipboard.SetText Label6.Caption
MsgBox "已经将颜色值复制到剪切板", vbExclamation, "Info"
End Sub
Private Sub Timer1_Timer()
DC = GetDC(0)
GetCursorPos P
Me.ColourView.BackColor = GetPixel(DC, P.X, P.Y)
pi1& = Me.ColourView.BackColor
Label2.Caption = P.X
Label3.Caption = P.Y
red = pi1& Mod 256
green = ((pi1& And &HFF00) / 256&) Mod 256&
blue = (pi1& And &HFF0000) / 65536
Label4.Caption = red & "," & green & "," & blue
Label5.Caption = "0x" & Hex(pi1&)
Select Case Len(Hex(pi1&))
Case 1
Label6.Caption = "&H0000000" & Hex(pi1&) & "&"
Case 2
Label6.Caption = "&H000000" & Hex(pi1&) & "&"
Case 3
Label6.Caption = "&H00000" & Hex(pi1&) & "&"
Case 4
Label6.Caption = "&H0000" & Hex(pi1&) & "&"
Case 5
Label6.Caption = "&H000" & Hex(pi1&) & "&"
Case 6
Label6.Caption = "&H00" & Hex(pi1&) & "&"
Case 7
Label6.Caption = "&H0" & Hex(pi1&) & "&"
Case 8
Label6.Caption = "&H" & Hex(pi1&) & "&"
Case 0
Label6.Caption = "&H00000000" & Hex(pi1&) & "&"
End Select
Select Case Len(Hex(pi1&))
Case 1
Label14.Caption = "#00000" & Hex(pi1&)
Case 2
Label14.Caption = "#0000" & Hex(pi1&)
Case 3
Label14.Caption = "#000" & Hex(pi1&)
Case 4
Label14.Caption = "#00" & Hex(pi1&)
Case 5
Label14.Caption = "#0" & Hex(pi1&)
Case 6
Label14.Caption = "#" & Hex(pi1&)
Case 0
Label14.Caption = "#000000" & Hex(pi1&)
End Select
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Dim ans As Integer
If Form1.Visible = True Then
Cancel = 666
End If
Cancel = 666
ans = MsgBox("确定要退出吗?", vbQuestion + vbYesNo, "Ask")
If ans = vbYes Then
Me.Hide
Form1.Visible = True
Form1.Show
Cancel = 0
Else
Cancel = 666
End If
End Sub
Private Sub Timer2_Timer()
On Error Resume Next
Dim a(2) As Integer
a(0) = GetAsyncKeyState(VK_CONTROL)
a(1) = GetAsyncKeyState(vbKeyS)
If a(0) + a(1) = 2 Then
Me.SetFocus
End If
End Sub
Private Sub Timer3_Timer()
On Error Resume Next
If blncaption = False Then
With Me
.Caption = "按CTRL+S激活窗口"
End With
blncaption = True
Exit Sub
End If
If blncaption = True Then
With Me
.Caption = "Screen Colour Getter"
End With
blncaption = False
Exit Sub
End If
End Sub
