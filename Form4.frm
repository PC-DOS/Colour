VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Colours"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9855
   Icon            =   "Form4.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   9855
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture1 
      Height          =   4695
      Left            =   2520
      ScaleHeight     =   4635
      ScaleWidth      =   7275
      TabIndex        =   13
      Top             =   0
      Width           =   7335
      Begin VB.Image Image1 
         Height          =   4665
         Left            =   0
         MousePointer    =   2  'Cross
         Picture         =   "Form4.frx":0442
         Stretch         =   -1  'True
         Top             =   0
         Width           =   7335
      End
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   270
      LargeChange     =   5
      Left            =   780
      Max             =   255
      Min             =   155
      TabIndex        =   10
      Top             =   4800
      Value           =   155
      Width           =   7950
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "透明度"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   4830
      Width           =   855
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
      Left            =   8760
      TabIndex        =   11
      Top             =   4800
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RGB颜色:"
      Height          =   180
      Index           =   0
      Left            =   15
      TabIndex        =   9
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "颜色预览:"
      Height          =   180
      Index           =   1
      Left            =   30
      TabIndex        =   8
      Top             =   2820
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "十六进制颜色:"
      Height          =   180
      Index           =   2
      Left            =   0
      TabIndex        =   7
      Top             =   675
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "网页颜色:"
      Height          =   180
      Index           =   3
      Left            =   15
      TabIndex        =   6
      Top             =   1350
      Width           =   810
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   15
      TabIndex        =   5
      Top             =   1605
      UseMnemonic     =   0   'False
      Width           =   2490
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   0
      TabIndex        =   4
      Top             =   915
      Width           =   2490
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   15
      TabIndex        =   3
      Top             =   225
      Width           =   2490
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   1575
      Left            =   0
      TabIndex        =   2
      Top             =   3105
      Width           =   2490
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   0
      TabIndex        =   1
      Top             =   2310
      UseMnemonic     =   0   'False
      Width           =   2490
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Photoshop颜色:"
      Height          =   180
      Index           =   5
      Left            =   0
      TabIndex        =   0
      Top             =   2085
      Width           =   1260
   End
   Begin VB.Menu copy 
      Caption         =   "copy"
      Visible         =   0   'False
      Begin VB.Menu chex 
         Caption         =   "复制十六进制颜色(&H)"
      End
      Begin VB.Menu RGBC 
         Caption         =   "复制RGB颜色值(&R)"
      End
      Begin VB.Menu web 
         Caption         =   "复制网页颜色值(&W)"
      End
      Begin VB.Menu PS 
         Caption         =   "复制Photoshop颜色值(&P)"
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
Unload Me
End If
End Sub
Private Sub Form_Load()
Dim rtn     As Long
Me.KeyPreview = True
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 255, LWA_ALPHA
Me.HScroll1.Value = 255
With Me.HScroll1
.Max = 255
.Min = 155
.SmallChange = 1
.LargeChange = 5
End With
End Sub
Private Sub HScroll1_Change()
Dim rtn     As Long
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, Me.HScroll1.Value, LWA_ALPHA
Label12.Caption = Me.HScroll1.Value
End Sub
Private Sub HScroll1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
Unload Me
End If
End Sub
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Dim pi1&
Dim red As Integer
Dim green As Integer
Dim blue As Integer
Label4.Caption = "X:" & X
Label5.Caption = "Y:" & Y
pi1& = Picture1.Point(X + Image1.Left, Y + Image1.Top)
Label2.BackColor = pi1&
red = pi1& Mod 256
green = ((pi1& And &HFF00) / 256&) Mod 256&
blue = (pi1& And &HFF0000) / 65536
Label6.Caption = red & "," & green & "," & blue
Label7.Caption = "0x" & Hex(pi1&)
Select Case Len(Hex(pi1&))
Case 1
Label8.Caption = "&H0000000" & Hex(pi1&) & "&"
Case 2
Label8.Caption = "&H000000" & Hex(pi1&) & "&"
Case 3
Label8.Caption = "&H00000" & Hex(pi1&) & "&"
Case 4
Label8.Caption = "&H0000" & Hex(pi1&) & "&"
Case 5
Label8.Caption = "&H000" & Hex(pi1&) & "&"
Case 6
Label8.Caption = "&H00" & Hex(pi1&) & "&"
Case 7
Label8.Caption = "&H0" & Hex(pi1&) & "&"
Case 8
Label8.Caption = "&H" & Hex(pi1&) & "&"
Case 0
Label8.Caption = "&H00000000" & Hex(pi1&) & "&"
End Select
Select Case Len(Hex(pi1&))
Case 1
Label11.Caption = "#00000" & Hex(pi1&)
Case 2
Label11.Caption = "#0000" & Hex(pi1&)
Case 3
Label11.Caption = "#000" & Hex(pi1&)
Case 4
Label11.Caption = "#00" & Hex(pi1&)
Case 5
Label11.Caption = "#0" & Hex(pi1&)
Case 6
Label11.Caption = "#" & Hex(pi1&)
Case 0
Label11.Caption = "#000000" & Hex(pi1&)
End Select
If Button = 2 Then
PopupMenu Me.copy
End If
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
Private Sub chex_Click()
On Error Resume Next
If Trim(Label7.Caption) = "" Then Exit Sub
Clipboard.clear
Clipboard.SetText Label7.Caption
MsgBox "已经将颜色值复制到剪切板", vbExclamation, "Info"
End Sub
Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu Me.copy
End If
End Sub
Private Sub PS_Click()
On Error Resume Next
If Trim(Label11.Caption) = "" Then Exit Sub
Clipboard.clear
Clipboard.SetText Label11.Caption
MsgBox "已经将颜色值复制到剪切板", vbExclamation, "Info"
End Sub
Private Sub RGBC_Click()
On Error Resume Next
If Trim(Label6.Caption) = "" Then Exit Sub
Clipboard.clear
Clipboard.SetText Label6.Caption
MsgBox "已经将颜色值复制到剪切板", vbExclamation, "Info"
End Sub
Private Sub web_Click()
On Error Resume Next
If Trim(Label8.Caption) = "" Then Exit Sub
Clipboard.clear
Clipboard.SetText Label8.Caption
MsgBox "已经将颜色值复制到剪切板", vbExclamation, "Info"
End Sub
Private Sub Label6_Click()
On Error Resume Next
If Trim(Label6.Caption) = "" Then Exit Sub
Clipboard.clear
Clipboard.SetText Label6.Caption
MsgBox "已经将颜色值复制到剪切板", vbExclamation, "Info"
End Sub
Private Sub Label7_Click()
On Error Resume Next
If Trim(Label7.Caption) = "" Then Exit Sub
Clipboard.clear
Clipboard.SetText Label7.Caption
MsgBox "已经将颜色值复制到剪切板", vbExclamation, "Info"
End Sub
Private Sub Label8_Click()
On Error Resume Next
If Trim(Label7.Caption) = "" Then Exit Sub
Clipboard.clear
Clipboard.SetText Label8.Caption
MsgBox "已经将颜色值复制到剪切板", vbExclamation, "Info"
End Sub
Private Sub Label11_Click()
On Error Resume Next
If Trim(Label11.Caption) = "" Then Exit Sub
Clipboard.clear
Clipboard.SetText Label11.Caption
MsgBox "已经将颜色值复制到剪切板", vbExclamation, "Info"
End Sub
