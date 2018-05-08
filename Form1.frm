VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Colour Getter Version 1.0 - PC-DOS Workshop"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   12360
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   12360
   Begin VB.Frame Frame1 
      Caption         =   "图片取色器"
      Height          =   7620
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   12285
      Begin VB.CommandButton Command2 
         Caption         =   "清除已经装载的图片(&C)"
         Enabled         =   0   'False
         Height          =   735
         Left            =   9945
         Picture         =   "Form1.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1140
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "装载图片(&L)"
         Default         =   -1  'True
         Height          =   915
         Left            =   9930
         Picture         =   "Form1.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   180
         Width           =   2310
      End
      Begin VB.HScrollBar HScroll1 
         Enabled         =   0   'False
         Height          =   255
         Left            =   60
         TabIndex        =   3
         Top             =   7275
         Width           =   9555
      End
      Begin VB.VScrollBar VScroll1 
         Enabled         =   0   'False
         Height          =   7080
         Left            =   9660
         TabIndex        =   2
         Top             =   180
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Enabled         =   0   'False
         Height          =   7065
         Left            =   60
         ScaleHeight     =   7005
         ScaleWidth      =   9525
         TabIndex        =   1
         Top             =   180
         Width           =   9585
         Begin VB.Image Image1 
            Enabled         =   0   'False
            Height          =   2865
            Left            =   60
            MouseIcon       =   "Form1.frx":0CC6
            MousePointer    =   2  'Cross
            Top             =   60
            Width           =   4695
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Photoshop颜色:"
         Height          =   180
         Index           =   5
         Left            =   9960
         TabIndex        =   19
         Top             =   4035
         Width           =   1260
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
         Left            =   9960
         TabIndex        =   18
         Top             =   4260
         UseMnemonic     =   0   'False
         Width           =   2265
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "PCDOS Workshop Presents"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1245
         Left            =   10995
         TabIndex        =   17
         Top             =   6300
         Width           =   1230
      End
      Begin VB.Image Image2 
         Height          =   1245
         Left            =   10005
         Picture         =   "Form1.frx":1108
         Top             =   6300
         Width           =   1065
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   10815
         TabIndex        =   16
         Top             =   4740
         Width           =   1425
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "X:"
         Height          =   360
         Left            =   9990
         TabIndex        =   15
         Top             =   5475
         Width           =   2250
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Y:"
         Height          =   360
         Left            =   9990
         TabIndex        =   14
         Top             =   5880
         Width           =   2250
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
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
         Left            =   9975
         TabIndex        =   13
         Top             =   2175
         Width           =   2265
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
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
         Left            =   9960
         TabIndex        =   12
         Top             =   2865
         Width           =   2280
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
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
         Left            =   9975
         TabIndex        =   11
         Top             =   3555
         UseMnemonic     =   0   'False
         Width           =   2265
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "网页颜色:"
         Height          =   180
         Index           =   3
         Left            =   9975
         TabIndex        =   10
         Top             =   3300
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "十六进制颜色:"
         Height          =   180
         Index           =   2
         Left            =   9960
         TabIndex        =   9
         Top             =   2625
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "鼠标位置(单位:Twip)"
         Height          =   180
         Left            =   9990
         TabIndex        =   8
         Top             =   5220
         Width           =   1710
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "颜色预览:"
         Height          =   180
         Index           =   1
         Left            =   9990
         TabIndex        =   7
         Top             =   4770
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RGB颜色:"
         Height          =   180
         Index           =   0
         Left            =   9975
         TabIndex        =   6
         Top             =   1950
         Width           =   720
      End
   End
   Begin VB.Image IC 
      Height          =   495
      Left            =   5580
      Top             =   3615
      Width           =   1215
   End
   Begin VB.Menu program 
      Caption         =   "程序(&P)"
      Begin VB.Menu load 
         Caption         =   "装载图片(&L)"
         Shortcut        =   ^L
      End
      Begin VB.Menu clear 
         Caption         =   "清除已经装载的图片(&C)"
         Shortcut        =   ^D
      End
      Begin VB.Menu fg1 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "退出(&E)"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu more 
      Caption         =   "更多颜色工具(&M)"
      Begin VB.Menu RGB 
         Caption         =   "RGB颜色调版(&R)..."
         Shortcut        =   ^R
      End
      Begin VB.Menu screen 
         Caption         =   "屏幕取色器(&S)..."
         Shortcut        =   ^S
      End
      Begin VB.Menu colourget 
         Caption         =   "RGB颜色吸管(&G)"
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu copy 
      Caption         =   "图像颜色数据复制(&O)"
      Begin VB.Menu chex 
         Caption         =   "复制16进制值(&H)"
         Shortcut        =   ^H
      End
      Begin VB.Menu RGBC 
         Caption         =   "复制RGB值(&R)"
         Shortcut        =   ^C
      End
      Begin VB.Menu web 
         Caption         =   "复制网页颜色值(&W)"
         Shortcut        =   ^W
      End
      Begin VB.Menu PS 
         Caption         =   "复制Photoshop颜色值(&P)"
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu info 
      Caption         =   "提示:为保证稳定性,部分功能在打开2个及以上实例时不可用"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CommonDialog1 As New CCommonDialog
Private Sub clear_Click()
On Error Resume Next
Dim answer As Integer
answer = MsgBox("确定清除图像吗?", vbQuestion + vbYesNo, "ASK")
If answer = vbYes Then
Image1.Picture = LoadPicture("")
Image1.Enabled = False
Picture1.Enabled = False
Label6.Caption = ""
Label2.BackColor = vbWhite
Me.HScroll1.Enabled = False
Me.VScroll1.Enabled = False
Command2.Enabled = False
Label6.BackColor = vbWhite
Label7.Caption = ""
Label8.Caption = ""
Label11.Caption = ""
Label6.Enabled = False
Label8.Enabled = False
Label7.Enabled = False
Label11.Enabled = False
copy.Enabled = False
Me.clear.Enabled = False
Else
Exit Sub
End If
End Sub
Private Sub colourget_Click()
Form4.Show
End Sub
Private Sub Command1_Click()
On Error GoTo ep
With CommonDialog1
.CancelError = True
.DialogTitle = "选择要装载的图片文件"
.Filter = "Pictures(*.jpg;*.jpeg;*.bmp;*.bip;*.gif;*.wmf)|*.jpg;*.jpeg;*.bmp;*.bip;*.gif;*.wmf"
.ShowModalWindow = True
.hWndCall = hWnd
.CancelError = True
Dim IsCanceled As Boolean
IsCanceled = .ShowOpen
End With
If IsCanceled = False Then
Exit Sub
End If
If Trim(CommonDialog1.FileName) <> "" Then
If Dir(CommonDialog1.FileName) <> "" Then
If Image1.Picture = IC.Picture Then
Image1.Picture = LoadPicture(CommonDialog1.FileName)
Picture1.Enabled = True
Image1.Enabled = True
Me.Image1.Left = 0
Me.Image1.Top = 0
Me.HScroll1.Enabled = True
Me.VScroll1.Enabled = True
If Me.Image1.Height <= Me.Picture1.Height And Me.Image1.Width <= Me.Picture1.Width Then
Me.Image1.Left = Me.Picture1.Width / 2 - Me.Image1.Width / 2
Me.Image1.Top = Me.Picture1.Height / 2 - Me.Image1.Height / 2
Me.VScroll1.Enabled = False
Me.HScroll1.Enabled = False
Command2.Enabled = True
clear.Enabled = True
Label6.Enabled = True
Label8.Enabled = True
Label7.Enabled = True
Label11.Enabled = True
copy.Enabled = True
Exit Sub
End If
If Me.Image1.Width <= Me.Picture1.Width Then
Me.HScroll1.Enabled = False
Me.VScroll1.Enabled = True
Me.Image1.Left = Me.Picture1.Width / 2 - Me.Image1.Width / 2
End If
If Me.Image1.Height <= Me.Picture1.Height Then
Me.VScroll1.Enabled = False
Me.HScroll1.Enabled = True
Me.Image1.Top = Me.Picture1.Height / 2 - Me.Image1.Height / 2
End If
Me.VScroll1.Max = Me.Image1.Height - Me.Picture1.Height
Me.HScroll1.Max = Me.Image1.Width - Me.Picture1.Width
Me.VScroll1.Value = 0
Me.HScroll1.Value = 0
Me.HScroll1.LargeChange = Image1.Width / 10
Me.HScroll1.SmallChange = Image1.Width / 50
Me.VScroll1.LargeChange = Image1.Width / 10
Me.VScroll1.SmallChange = Image1.Width / 50
Command2.Enabled = True
clear.Enabled = True
Label6.Enabled = True
Label8.Enabled = True
Label7.Enabled = True
Label11.Enabled = True
copy.Enabled = True
Else
Dim ans As Integer
ans = MsgBox("已经导入了图片,是否替换?", vbQuestion + vbYesNo, "Ask")
If ans = vbYes Then
Image1.Picture = LoadPicture(CommonDialog1.FileName)
Picture1.Enabled = True
Image1.Enabled = True
Me.Image1.Left = 0
Me.Image1.Top = 0
Me.HScroll1.Enabled = True
Me.VScroll1.Enabled = True
If Me.Image1.Height <= Me.Picture1.Height And Me.Image1.Width <= Me.Picture1.Width Then
Me.Image1.Left = Me.Picture1.Width / 2 - Me.Image1.Width / 2
Me.Image1.Top = Me.Picture1.Height / 2 - Me.Image1.Height / 2
Me.VScroll1.Enabled = False
Me.HScroll1.Enabled = False
Command2.Enabled = True
clear.Enabled = True
Label6.Enabled = True
Label8.Enabled = True
Label7.Enabled = True
Label11.Enabled = True
copy.Enabled = True
Exit Sub
End If
If Me.Image1.Width <= Me.Picture1.Width Then
Me.HScroll1.Enabled = False
Me.VScroll1.Enabled = True
Me.Image1.Left = Me.Picture1.Width / 2 - Me.Image1.Width / 2
End If
If Me.Image1.Height <= Me.Picture1.Height Then
Me.VScroll1.Enabled = False
Me.HScroll1.Enabled = True
Me.Image1.Top = Me.Picture1.Height / 2 - Me.Image1.Height / 2
End If
Me.VScroll1.Max = Me.Image1.Height - Me.Picture1.Height
Me.HScroll1.Max = Me.Image1.Width - Me.Picture1.Width
Me.VScroll1.Value = 0
Me.HScroll1.Value = 0
Me.HScroll1.LargeChange = Image1.Width / 10
Me.HScroll1.SmallChange = Image1.Width / 50
Me.VScroll1.LargeChange = Image1.Width / 10
Me.VScroll1.SmallChange = Image1.Width / 50
Command2.Enabled = True
clear.Enabled = True
Label6.Enabled = True
Label8.Enabled = True
Label7.Enabled = True
copy.Enabled = True
Else
Exit Sub
End If
End If
Else
MsgBox "找不到文件", vbCritical, "Error"
Exit Sub
End If
End If
Exit Sub
ep:
If Err.Number <> 32755 Then
MsgBox "发生错误:" & Chr(13) & Err.Description, vbCritical, "Error"
If Err.Description = "溢出" Or Err.Number = 6 Then
Me.VScroll1.Max = 32755
Me.VScroll1.Min = 0
Me.HScroll1.Max = 32755
Me.HScroll1.Min = 0
Me.VScroll1.Value = 0
Me.HScroll1.Value = 0
Me.HScroll1.LargeChange = 1000
Me.HScroll1.SmallChange = 100
Me.VScroll1.LargeChange = 1000
Me.VScroll1.SmallChange = 100
Command2.Enabled = True
clear.Enabled = True
Label6.Enabled = True
Label8.Enabled = True
Label7.Enabled = True
Label11.Enabled = True
copy.Enabled = True
Else
Me.VScroll1.Max = 0
Me.VScroll1.Min = 0
Me.HScroll1.Max = 0
Me.HScroll1.Min = 0
Me.VScroll1.Value = 0
Me.HScroll1.Value = 0
Me.HScroll1.LargeChange = 1000
Me.HScroll1.SmallChange = 100
Me.VScroll1.LargeChange = 1000
Me.VScroll1.SmallChange = 100
Me.HScroll1.Enabled = False
Me.VScroll1.Enabled = False
Command2.Enabled = False
clear.Enabled = False
Label6.Enabled = False
Label8.Enabled = False
Label7.Enabled = False
Label11.Enabled = False
copy.Enabled = False
Me.Image1.Picture = LoadPicture()
Image1.Picture = LoadPicture("")
Image1.Enabled = False
Picture1.Enabled = False
Label6.Caption = ""
Label2.BackColor = vbWhite
Me.HScroll1.Enabled = False
Me.VScroll1.Enabled = False
Command2.Enabled = False
Label6.BackColor = vbWhite
Label7.Caption = ""
Label8.Caption = ""
Label6.Enabled = False
Label8.Enabled = False
Label7.Enabled = False
Label11.Enabled = False
copy.Enabled = False
Me.clear.Enabled = False
Label11.Caption = ""
End If
Else
Exit Sub
End If
If Err.Number = 481 Or Err.Number = 53 Then
If Image1.Picture = LoadPicture() Then
Me.VScroll1.Max = 0
Me.VScroll1.Min = 0
Me.HScroll1.Max = 0
Me.HScroll1.Min = 0
Me.VScroll1.Value = 0
Me.HScroll1.Value = 0
Me.HScroll1.LargeChange = 1000
Me.HScroll1.SmallChange = 100
Me.VScroll1.LargeChange = 1000
Me.VScroll1.SmallChange = 100
Me.HScroll1.Enabled = False
Me.VScroll1.Enabled = False
Command2.Enabled = False
clear.Enabled = False
Label6.Enabled = False
Label8.Enabled = False
Label7.Enabled = False
Label11.Enabled = False
copy.Enabled = False
Me.Image1.Picture = LoadPicture()
Image1.Picture = LoadPicture("")
Image1.Enabled = False
Picture1.Enabled = False
Label6.Caption = ""
Label2.BackColor = vbWhite
Me.HScroll1.Enabled = False
Me.VScroll1.Enabled = False
Command2.Enabled = False
Label6.BackColor = vbWhite
Label7.Caption = ""
Label8.Caption = ""
Label6.Enabled = False
Label8.Enabled = False
Label7.Enabled = False
Label11.Enabled = False
copy.Enabled = False
Me.clear.Enabled = False
Label11.Caption = ""
Else
Me.VScroll1.Max = 0
Me.VScroll1.Min = 0
Me.HScroll1.Max = 0
Me.HScroll1.Min = 0
Me.VScroll1.Value = 0
Me.HScroll1.Value = 0
Me.HScroll1.LargeChange = 1000
Me.HScroll1.SmallChange = 100
Me.VScroll1.LargeChange = 1000
Me.VScroll1.SmallChange = 100
Me.HScroll1.Enabled = False
Me.VScroll1.Enabled = False
Command2.Enabled = False
clear.Enabled = False
Label6.Enabled = False
Label8.Enabled = False
Label7.Enabled = False
Label11.Enabled = False
copy.Enabled = False
Me.Image1.Picture = LoadPicture()
Image1.Picture = LoadPicture("")
Image1.Enabled = False
Picture1.Enabled = False
Label6.Caption = ""
Label2.BackColor = vbWhite
Me.HScroll1.Enabled = False
Me.VScroll1.Enabled = False
Command2.Enabled = False
Label6.BackColor = vbWhite
Label7.Caption = ""
Label8.Caption = ""
Label6.Enabled = False
Label8.Enabled = False
Label7.Enabled = False
Label11.Enabled = False
copy.Enabled = False
Me.clear.Enabled = False
Label11.Caption = ""
End If
End If
End Sub
Private Sub Command2_Click()
On Error Resume Next
Dim answer As Integer
answer = MsgBox("确定清除图像吗?", vbQuestion + vbYesNo, "ASK")
If answer = vbYes Then
Image1.Picture = LoadPicture("")
Image1.Enabled = False
Picture1.Enabled = False
Label6.Caption = ""
Label2.BackColor = vbWhite
Me.HScroll1.Enabled = False
Me.VScroll1.Enabled = False
Command2.Enabled = False
Label6.BackColor = vbWhite
Label7.Caption = ""
Label8.Caption = ""
Label11.Caption = ""
Label6.Enabled = False
Label8.Enabled = False
Label7.Enabled = False
Label11.Enabled = False
copy.Enabled = False
Me.clear.Enabled = False
Else
Exit Sub
End If
End Sub
Private Sub exit_Click()
On Error Resume Next
Dim ans As Integer
ans = MsgBox("确定要退出吗?", vbQuestion + vbYesNo, "Ask")
If ans = vbYes Then
End
Else
Exit Sub
End If
End Sub
Private Sub Form_Activate()
Command1.SetFocus
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
Unload Me
End If
End Sub
Private Sub Form_Load()
On Error Resume Next
If App.PrevInstance = False Then
Me.KeyPreview = True
Command2.Enabled = False
clear.Enabled = False
Dim forcunt As Integer
For forcunt = 0 To Me.Label1.UBound
Me.Label1(forcunt).AutoSize = True
Next
With Me.Label9
.Height = Me.Image2.Height
.BackColor = vbBlack
.Caption = vbCrLf & "PC_DOS" & vbCrLf & "Workshop" & vbCrLf & "Presents"
.FontBold = True
.FontItalic = False
.FontName = "System"
.FontStrikethru = False
.FontUnderline = False
.FontSize = 18
.Alignment = 2
End With
Me.Label3.AutoSize = True
With Me.VScroll1
.Height = Picture1.Height
.Width = 255
.Enabled = False
.Min = 0
.Max = 0
.Value = 0
End With
With Me.HScroll1
.Width = Picture1.Width
.Height = 255
.Enabled = False
.Min = 0
.Max = 0
.Value = 0
End With
With Me.Image1
.Left = 0
.Top = 0
.Picture = LoadPicture("")
.Width = 0
.Height = 0
End With
Label6.Enabled = False
Label8.Enabled = False
Label7.Enabled = False
Label11.Enabled = False
copy.Enabled = False
[screen].Enabled = True
[screen].Caption = "屏幕取色器(&S)"
info.Visible = False
Me.clear.Enabled = False
With Me
.Left = 0
.Top = 0
.Height = 8385
.Width = 12450
End With
Else
Form1.Hide
Command2.Enabled = False
clear.Enabled = False
For forcunt = 0 To Me.Label1.UBound
Me.Label1(forcunt).AutoSize = True
Next
Me.Label3.AutoSize = True
With Me.VScroll1
.Height = Picture1.Height
.Width = 255
.Enabled = False
.Min = 0
.Max = 0
.Value = 0
End With
With Me.HScroll1
.Width = Picture1.Width
.Height = 255
.Enabled = False
.Min = 0
.Max = 0
.Value = 0
End With
With Me.Image1
.Left = 0
.Top = 0
.Picture = LoadPicture("")
.Width = 0
.Height = 0
End With
With Me.Label9
.Height = Me.Image2.Height
.BackColor = vbBlack
.Caption = vbCrLf & "PC_DOS" & vbCrLf & "Workshop" & vbCrLf & "Presents"
.FontBold = True
.FontItalic = False
.FontName = "System"
.FontStrikethru = False
.FontUnderline = False
.FontSize = 18
.Alignment = 2
End With
Label6.Enabled = False
Label8.Enabled = False
Label7.Enabled = False
Label11.Enabled = False
copy.Enabled = False
[screen].Enabled = False
[screen].Caption = "功能已被稳定性管理器禁用"
info.Visible = True
Me.clear.Enabled = False
With Me
.Left = 0
.Top = 0
.Height = 8385
.Width = 12450
End With
Me.KeyPreview = True
Dim ans As Integer
ans = MsgBox("检测到应用程序已经有一个实例在运行,为了保证系统稳定性,部分功能将不可用" & vbCrLf & vbCrLf & "点击'确定'继续" & vbCrLf & "点击'取消'退出", vbExclamation + vbOKCancel, "Info")
If ans = vbOK Then
Form1.Show
Else
End
End If
End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Dim ans As Integer
ans = MsgBox("确定要退出吗?", vbQuestion + vbYesNo, "Ask")
If ans = vbYes Then
Cancel = 0
End
Else
Cancel = 666
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
End
End Sub
Private Sub chex_Click()
On Error Resume Next
If Trim(Label7.Caption) = "" Then Exit Sub
Clipboard.clear
Clipboard.SetText Label7.Caption
MsgBox "已经将颜色值复制到剪切板", vbExclamation, "Info"
End Sub
Private Sub HScroll1_Change()
On Error Resume Next
Image1.Left = -HScroll1.Value
If -HScroll1.Value > 0 Then
Image1.Left = HScroll1.Value
End If
End Sub
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next
Dim pi1&
Dim red As Integer
Dim green As Integer
Dim blue As Integer
Label4.Caption = "X:" & X - Me.HScroll1.Value
Label5.Caption = "Y:" & y - Me.VScroll1.Value
pi1& = Picture1.Point(X + Image1.Left, y + Image1.Top)
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
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next
Dim pi1&
Dim red As Integer
Dim green As Integer
Dim blue As Integer
Label4.Caption = "X:" & X - Me.HScroll1.Value
Label5.Caption = "Y:" & y - Me.VScroll1.Value
pi1& = Picture1.Point(X + Image1.Left, y + Image1.Top)
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
End Sub
Private Sub Image2_Click()
On Error Resume Next
frmAbout.Show 1
End Sub
Private Sub Label11_Click()
On Error Resume Next
If Trim(Label11.Caption) = "" Then Exit Sub
Clipboard.clear
Clipboard.SetText Label11.Caption
MsgBox "已经将颜色值复制到剪切板", vbExclamation, "Info"
End Sub
Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 2 Then
PopupMenu Me.copy
End If
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
Private Sub Label9_Click()
On Error Resume Next
frmAbout.Show 1
End Sub
Private Sub load_Click()
On Error GoTo ep
With CommonDialog1
.CancelError = True
.DialogTitle = "选择要装载的图片文件"
.Filter = "Pictures(*.jpg;*.jpeg;*.bmp;*.bip;*.gif;*.wmf)|*.jpg;*.jpeg;*.bmp;*.bip;*.gif;*.wmf"
.ShowModalWindow = True
.hWndCall = hWnd
.CancelError = True
Dim IsCanceled As Boolean
IsCanceled = .ShowOpen
End With
If IsCanceled = False Then
Exit Sub
End If
If Trim(CommonDialog1.FileName) <> "" Then
If Dir(CommonDialog1.FileName) <> "" Then
If Image1.Picture = IC.Picture Then
Image1.Picture = LoadPicture(CommonDialog1.FileName)
Picture1.Enabled = True
Image1.Enabled = True
Me.Image1.Left = 0
Me.Image1.Top = 0
Me.HScroll1.Enabled = True
Me.VScroll1.Enabled = True
If Me.Image1.Height <= Me.Picture1.Height And Me.Image1.Width <= Me.Picture1.Width Then
Me.Image1.Left = Me.Picture1.Width / 2 - Me.Image1.Width / 2
Me.Image1.Top = Me.Picture1.Height / 2 - Me.Image1.Height / 2
Me.VScroll1.Enabled = False
Me.HScroll1.Enabled = False
Command2.Enabled = True
clear.Enabled = True
Label6.Enabled = True
Label8.Enabled = True
Label7.Enabled = True
Label11.Enabled = True
copy.Enabled = True
Exit Sub
End If
If Me.Image1.Width <= Me.Picture1.Width Then
Me.HScroll1.Enabled = False
Me.VScroll1.Enabled = True
Me.Image1.Left = Me.Picture1.Width / 2 - Me.Image1.Width / 2
End If
If Me.Image1.Height <= Me.Picture1.Height Then
Me.VScroll1.Enabled = False
Me.HScroll1.Enabled = True
Me.Image1.Top = Me.Picture1.Height / 2 - Me.Image1.Height / 2
End If
Me.VScroll1.Max = Me.Image1.Height - Me.Picture1.Height
Me.HScroll1.Max = Me.Image1.Width - Me.Picture1.Width
Me.VScroll1.Value = 0
Me.HScroll1.Value = 0
Me.HScroll1.LargeChange = Image1.Width / 10
Me.HScroll1.SmallChange = Image1.Width / 50
Me.VScroll1.LargeChange = Image1.Width / 10
Me.VScroll1.SmallChange = Image1.Width / 50
Command2.Enabled = True
clear.Enabled = True
Label6.Enabled = True
Label8.Enabled = True
Label7.Enabled = True
Label11.Enabled = True
copy.Enabled = True
Else
Dim ans As Integer
ans = MsgBox("已经导入了图片,是否替换?", vbQuestion + vbYesNo, "Ask")
If ans = vbYes Then
Image1.Picture = LoadPicture(CommonDialog1.FileName)
Picture1.Enabled = True
Image1.Enabled = True
Me.Image1.Left = 0
Me.Image1.Top = 0
Me.HScroll1.Enabled = True
Me.VScroll1.Enabled = True
If Me.Image1.Height <= Me.Picture1.Height And Me.Image1.Width <= Me.Picture1.Width Then
Me.Image1.Left = Me.Picture1.Width / 2 - Me.Image1.Width / 2
Me.Image1.Top = Me.Picture1.Height / 2 - Me.Image1.Height / 2
Me.VScroll1.Enabled = False
Me.HScroll1.Enabled = False
Command2.Enabled = True
clear.Enabled = True
Label6.Enabled = True
Label8.Enabled = True
Label7.Enabled = True
Label11.Enabled = True
copy.Enabled = True
Exit Sub
End If
If Me.Image1.Width <= Me.Picture1.Width Then
Me.HScroll1.Enabled = False
Me.VScroll1.Enabled = True
Me.Image1.Left = Me.Picture1.Width / 2 - Me.Image1.Width / 2
End If
If Me.Image1.Height <= Me.Picture1.Height Then
Me.VScroll1.Enabled = False
Me.HScroll1.Enabled = True
Me.Image1.Top = Me.Picture1.Height / 2 - Me.Image1.Height / 2
End If
Me.VScroll1.Max = Me.Image1.Height - Me.Picture1.Height
Me.HScroll1.Max = Me.Image1.Width - Me.Picture1.Width
Me.VScroll1.Value = 0
Me.HScroll1.Value = 0
Me.HScroll1.LargeChange = Image1.Width / 10
Me.HScroll1.SmallChange = Image1.Width / 50
Me.VScroll1.LargeChange = Image1.Width / 10
Me.VScroll1.SmallChange = Image1.Width / 50
Command2.Enabled = True
clear.Enabled = True
Label6.Enabled = True
Label8.Enabled = True
Label7.Enabled = True
copy.Enabled = True
Else
Exit Sub
End If
End If
Else
MsgBox "找不到文件", vbCritical, "Error"
Exit Sub
End If
End If
Exit Sub
ep:
If Err.Number <> 32755 Then
MsgBox "发生错误:" & Chr(13) & Err.Description, vbCritical, "Error"
If Err.Description = "溢出" Or Err.Number = 6 Then
Me.VScroll1.Max = 32755
Me.VScroll1.Min = 0
Me.HScroll1.Max = 32755
Me.HScroll1.Min = 0
Me.VScroll1.Value = 0
Me.HScroll1.Value = 0
Me.HScroll1.LargeChange = 1000
Me.HScroll1.SmallChange = 100
Me.VScroll1.LargeChange = 1000
Me.VScroll1.SmallChange = 100
Command2.Enabled = True
clear.Enabled = True
Label6.Enabled = True
Label8.Enabled = True
Label7.Enabled = True
Label11.Enabled = True
copy.Enabled = True
Else
Me.VScroll1.Max = 0
Me.VScroll1.Min = 0
Me.HScroll1.Max = 0
Me.HScroll1.Min = 0
Me.VScroll1.Value = 0
Me.HScroll1.Value = 0
Me.HScroll1.LargeChange = 1000
Me.HScroll1.SmallChange = 100
Me.VScroll1.LargeChange = 1000
Me.VScroll1.SmallChange = 100
Me.HScroll1.Enabled = False
Me.VScroll1.Enabled = False
Command2.Enabled = False
clear.Enabled = False
Label6.Enabled = False
Label8.Enabled = False
Label7.Enabled = False
Label11.Enabled = False
copy.Enabled = False
Me.Image1.Picture = LoadPicture()
Image1.Picture = LoadPicture("")
Image1.Enabled = False
Picture1.Enabled = False
Label6.Caption = ""
Label2.BackColor = vbWhite
Me.HScroll1.Enabled = False
Me.VScroll1.Enabled = False
Command2.Enabled = False
Label6.BackColor = vbWhite
Label7.Caption = ""
Label8.Caption = ""
Label6.Enabled = False
Label8.Enabled = False
Label7.Enabled = False
Label11.Enabled = False
copy.Enabled = False
Me.clear.Enabled = False
Label11.Caption = ""
End If
Else
Exit Sub
End If
If Err.Number = 481 Or Err.Number = 53 Then
If Image1.Picture = LoadPicture() Then
Me.VScroll1.Max = 0
Me.VScroll1.Min = 0
Me.HScroll1.Max = 0
Me.HScroll1.Min = 0
Me.VScroll1.Value = 0
Me.HScroll1.Value = 0
Me.HScroll1.LargeChange = 1000
Me.HScroll1.SmallChange = 100
Me.VScroll1.LargeChange = 1000
Me.VScroll1.SmallChange = 100
Me.HScroll1.Enabled = False
Me.VScroll1.Enabled = False
Command2.Enabled = False
clear.Enabled = False
Label6.Enabled = False
Label8.Enabled = False
Label7.Enabled = False
Label11.Enabled = False
copy.Enabled = False
Me.Image1.Picture = LoadPicture()
Image1.Picture = LoadPicture("")
Image1.Enabled = False
Picture1.Enabled = False
Label6.Caption = ""
Label2.BackColor = vbWhite
Me.HScroll1.Enabled = False
Me.VScroll1.Enabled = False
Command2.Enabled = False
Label6.BackColor = vbWhite
Label7.Caption = ""
Label8.Caption = ""
Label6.Enabled = False
Label8.Enabled = False
Label7.Enabled = False
Label11.Enabled = False
copy.Enabled = False
Me.clear.Enabled = False
Label11.Caption = ""
Else
Me.VScroll1.Max = 0
Me.VScroll1.Min = 0
Me.HScroll1.Max = 0
Me.HScroll1.Min = 0
Me.VScroll1.Value = 0
Me.HScroll1.Value = 0
Me.HScroll1.LargeChange = 1000
Me.HScroll1.SmallChange = 100
Me.VScroll1.LargeChange = 1000
Me.VScroll1.SmallChange = 100
Me.HScroll1.Enabled = False
Me.VScroll1.Enabled = False
Command2.Enabled = False
clear.Enabled = False
Label6.Enabled = False
Label8.Enabled = False
Label7.Enabled = False
Label11.Enabled = False
copy.Enabled = False
Me.Image1.Picture = LoadPicture()
Image1.Picture = LoadPicture("")
Image1.Enabled = False
Picture1.Enabled = False
Label6.Caption = ""
Label2.BackColor = vbWhite
Me.HScroll1.Enabled = False
Me.VScroll1.Enabled = False
Command2.Enabled = False
Label6.BackColor = vbWhite
Label7.Caption = ""
Label8.Caption = ""
Label6.Enabled = False
Label8.Enabled = False
Label7.Enabled = False
Label11.Enabled = False
copy.Enabled = False
Me.clear.Enabled = False
Label11.Caption = ""
End If
End If
End Sub
Private Sub PS_Click()
On Error Resume Next
If Trim(Label11.Caption) = "" Then Exit Sub
Clipboard.clear
Clipboard.SetText Label11.Caption
MsgBox "已经将颜色值复制到剪切板", vbExclamation, "Info"
End Sub
Private Sub RGB_Click()
Form2.Show
End Sub
Private Sub RGBC_Click()
On Error Resume Next
If Trim(Label6.Caption) = "" Then Exit Sub
Clipboard.clear
Clipboard.SetText Label6.Caption
MsgBox "已经将颜色值复制到剪切板", vbExclamation, "Info"
End Sub
Private Sub screen_Click()
Me.Visible = False
Me.Hide
Form2.Hide
Form3.Show
End Sub
Private Sub VScroll1_Change()
On Error Resume Next
Image1.Top = -VScroll1.Value
If -VScroll1.Value > 0 Then
Image1.Top = VScroll1.Value
End If
End Sub
Private Sub web_Click()
On Error Resume Next
If Trim(Label8.Caption) = "" Then Exit Sub
Clipboard.clear
Clipboard.SetText Label8.Caption
MsgBox "已经将颜色值复制到剪切板", vbExclamation, "Info"
End Sub
