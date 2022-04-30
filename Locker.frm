VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Locker 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Locker"
   ClientHeight    =   7755
   ClientLeft      =   120
   ClientTop       =   1020
   ClientWidth     =   14595
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Locker.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   14595
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   410
      Left            =   120
      Picture         =   "Locker.frx":1084A
      ScaleHeight     =   405
      ScaleWidth      =   375
      TabIndex        =   8
      Top             =   7200
      Width           =   375
   End
   Begin VB.PictureBox Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   17880
      Picture         =   "Locker.frx":10C68
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   7
      Top             =   70
      Width           =   495
   End
   Begin VB.PictureBox REFRESH 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   2180
      Picture         =   "Locker.frx":10F73
      ScaleHeight     =   345
      ScaleWidth      =   510
      TabIndex        =   6
      Top             =   100
      Width           =   510
   End
   Begin VB.PictureBox FORWARD 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   1560
      Picture         =   "Locker.frx":1139D
      ScaleHeight     =   450
      ScaleWidth      =   345
      TabIndex        =   5
      Top             =   50
      Width           =   345
   End
   Begin VB.PictureBox BACK 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   860
      Picture         =   "Locker.frx":1174C
      ScaleHeight     =   360
      ScaleWidth      =   345
      TabIndex        =   4
      Top             =   90
      Width           =   345
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6735
      Left            =   600
      TabIndex        =   3
      Top             =   585
      Width           =   13935
      ExtentX         =   24580
      ExtentY         =   11880
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2900
      TabIndex        =   1
      Top             =   60
      Width           =   10455
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "sdf"
      CausesValidation=   0   'False
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   800
      Width           =   975
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000000&
      X1              =   0
      X2              =   2040
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000000&
      X1              =   600
      X2              =   2040
      Y1              =   565
      Y2              =   565
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000000&
      X1              =   580
      X2              =   580
      Y1              =   0
      Y2              =   1080
   End
   Begin VB.Image Image3 
      Height          =   585
      Left            =   0
      Picture         =   "Locker.frx":11AF6
      Top             =   840
      Width           =   585
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   0
      Picture         =   "Locker.frx":12001
      Top             =   120
      Width           =   585
   End
   Begin VB.Menu Control 
      Caption         =   "asdf"
      Visible         =   0   'False
      Begin VB.Menu Print 
         Caption         =   "打印"
      End
      Begin VB.Menu file 
         Caption         =   "文件"
         Begin VB.Menu NEW 
            Caption         =   "打开"
         End
         Begin VB.Menu Save 
            Caption         =   "另存为"
         End
         Begin VB.Menu T 
            Caption         =   "查找"
         End
      End
      Begin VB.Menu Do 
         Caption         =   "操作"
         Begin VB.Menu Bac 
            Caption         =   "后退"
         End
         Begin VB.Menu FORWAR 
            Caption         =   "前进"
         End
         Begin VB.Menu FRESH 
            Caption         =   "刷新"
         End
         Begin VB.Menu STOP 
            Caption         =   "停止"
         End
         Begin VB.Menu GoHome 
            Caption         =   "主页"
         End
      End
      Begin VB.Menu View 
         Caption         =   "缩放"
         Begin VB.Menu big150 
            Caption         =   "缩放150%"
         End
         Begin VB.Menu Big125 
            Caption         =   "缩放125%"
         End
         Begin VB.Menu Big100 
            Caption         =   "缩放100%"
         End
         Begin VB.Menu little75 
            Caption         =   "缩放75%"
         End
         Begin VB.Menu Little50 
            Caption         =   "缩放50%"
         End
         Begin VB.Menu Little25 
            Caption         =   "缩放25%"
         End
      End
      Begin VB.Menu wnm 
         Caption         =   "-"
      End
      Begin VB.Menu Choose 
         Caption         =   "模式"
         Begin VB.Menu Little 
            Caption         =   "小窗模式"
         End
         Begin VB.Menu Morden 
            Caption         =   "MordenUI模式"
         End
      End
      Begin VB.Menu az 
         Caption         =   "-"
      End
      Begin VB.Menu function 
         Caption         =   "工具"
         Begin VB.Menu showphoto 
            Caption         =   "展示图片(非PNG格式）"
         End
         Begin VB.Menu Command2 
            Caption         =   "查看网页源代码"
         End
         Begin VB.Menu UPDATE 
            Caption         =   "查看更新公告"
         End
      End
      Begin VB.Menu M 
         Caption         =   "-"
      End
      Begin VB.Menu ab 
         Caption         =   "关于 Locker"
         Visible         =   0   'False
      End
      Begin VB.Menu dd 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu Exit 
         Caption         =   "退出"
      End
   End
   Begin VB.Menu flie 
      Caption         =   "网页"
      Begin VB.Menu op 
         Caption         =   "打开"
      End
      Begin VB.Menu oppppperate 
         Caption         =   "操作"
         Begin VB.Menu bbbbaack 
            Caption         =   "后退"
         End
         Begin VB.Menu fffffffffffffforwardddddd 
            Caption         =   "前进"
         End
         Begin VB.Menu reeefffressssssssshhhh 
            Caption         =   "刷新"
         End
      End
   End
End
Attribute VB_Name = "Locker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
Private Sub ab_Click()
Form1.Show
End Sub
Private Sub Bac_Click()
On Error Resume Next
WebBrowser1.GoBack
End Sub
Private Sub BACK_Click() '返回
    On Error Resume Next
    WebBrowser1.GoBack
    GoAdDress.Text = WebBrowser1.LocationURL
End Sub

Private Sub bbbbaack_Click()
On Error Resume Next
WebBrowser1.GoBack
End Sub

Private Sub Big100_Click()
On Error Resume Next
WebBrowser1.Document.body.Style.Zoom = "100%"
End Sub
Private Sub Big125_Click()
On Error Resume Next
WebBrowser1.Document.body.Style.Zoom = "125%"
End Sub
Private Sub Big150_Click()
On Error Resume Next
WebBrowser1.Document.body.Style.Zoom = "150%"
End Sub
Private Sub Command1_Click()
On Error Resume Next
If Not Text1.Text = "about:" Then
WebBrowser1.Navigate Trim(Text1.Text) '打开网页
Else
MsgBox "版本8.1"
End If
End Sub
Private Sub About_Click()
Form1.Show
End Sub
Private Sub Command2_Click()
On Error Resume Next
Form2.Text1.Text = Locker.WebBrowser1.Document.body.innerHTML
Form2.Show
Form2.Combo1.Text = Locker.WebBrowser1.LocationURL
End Sub
Private Sub Command4_Click()
ab.Visible = True
dd.Visible = True
PopupMenu Control, vbPopupMenuLeftAlign
ab.Visible = False
dd.Visible = False
End Sub

Private Sub fffffffffffffforwardddddd_Click()
On Error Resume Next
WebBrowser1.GoForward
End Sub

Private Sub FORWAR_Click()
On Error Resume Next
WebBrowser1.GoForward
End Sub
Private Sub FRESH_Click()
On Error Resume Next
 WebBrowser1.REFRESH
End Sub
Private Sub GoHome_Click()
On Error Resume Next
WebBrowser1.GoHome
End Sub
Private Sub EXIT_Click()
End
End Sub
Private Sub Image1_Click()
On Error Resume Next
WebBrowser1.Navigate "https://cn.bing.com"
End Sub
Private Sub Image2_Click()
On Error Resume Next
WebBrowser1.Navigate "https://fanyi.baidu.com/"
End Sub
Private Sub Image3_Click()
On Error Resume Next
WebBrowser1.Navigate "https://baike.baidu.com"
End Sub
Private Sub Little25_Click()
On Error Resume Next
WebBrowser1.Document.body.Style.Zoom = "25%"
End Sub
Private Sub Little50_Click()
On Error Resume Next
WebBrowser1.Document.body.Style.Zoom = "50%"
End Sub
Private Sub little75_Click()
On Error Resume Next
WebBrowser1.Document.body.Style.Zoom = "75%"
End Sub
Private Sub Morden_Click()
zhuan.Show
End Sub
Private Sub New_Click()
WebBrowser1.ExecWB OLECMDID_OPEN, OLECMDEXECOPT_DODEFAULT
End Sub
Private Sub op_Click()
New_Click
End Sub
Private Sub Picture1_Click()
On Error Resume Next
 WebBrowser1.GoHome '返回主页
End Sub
Private Sub Print_Click()
WebBrowser1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT
End Sub

Private Sub reeefffressssssssshhhh_Click()
On Error Resume Next
WebBrowser1.REFRESH
End Sub

Private Sub Save_Click()
WebBrowser1.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DODEFAULT
End Sub
Private Sub showphoto_Click()
Form5.Show
Me.Hide
End Sub
Private Sub T_Click()
WebBrowser1.ExecWB OLECMDID_FIND, OLECMDEXECOPT_DODEFAULT
End Sub
Private Sub Form_Load()
    Call Command1_Click
  HomeAddress = " https://cn.bing.com" '填写主页地址
    WebBrowser1.Navigate HomeAddress
Text1.Width = 11175
Set W = CreateObject("wscript.shell")
W.regwrite "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION\" & App.EXEName + ".exe", "11000", "REG_DWORD"
Set W = Nothing
End Sub
Private Sub FORWARD_Click() '前进
On Error Resume Next
    WebBrowser1.GoForward
    GoAdDress.Text = WebBrowser1.LocationURL
End Sub
Private Sub Little_Click()
On Error Resume Next
Form3.Show
Locker.Hide
End Sub
Private Sub REFRESH_Click() '刷新
On Error Resume Next
    WebBrowser1.REFRESH
   End Sub
Private Sub UPDATE_Click()
Form1.Show
Form1.Width = 17055
Form1.Command1.Caption = "收起公告栏"
End Sub
Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
On Error Resume Next
Cancel = True
WebBrowser1.Navigate2 WebBrowser1.Document.activeElement.href
End Sub
Private Sub STOP_Click()
On Error Resume Next
WebBrowser1.STOP
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    WebBrowser1.Height = Locker.Height - 1000
    WebBrowser1.Width = Locker.Width - 825
    Text1.Width = Me.Width - 4100
    Command4.Left = Text1.Left + Text1.Width + 120
    Line1.Y2 = Me.Height
    Line2.X2 = WebBrowser1.Left + WebBrowser1.Width
    Line3.X2 = WebBrowser1.Left + WebBrowser1.Width
    Picture1.Top = WebBrowser1.Height - 500

    End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer) '回车键，需要改按钮
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        Command1_Click
        Text1.Text = Text1.Text
    End If
End Sub
Private Sub WebBrowser1_StatuTextChange()
On Error Resume Next
End Sub
Private Sub WebBrowser1_TitleChange(ByVal Text As String)
Text1.Text = WebBrowser1.LocationURL
    If InternetGetConnectedState(0&, 0&) Then
    Else
        Form9.Show
    End If
End Sub
Private Sub Text1_Click()
On Error Resume Next
Text1.Text = ""
End Sub
Private Sub Me_Unload(Cancel As Integer)
End
End Sub
