VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Locker 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Locker"
   ClientHeight    =   4515
   ClientLeft      =   105
   ClientTop       =   765
   ClientWidth     =   9360
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
   ScaleHeight     =   4515
   ScaleWidth      =   9360
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "访问"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   17
      Top             =   800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6135
      Left            =   600
      ScaleHeight     =   6105
      ScaleWidth      =   12345
      TabIndex        =   10
      Top             =   650
      Visible         =   0   'False
      Width           =   12375
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "尝试重启无线路由器、宽带或信号交换机"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1080
         TabIndex        =   21
         Top             =   5040
         Width           =   3975
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "确保网线已插好"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1080
         TabIndex        =   20
         Top             =   4560
         Width           =   3735
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "确保已关闭飞行模式"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1080
         TabIndex        =   19
         Top             =   4080
         Width           =   3735
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "查看网络共享中心"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   495
         Left            =   1080
         TabIndex        =   18
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "查看网络连接面板"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   495
         Left            =   1080
         TabIndex        =   16
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "确保Locker没有被杀毒软件或防火墙拦截"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   15
         Top             =   2640
         Width           =   3855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "点此重新加载页面（请勿点击刷新按钮）"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   495
         Left            =   1080
         TabIndex        =   14
         Top             =   2160
         Width           =   4455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "检查路由器是否在正常工作且信号稳定"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1080
         TabIndex        =   13
         Top             =   1680
         Width           =   3735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "您可以："
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   12
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "网络异常"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   840
         TabIndex        =   11
         Top             =   480
         Width           =   3615
      End
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   450
      Left            =   13440
      TabIndex        =   8
      Text            =   " 搜索..."
      Top             =   60
      Width           =   4800
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   120
      Picture         =   "Locker.frx":1084A
      ScaleHeight     =   585
      ScaleWidth      =   405
      TabIndex        =   7
      Top             =   6480
      Width           =   405
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   410
      Left            =   120
      Picture         =   "Locker.frx":10D66
      ScaleHeight     =   405
      ScaleWidth      =   375
      TabIndex        =   6
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
      Picture         =   "Locker.frx":112CD
      ScaleHeight     =   375
      ScaleWidth      =   495
      TabIndex        =   5
      Top             =   70
      Width           =   495
   End
   Begin VB.PictureBox REFRESH 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   2040
      Picture         =   "Locker.frx":1161C
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   4
      Top             =   120
      Width           =   360
   End
   Begin VB.PictureBox FORWARD 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   1440
      Picture         =   "Locker.frx":11B26
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   3
      Top             =   120
      Width           =   360
   End
   Begin VB.PictureBox BACK 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   840
      Picture         =   "Locker.frx":11FC2
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   2
      Top             =   120
      Width           =   360
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1560
      Visible         =   0   'False
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
      ForeColor       =   &H00808080&
      Height          =   450
      Left            =   2640
      TabIndex        =   0
      Top             =   60
      Width           =   10335
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6735
      Left            =   600
      TabIndex        =   9
      Top             =   630
      Width           =   13455
      ExtentX         =   23733
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
   Begin VB.Line Line4 
      BorderColor     =   &H8000000C&
      X1              =   598
      X2              =   598
      Y1              =   0
      Y2              =   108
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
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000000&
      X1              =   590
      X2              =   590
      Y1              =   0
      Y2              =   1080
   End
   Begin VB.Image Image3 
      Height          =   585
      Left            =   0
      Picture         =   "Locker.frx":12468
      Top             =   840
      Width           =   585
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   0
      Picture         =   "Locker.frx":12973
      Top             =   120
      Width           =   585
   End
   Begin VB.Menu Control 
      Caption         =   "az"
      Visible         =   0   'False
      Begin VB.Menu print 
         Caption         =   "打印"
      End
      Begin VB.Menu file 
         Caption         =   "文件"
         Begin VB.Menu NEW 
            Caption         =   "打开"
         End
         Begin VB.Menu save 
            Caption         =   "另存为"
         End
         Begin VB.Menu t 
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
         Begin VB.Menu stop 
            Caption         =   "停止"
         End
         Begin VB.Menu gohome 
            Caption         =   "主页"
         End
      End
      Begin VB.Menu view 
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
            Caption         =   "专注模式"
         End
         Begin VB.Menu ballonscreen 
            Caption         =   "Locker栏"
         End
      End
      Begin VB.Menu az 
         Caption         =   "-"
      End
      Begin VB.Menu function 
         Caption         =   "工具"
         Begin VB.Menu Command2 
            Caption         =   "查看网页源代码"
         End
      End
      Begin VB.Menu M 
         Caption         =   "-"
      End
      Begin VB.Menu Command3 
         Caption         =   "设置"
      End
      Begin VB.Menu update 
         Caption         =   "查看 Locker 更新"
      End
      Begin VB.Menu ab 
         Caption         =   "关于 Locker"
      End
      Begin VB.Menu dd 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu exit 
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
         Begin VB.Menu stoppppppp 
            Caption         =   "停止"
         End
         Begin VB.Menu thexxhomepage 
            Caption         =   "主页"
         End
      End
      Begin VB.Menu saveastop 
         Caption         =   "另存为"
      End
      Begin VB.Menu printtop 
         Caption         =   "打印"
      End
      Begin VB.Menu findtop 
         Caption         =   "查找"
      End
      Begin VB.Menu junkline 
         Caption         =   "-"
      End
      Begin VB.Menu exittop 
         Caption         =   "退出"
      End
   End
   Begin VB.Menu viewtop 
      Caption         =   "查看"
      Begin VB.Menu zoom150top 
         Caption         =   "缩放150%"
      End
      Begin VB.Menu zoom125top 
         Caption         =   "缩放125%"
      End
      Begin VB.Menu zoom100top 
         Caption         =   "缩放100%"
      End
      Begin VB.Menu zoom75top 
         Caption         =   "缩放75%"
      End
      Begin VB.Menu zoom50top 
         Caption         =   "缩放50%"
      End
      Begin VB.Menu zoom25top 
         Caption         =   "缩放25%"
      End
   End
   Begin VB.Menu optiontop 
      Caption         =   "选项"
      Begin VB.Menu modetop 
         Caption         =   "模式"
         Begin VB.Menu smalltop 
            Caption         =   "小窗模式"
         End
         Begin VB.Menu Mordentop 
            Caption         =   "专注模式"
         End
         Begin VB.Menu lockerbartop 
            Caption         =   "Locker栏"
         End
      End
      Begin VB.Menu tool 
         Caption         =   "工具"
         Begin VB.Menu webcodetop 
            Caption         =   "查看网页源代码                 "
         End
      End
   End
   Begin VB.Menu help 
      Caption         =   "帮助"
      Begin VB.Menu Commandyee 
         Caption         =   "设置"
      End
      Begin VB.Menu updatetop 
         Caption         =   "查看 Locker 更新"
      End
      Begin VB.Menu abouttop 
         Caption         =   "关于 Locker"
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
Private Sub abouttop_Click()
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
Private Sub ballonscreen_Click()
LockerBar.Show
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
WebBrowser1.Navigate Trim(UTF8EncodeURI(Text1.Text)) '打开网页
Else
MsgBox "Locker X 10.0.6 Generations"
End If
End Sub
Private Sub About_Click()
Form1.Show
End Sub
Private Sub Command2_Click()
On Error Resume Next
Form2.Text1.Locked = False
Form2.Text1.Text = Locker.WebBrowser1.Document.body.innerHtml
Form2.Show
Form2.Text1.Locked = True
End Sub
Private Sub Command3_Click()
Setting.Show
End Sub
Private Sub Command4_Click()
ab.Visible = True
dd.Visible = True
PopupMenu Control, vbPopupMenuLeftAlign
ab.Visible = False
dd.Visible = False
End Sub
Private Sub Commandyee_Click()
Setting.Show
End Sub
Private Sub exittop_Click()
End
End Sub
Private Sub fffffffffffffforwardddddd_Click()
On Error Resume Next
WebBrowser1.GoForward
End Sub
Private Sub findtop_Click()
On Error Resume Next
WebBrowser1.ExecWB OLECMDID_FIND, OLECMDEXECOPT_DODEFAULT
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
WebBrowser1.gohome
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
Private Sub Label4_Click()
WebBrowser1.Navigate (Text1.Text)
End Sub
Private Sub Label6_Click()
Shell "explorer.exe shell:::{7007ACC7-3202-11D1-AAD2-00805FC1270E}", 1
End Sub
Private Sub Label7_Click()
Shell "explorer.exe shell:::{8E908FC9-BECC-40f6-915B-F4CA0E70D03D}", 1
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
Private Sub lockerbartop_Click()
LockerBar.Show
End Sub
Private Sub Morden_Click()
Warn.Text1.Text = "chenjin"
Warn.Show
End Sub
Private Sub Mordentop_Click()
Warn.Text1.Text = "chenjin"
Warn.Show
End Sub
Private Sub New_Click()
WebBrowser1.ExecWB OLECMDID_OPEN, OLECMDEXECOPT_DODEFAULT
End Sub
Private Sub op_Click()
New_Click
End Sub
Private Sub Picture1_Click()
On Error Resume Next
 WebBrowser1.gohome '返回主页
End Sub
Private Sub Picture2_Click()
Form2.Show
Form2.Text1.Locked = False
    Dim doc As Object
    Dim i As Object
    Dim strHtml As String
    Set doc = WebBrowser1.Document
    For Each i In doc.All
        strHtml = strHtml & Chr(13) & i.innerHtml
    Next
Form2.Text1.Text = strHtml
Form2.Text1.Locked = True
End Sub

Private Sub Picture3_Click()

End Sub

Private Sub Print_Click()
WebBrowser1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT
End Sub
Private Sub printtop_Click()
WebBrowser1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT
End Sub
Private Sub reeefffressssssssshhhh_Click()
On Error Resume Next
WebBrowser1.REFRESH
End Sub
Private Sub Save_Click()
WebBrowser1.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DODEFAULT
End Sub
Private Sub saveastop_Click()
WebBrowser1.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DODEFAULT
End Sub
Private Sub smalltop_Click()
Warn.Text1.Text = "small"
Warn.Show
End Sub
Private Sub stoppppppp_Click()
On Error Resume Next
WebBrowser1.stop
End Sub
Private Sub T_Click()
WebBrowser1.ExecWB OLECMDID_FIND, OLECMDEXECOPT_DODEFAULT
End Sub
Private Sub FORWARD_Click() '前进
On Error Resume Next
    WebBrowser1.GoForward
    GoAdDress.Text = WebBrowser1.LocationURL
End Sub
Private Sub Little_Click()
On Error Resume Next
Warn.Text1.Text = "small"
Warn.Show
End Sub
Private Sub REFRESH_Click() '刷新
On Error Resume Next
    WebBrowser1.REFRESH
   End Sub
Private Sub thexxhomepage_Click()
On Error Resume Next
WebBrowser1.gohome
End Sub
Private Sub Timer1_Timer()
Me.Left = -10
Me.Top = -10
Timer1.Enabled = False
Timer2.Enabled = True
End Sub
Private Sub Timer2_Timer()
Me.Width = 12848
Me.Height = 7227
Timer2.Enabled = False
Timer3.Enabled = True
End Sub
Private Sub Timer3_Timer()
Me.WindowState = 2
Timer3.Enabled = False
End Sub
Private Sub UPDATE_Click()
Form1.Show
Form1.Width = 15570
Form1.Command1.Caption = "收起公告栏"
End Sub
Private Sub updatetop_Click()
Form1.Show
Form1.Width = 15570
End Sub
Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
On Error Resume Next
Cancel = True
WebBrowser1.Navigate2 WebBrowser1.Document.activeElement.href
End Sub
Private Sub STOP_Click()
On Error Resume Next
WebBrowser1.stop
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    WebBrowser1.Height = Locker.Height - 1000
    WebBrowser1.Width = Locker.Width - 840
    Text1.Width = Me.Width - 8800
    Text3.Left = Text1.Left + Text1.Width + 100
    Text3.Width = Locker.Width - Text1.Left - Text1.Width - 100 - 860
    Command4.Left = Text3.Left + Text3.Width + 120
    Line1.Y2 = Me.Height
    Line2.X2 = WebBrowser1.Left + WebBrowser1.Width
    Line3.X2 = WebBrowser1.Left + WebBrowser1.Width
    Picture1.Top = WebBrowser1.Height - 500
    Picture2.Top = Picture1.Top - 720
    Line4.Y2 = Me.Height
    Picture3.Left = WebBrowser1.Left
    Picture3.Top = WebBrowser1.Top
    Picture3.Width = WebBrowser1.Width
    Picture3.Height = WebBrowser1.Height
    End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer) '回车键，需要改按钮
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        Command1_Click
        Text1.Text = Text1.Text
    End If
    If Text1.ForeColor = &H808080 Then
Text1.ForeColor = &H0&
End If
End Sub
Private Sub WebBrowser1_StatuTextChange()
On Error Resume Next
End Sub
Private Sub WebBrowser1_TitleChange(ByVal Text As String)
Text1.Text = WebBrowser1.LocationURL
If WebBrowser1.Busy = True Then
Text1.ForeColor = &H808080
Else
Text1.ForeColor = &H0&
End If
    If InternetGetConnectedState(0&, 0&) Then
       Picture3.Visible = False
    Else
        Picture3.Visible = True
    End If
Me.Caption = WebBrowser1.LocationName + " - Locker"
WebBrowser1.Silent = True
End Sub
Private Sub Me_Unload(Cancel As Integer)
End
End Sub
Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
'如果单击搜索框，让搜索文字消失
If Text1.ForeColor = &H808080 Then
Text1.ForeColor = &H0&
End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
'如果按下回车键，则访问网页
If KeyAscii = vbKeyReturn Then
    WebBrowser1.Navigate "https://cn.bing.com/search?q=" + (UTF8EncodeURI(Text3.Text))
End If
End Sub
Private Sub Text3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
'如果单击搜索框，让搜索文字消失
If Text3.Text = Text3.Text Then
Text3.ForeColor = &H0&
Text3.Text = ""
End If
End Sub
Function UTF8EncodeURI(szInput)                 '转UTF8码声明
Dim wch, uch, szRet
Dim x
Dim nAsc, nAsc2, nAsc3
If szInput = "" Then
UTF8EncodeURI = szInput
Exit Function
End If
For x = 1 To Len(szInput)
wch = Mid(szInput, x, 1)
nAsc = AscW(wch)
If nAsc < 0 Then nAsc = nAsc + 65536
If (nAsc And &HFF80) = 0 Then
szRet = szRet & wch
Else
If (nAsc And &HF000) = 0 Then
uch = "%" & Hex(((nAsc \ 2 ^ 6)) Or &HC0) & Hex(nAsc And &H3F Or &H80)
szRet = szRet & uch
Else
uch = "%" & Hex((nAsc \ 2 ^ 12) Or &HE0) & "%" & _
Hex((nAsc \ 2 ^ 6) And &H3F Or &H80) & "%" & _
Hex(nAsc And &H3F Or &H80)
szRet = szRet & uch
End If
End If
Next
UTF8EncodeURI = szRet
End Function
Function GBKEncodeURI(szInput)          '声明GBK15编码
Dim i As Long
Dim x() As Byte
Dim szRet As String
szRet = ""
x = StrConv(szInput, vbFromUnicode)
For i = LBound(x) To UBound(x)
szRet = szRet & "%" & Hex(x(i))
Next
GBKEncodeURI = szRet
End Function
Private Sub webcodetop_Click()
Form2.Show
Form2.Text1.Locked = False
    Dim doc As Object
    Dim i As Object
    Dim strHtml As String
    Set doc = WebBrowser1.Document
    For Each i In doc.All
        strHtml = strHtml & Chr(13) & i.innerHtml
    Next
Form2.Text1.Text = strHtml
Form2.Text1.Locked = True
End Sub
Private Sub zoom100top_Click()
WebBrowser1.Document.body.Style.Zoom = "100%"
End Sub
Private Sub zoom125top_Click()
WebBrowser1.Document.body.Style.Zoom = "125%"
End Sub
Private Sub zoom150top_Click()
WebBrowser1.Document.body.Style.Zoom = "150%"
End Sub
Private Sub zoom25top_Click()
WebBrowser1.Document.body.Style.Zoom = "25%"
End Sub
Private Sub zoom50top_Click()
WebBrowser1.Document.body.Style.Zoom = "50%"
End Sub
Private Sub zoom75top_Click()
WebBrowser1.Document.body.Style.Zoom = "75%"
End Sub
