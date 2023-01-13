VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form1 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "关于 Locker"
   ClientHeight    =   7950
   ClientLeft      =   1050
   ClientTop       =   1785
   ClientWidth     =   6570
   BeginProperty Font 
      Name            =   "OPPOSans R"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   6570
   Begin VB.CommandButton Command1 
      Caption         =   "展开公告栏"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   7250
      Width           =   1095
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   7335
      Left            =   7080
      TabIndex        =   2
      Top             =   120
      Width           =   7455
      ExtentX         =   13150
      ExtentY         =   12938
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
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "宪逸"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   13
      ToolTipText     =   "Phosoft创始人之一，他提出了许多细节问题，让Locker更加完美！"
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "What_Damon"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   12
      ToolTipText     =   "我不是很好形容他，但是我知道他很擅长软件UI的设计，也在很多方面帮助过Locker"
      Top             =   3490
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "张六六Zpcin"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   11
      ToolTipText     =   "Phosoft创始人之一，他提出了许多细节问题，让Locker更加完美！"
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "版本 10.0.6 Gen2"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   10
      Top             =   1850
      Width           =   3495
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Locker"
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
      Left            =   720
      TabIndex        =   9
      Top             =   6200
      Width           =   2295
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Locker 的诞生离不开 Internet Explorer 和 NetExplore"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   8
      ToolTipText     =   $"About.frx":1084A
      Top             =   5600
      Width           =   5175
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Redmountain2018"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      ToolTipText     =   "AuroraStudio创始人，多次帮助Locker的UI和功能改进，他让Locker更完美！"
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "苏方华_RTC"
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
      Left            =   1200
      TabIndex        =   6
      ToolTipText     =   "INTRON软件组织的创始人，NetExplore的开发者，Locker的诞生离不开他的帮助！"
      Top             =   3050
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "鸣谢："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   5
      ToolTipText     =   "向列表中的人表示深深的敬意"
      Top             =   2480
      Width           =   2175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000011&
      X1              =   240
      X2              =   6120
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Internet 更新公告:"
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
      Left            =   720
      TabIndex        =   4
      Top             =   7320
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   1350
      Left            =   600
      Picture         =   "About.frx":108D2
      ToolTipText     =   "告诉你个秘密：LockerX很环保，因为他没附带充电器"
      Top             =   120
      Width           =   1350
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "此软件由狗_君开发。保留所有权利。"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Top             =   6650
      Width           =   7215
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Pilot Locker"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   1695
      Left            =   2280
      TabIndex        =   0
      ToolTipText     =   "Pilot Locker是 The Locker Project 的第一个软件项目 "
      Top             =   480
      Width           =   5055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Sub Command1_Click()
If Form1.Width = 6660 Then
Form1.Width = 14685
Command1.Caption = "收起公告栏"
Else
Form1.Width = 6660
Command1.Caption = "展开公告栏"
End If
End Sub
Private Sub Form_Load()
  HomeAddress = "http://windows.3vhost.net/locker/lockerx.htm " '填写主页地址
    WebBrowser1.Navigate HomeAddress
    myval = SetWindowPos(Form1.hwnd, -1, 0, 0, 0, 0, 3)
End Sub

