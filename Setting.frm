VERSION 5.00
Begin VB.Form Setting 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6480
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command5 
      Caption         =   "设置为搜狐"
      Height          =   375
      Left            =   5040
      TabIndex        =   7
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "设置为搜狗"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "设置为百度"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "设置为必应"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "保存设置"
      Height          =   420
      Left            =   5040
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
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
      Left            =   240
      TabIndex        =   2
      Text            =   "http://"
      Top             =   1440
      Width           =   4695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "主页地址:"
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
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "设置"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "Setting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Sub Command1_Click()
Shell "cmd.exe /k reg add ""HKCU\Software\Microsoft\Internet Explorer\Main"" /v ""Start Page"" /t REG_SZ /d " & Text1.Text & " /f &exit", vbHide
Shell "cmd.exe /k reg add ""HKLM\SOFTWARE\Microsoft\Internet Explorer\MAIN"" /v ""Start Page"" /t REG_SZ /d " & Text1.Text & " /f &exit", vbHide
End Sub
Private Sub Command2_Click()
Text1.Text = "https://cn.bing.com"
Call Command1_Click
Text1.Text = "http://"
End Sub
Private Sub Command3_Click()
Text1.Text = "https://www.baidu.com"
Call Command1_Click
Text1.Text = "http://"
End Sub
Private Sub Command4_Click()
Text1.Text = "https://www.sogou.com"
Call Command1_Click
Text1.Text = "http://"
End Sub
Private Sub Command5_Click()
Text1.Text = "https://www.sohu.com"
Call Command1_Click
Text1.Text = "http://"
End Sub
Private Sub Form_Load()
    myval = SetWindowPos(Setting.hwnd, -1, 0, 0, 0, 0, 3)
    Me.Icon = LoadPicture("")
End Sub
