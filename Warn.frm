VERSION 5.00
Begin VB.Form Warn 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "警告！"
   ClientHeight    =   2475
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   6555
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   5400
      Top             =   1680
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   4560
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "不要切换模式"
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
      Left            =   3600
      TabIndex        =   3
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定切换"
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
      TabIndex        =   2
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "请您再三思索，再切换模式！"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "切换浏览模式会导致您的工作进度丢失！"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5415
   End
End
Attribute VB_Name = "Warn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Sub Command1_Click()
If Text1.Text = "chenjin" Then
zhuan.Show
Locker.Hide
Unload Me
ElseIf Text1.Text = "small" Then
Form3.Show
Locker.Hide
Unload Me
End If
End Sub
Private Sub Command2_Click()
Me.Hide
End Sub
Private Sub Form_Load()
Me.Icon = LoadPicture("")
    myval = SetWindowPos(Form1.hwnd, -1, 0, 0, 0, 0, 3)
End Sub
Private Sub Timer1_Timer()
Me.Left = 300
Me.Top = 400
Timer1.Enabled = False
End Sub
