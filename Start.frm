VERSION 5.00
Begin VB.Form Form4 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   4710
   ClientLeft      =   4005
   ClientTop       =   3300
   ClientWidth     =   10125
   Icon            =   "Start.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   10125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Timer Timer4 
      Left            =   4320
      Top             =   3960
   End
   Begin VB.Timer Timer3 
      Left            =   5280
      Top             =   480
   End
   Begin VB.Timer Timer2 
      Left            =   1080
      Top             =   4080
   End
   Begin VB.Timer Timer1 
      Left            =   4200
      Top             =   480
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000000&
      FillColor       =   &H80000000&
      Height          =   4695
      Left            =   0
      Top             =   0
      Width           =   10095
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Locker X"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   72
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   2535
      Left            =   7680
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.Image Image1 
      Height          =   2550
      Left            =   480
      Picture         =   "Start.frx":1084A
      Top             =   1080
      Width           =   2550
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Timer2_Timer()
Label1.Left = 3120
Timer3.Enabled = True
Timer3.Interval = 20
End Sub
Private Sub Timer1_Timer()
Label1.Visible = True
Timer2.Enabled = True
Timer2.Interval = 65
End Sub
Private Sub Form_Load()
Timer1.Enabled = True
Timer1.Interval = 80
End Sub
Private Sub Timer3_Timer()
Label1.ForeColor = &H8000000C
Timer4.Enabled = True
Timer4.Interval = 1000
End Sub

Private Sub Timer4_Timer()
Unload Me
Locker.Show
End Sub
