VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form3 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ð¡´°Ä£Ê½"
   ClientHeight    =   8595
   ClientLeft      =   4245
   ClientTop       =   2490
   ClientWidth     =   12315
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   12315
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CommandButton Command4 
      Caption         =   "Í£Ö¹"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10560
      TabIndex        =   8
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ö÷Ò³"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ë¢ÐÂ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   120
      Width           =   6975
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6735
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   12285
      ExtentX         =   21669
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
   Begin VB.CommandButton GO 
      Caption         =   "·ÃÎÊ"
      Default         =   -1  'True
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton EXIT 
      Caption         =   "ÍË³ö"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11400
      MaskColor       =   &H00808080&
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton FORWARD 
      Caption         =   "¡ú"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton BACK 
      Caption         =   "¡û"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
Private Sub BACK_Click()
 WebBrowser1.GoBack
End Sub
Private Sub Command1_Click()
WebBrowser1.REFRESH
End Sub
Private Sub Command2_Click()
WebBrowser1.Navigate "https://cn.bing.com"
End Sub
Private Sub Command4_Click()
WebBrowser1.stop
End Sub
Private Sub EXIT_Click()
Locker.Show
  HomeAddress = Form3.WebBrowser1.LocationURL
   Locker.WebBrowser1.Navigate HomeAddress
Unload Me
End Sub
Private Sub FORWARD_Click()
On Error Resume Next
 WebBrowser1.GoForward
End Sub
Private Sub Go_Click()
On Error Resume Next
    WebBrowser1.Navigate Trim(Text1.Text) '´ò¿ªÍøÒ³
End Sub
Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
Cancel = True
WebBrowser1.Navigate2 WebBrowser1.Document.activeElement.href
End Sub
Private Sub Form_Load()
 Call Go_Click
  HomeAddress = Locker.WebBrowser1.LocationURL
    WebBrowser1.Navigate HomeAddress
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    WebBrowser1.Width = Form3.Width - 120
    WebBrowser1.Height = Form3.Height - 1000
End Sub
Private Sub WebBrowser1_TitleChange(ByVal Text As String)
Text1.Text = WebBrowser1.LocationURL
If WebBrowser1.Busy = True Then
Text1.ForeColor = &H808080
Else
Text1.ForeColor = &H0&
End If
If InternetGetConnectedState(0&, 0&) Then
Else
Form9.Show
End If
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer) '»Ø³µ¼ü£¬ÐèÒª¸Ä°´Å¥
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        Go_Click
        Text1.Text = Text1.Text
    End If
    If Text1.ForeColor = &H808080 Then
Text1.ForeColor = &H0&
End If
End Sub
Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
'Èç¹ûµ¥»÷ËÑË÷¿ò£¬ÈÃËÑË÷ÎÄ×ÖÏûÊ§
If Text1.ForeColor = &H808080 Then
Text1.ForeColor = &H0&
End If
End Sub
