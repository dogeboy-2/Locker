VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form zhuan 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "×¨×¢Ä£Ê½"
   ClientHeight    =   10800
   ClientLeft      =   0
   ClientTop       =   345
   ClientWidth     =   18960
   Icon            =   "×¨×¢Ä£Ê½.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   10800
   ScaleWidth      =   18960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "Ö÷Ò³"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14280
      TabIndex        =   7
      Top             =   10080
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ë¢ÐÂ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12960
      TabIndex        =   6
      Top             =   10080
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   10080
      Width           =   10335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Default         =   -1  'True
      Height          =   375
      Left            =   11040
      TabIndex        =   4
      Top             =   10200
      Width           =   615
   End
   Begin VB.CommandButton EX 
      Caption         =   "ÍË³ö"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   17040
      TabIndex        =   3
      Top             =   10080
      Width           =   855
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
      Height          =   495
      Left            =   15720
      TabIndex        =   2
      Top             =   10080
      Width           =   855
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
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   10080
      Width           =   735
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   9855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   19335
      ExtentX         =   34105
      ExtentY         =   17383
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
End
Attribute VB_Name = "zhuan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BACK_Click() '·µ»Ø
    On Error Resume Next
    WebBrowser1.GoBack
    GoAdDress.Text = WebBrowser1.LocationURL
End Sub
Private Sub Command1_Click()
On Error Resume Next
    WebBrowser1.Navigate Trim(Text1.Text) '´ò¿ªÍøÒ³
    End Sub
Private Sub Command2_Click()
On Error Resume Next
WebBrowser1.REFRESH
End Sub
Private Sub Command3_Click()
On Error Resume Next
WebBrowser1.Navigate "https://cn.bing.com"
End Sub
Private Sub EX_Click()
Locker.Show
Locker.WebBrowser1.Navigate (zhuan.WebBrowser1.LocationURL)
Unload Me
End Sub
Private Sub Form_Load()
WebBrowser1.Navigate (Locker.WebBrowser1.LocationURL)
End Sub
Private Sub FORWARD_Click() 'Ç°½ø
On Error Resume Next
    WebBrowser1.GoForward
    GoAdDress.Text = WebBrowser1.LocationURL
End Sub
Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
Cancel = True
WebBrowser1.Navigate2 WebBrowser1.Document.activeElement.href
End Sub
Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
Text1.Text = WebBrowser1.LocationURL
End Sub
Private Sub Form_Resize()
WebBrowser1.Height = Me.Height - 945
WebBrowser1.Width = Me.Width
BACK.Top = WebBrowser1.Top + WebBrowser1.Height + 200
Command1.Top = WebBrowser1.Top + WebBrowser1.Height + 200
Command2.Top = WebBrowser1.Top + WebBrowser1.Height + 200
Command3.Top = WebBrowser1.Top + WebBrowser1.Height + 200
FORWARD.Top = WebBrowser1.Top + WebBrowser1.Height + 200
EX.Top = WebBrowser1.Top + WebBrowser1.Height + 200
Text1.Top = WebBrowser1.Top + WebBrowser1.Height + 200
End Sub
