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
      Left            =   13560
      TabIndex        =   7
      Top             =   10200
      Visible         =   0   'False
      Width           =   855
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
      Left            =   14520
      TabIndex        =   8
      Top             =   10200
      Visible         =   0   'False
      Width           =   855
   End
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
      Left            =   12600
      TabIndex        =   9
      Top             =   10200
      Visible         =   0   'False
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
      Left            =   11640
      TabIndex        =   5
      Top             =   10200
      Visible         =   0   'False
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
      Left            =   120
      TabIndex        =   2
      Top             =   10200
      Visible         =   0   'False
      Width           =   735
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
      Left            =   1200
      TabIndex        =   3
      Top             =   10200
      Visible         =   0   'False
      Width           =   10335
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6720
      ScaleHeight     =   345
      ScaleWidth      =   945
      TabIndex        =   0
      Top             =   6960
      Width           =   975
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "¹¤¾ßÀ¸"
         BeginProperty Font 
            Name            =   "Î¢ÈíÑÅºÚ"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -240
         TabIndex        =   1
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   10320
      Visible         =   0   'False
      Width           =   615
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   10815
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   19335
      ExtentX         =   34105
      ExtentY         =   19076
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
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Dim p2Xa As String, p2Ya As String, p2FClick As String
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
    myval = SetWindowPos(Form1.hwnd, -1, 0, 0, 0, 0, 3)
End Sub
Private Sub FORWARD_Click() 'Ç°½ø
On Error Resume Next
    WebBrowser1.GoForward
    GoAdDress.Text = WebBrowser1.LocationURL
End Sub
Private Sub Label1_Click()
BACK.Visible = Not BACK.Visible
Text1.Visible = Not Text1.Visible
Command1.Visible = Not Command1.Visible
Command2.Visible = Not Command2.Visible
Command3.Visible = Not Command3.Visible
FORWARD.Visible = Not FORWARD.Visible
EX.Visible = Not EX.Visible
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer) '»Ø³µ¼ü£¬ÐèÒª¸Ä°´Å¥
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        Command1_Click
        Text1.Text = Text1.Text
    End If
End Sub
Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
Cancel = True
WebBrowser1.Navigate2 WebBrowser1.Document.activeElement.href
End Sub
Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
Text1.Text = WebBrowser1.LocationURL
End Sub
Private Sub Form_Resize()
WebBrowser1.Height = Me.Height
WebBrowser1.Width = Me.Width
BACK.Top = Me.Height - 500
Command1.Top = Me.Height - 500
Command2.Top = Me.Height - 500
Command3.Top = Me.Height - 500
FORWARD.Top = Me.Height - 500
EX.Top = Me.Height - 500
Text1.Top = Me.Height - 500
End Sub
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
p2Xa = x
p2Ya = y
p2FClick = "Yes"
End Sub
Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
p2Xa = ""
p2Ya = ""
p2FClick = "No"
End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If p2FClick = "Yes" Then
Picture2.Left = Picture2.Left - p2Xa + x
Picture2.Top = Picture2.Top - p2Ya + y
End If
End Sub
