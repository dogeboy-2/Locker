VERSION 5.00
Begin VB.Form LockerBar 
   BackColor       =   &H8000000B&
   BorderStyle     =   0  'None
   Caption         =   "Form7"
   ClientHeight    =   870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11850
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   870
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
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
      Left            =   7920
      TabIndex        =   1
      Text            =   " 搜索..."
      Top             =   240
      Width           =   3840
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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7695
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000000&
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   11820
   End
End
Attribute VB_Name = "LockerBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Dim Xa As String, Ya As String, FClick As String
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Xa = x
Ya = y
FClick = "Yes"
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Xa = ""
Ya = ""
FClick = "No"
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If FClick = "Yes" Then
Me.Left = Me.Left - Xa + x
Me.Top = Me.Top - Ya + y
End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
'如果按下回车键，则访问网页
If KeyAscii = vbKeyReturn Then
Locker.WebBrowser1.Navigate "https://cn.bing.com/search?q=" + (UTF8EncodeURI(Text3.Text))
Unload Me
Locker.Show
End If
End Sub
Private Sub Text3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
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
Private Sub Form_Load()
    Dim retValue As Long
    retValue = SetWindowPos(Me.hwnd, HWND_TOPMOST, Me.CurrentX, Me.CurrentY, 800, 70, SWP_SHOWWINDOW)
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer) '回车键，需要改按钮
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
Locker.WebBrowser1.Navigate Trim(Text1.Text)
Unload Me
Locker.Show
    End If
    If Text1.ForeColor = &H808080 Then
Text1.ForeColor = &H0&
End If
End Sub
Private Sub Form_Resize()
On Error Resume Next
Shape1.Height = Me.Height - 10
Shape1.Top = 10
Shape1.Width = Me.Width - 10
Shape1.Left = 10
End Sub

