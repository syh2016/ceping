VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "密码修改"
   ClientHeight    =   4380
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   5670
   ControlBox      =   0   'False
   Icon            =   "密码修改界面.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   2520
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   2280
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "返回"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   5
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "确认"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   4
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   2520
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1440
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      IMEMode         =   3  'DISABLE
      Left            =   2520
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label3 
      Caption         =   "密码验证 ："
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "设定密码 ："
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "原始密码 ："
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b, c
Private Sub Command1_Click()
   
 Call mmxgjc
 
End Sub

Private Sub Command2_Click()
  Form2.Show
  Form3.Hide
  text1.Text = ""
  Text2.Text = ""
  Text3.Text = ""
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
   Form2.Show
   Form3.Hide
   text1.Text = ""
   Text2.Text = ""
   Text3.Text = ""
  End If
  If KeyAscii = 13 Then Call mmxgjc
End Sub

Private Sub Form_Load()
  Call jc
 
End Sub

Private Sub Text1_DblClick()
    text1.Text = ""
End Sub





Private Sub Text1_KeyPress(KeyAscii As Integer)
   
    If KeyAscii < 48 And KeyAscii <> 13 And KeyAscii <> 8 And KeyAscii <> 27 Then
    KeyAscii = 0
    text1.Text = ""
    MsgBox "输入字符必须为字母或数字"
    ElseIf KeyAscii > 90 And KeyAscii <> 13 And KeyAscii <> 8 Then
    KeyAscii = 0
     text1.Text = ""
    MsgBox "输入字符必须为字母或数字"
   
   End If
End Sub

Private Sub Text2_DblClick()
  Text2.Text = ""
End Sub




Private Sub Text2_KeyPress(KeyAscii As Integer)
 If KeyAscii < 48 And KeyAscii <> 13 And KeyAscii <> 8 And KeyAscii <> 27 Then
    KeyAscii = 0
    Text2.Text = ""
    MsgBox "输入字符必须为字母或数字"
    ElseIf KeyAscii > 90 And KeyAscii <> 13 And KeyAscii <> 8 Then
    KeyAscii = 0
     Text2.Text = ""
    MsgBox "输入字符必须为字母或数字"
   
   End If
End Sub

Private Sub Text3_DblClick()
   
   Text3.Text = ""
End Sub



Private Sub Text3_KeyPress(KeyAscii As Integer)
  If KeyAscii < 48 And KeyAscii <> 13 And KeyAscii <> 8 And KeyAscii <> 27 Then
    KeyAscii = 0
    Text3.Text = ""
    MsgBox "输入字符必须为字母或数字"
    ElseIf KeyAscii > 90 And KeyAscii <> 13 And KeyAscii <> 8 Then
    KeyAscii = 0
     Text3.Text = ""
    MsgBox "输入字符必须为字母或数字"
   
   End If
End Sub
