VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "倒计时设定"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4800
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
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
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   2
      Text            =   " "
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1440
      MaxLength       =   2
      TabIndex        =   1
      Text            =   " "
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   " ："
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  If Val(text1.Text) > 60 Or Val(Text2.Text) > 60 Then
  MsgBox "数据超出范围", 48
  text1.Text = ""
  Text2.Text = ""
   Else
    MsgBox "设定成功"
   mm = text1.Text
  ss = Text2.Text
  Call xrsj
   main_del.Show
   Form2.Hide
   Form4.Hide
 
  
 End If
  
 
End Sub

Private Sub Command2_Click()
Form4.Hide
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

 If KeyAscii = 27 Then Form2.Show: Form4.Hide
End Sub

Private Sub Form_Load()
 text1 = ""
 text1.PasswordChar = "*"
 text1.SelStart = 0
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 If KeyAscii < 48 And KeyAscii <> 13 And KeyAscii <> 8 And KeyAscii <> 27 Then
    KeyAscii = 0
    MsgBox "输入字符必须为数字"
    ElseIf KeyAscii > 57 And KeyAscii <> 13 And KeyAscii <> 8 Then
    KeyAscii = 0
    MsgBox "输入字符必须为数字"
   
   End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
 If KeyAscii < 48 And KeyAscii <> 13 And KeyAscii <> 8 And KeyAscii <> 27 Then
    KeyAscii = 0
    MsgBox "输入字符必须为数字"
    ElseIf KeyAscii > 57 And KeyAscii <> 13 And KeyAscii <> 8 Then
    KeyAscii = 0
    MsgBox "输入字符必须为数字"
   
   End If
End Sub
