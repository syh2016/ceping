VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ѡ�������ϵͳ"
   ClientHeight    =   9405
   ClientLeft      =   2190
   ClientTop       =   510
   ClientWidth     =   8175
   FillStyle       =   0  'Solid
   Icon            =   "��½����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   8175
   StartUpPosition =   2  '��Ļ����
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      IMEMode         =   2  'OFF
      ItemData        =   "��½����.frx":C84A
      Left            =   2520
      List            =   "��½����.frx":C84C
      TabIndex        =   7
      Text            =   "��ѡ���û�"
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Timer Timer1 
      Left            =   480
      Top             =   1080
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   2520
      MaxLength       =   10
      TabIndex        =   6
      Text            =   " "
      Top             =   4050
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "��ʾ/����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4200
      MaskColor       =   &H00000001&
      TabIndex        =   5
      Top             =   4050
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "��½"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2520
      TabIndex        =   4
      Top             =   5640
      Width           =   3015
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�޸�����"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2520
      TabIndex        =   1
      Top             =   4845
      Width           =   3015
   End
   Begin VB.Label Label4 
      Caption         =   "ѡ�������ϵͳ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   735
      Left            =   2100
      TabIndex        =   8
      Top             =   1320
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "���ߣ����ֻ�ܶ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   7680
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "����ʱ�䣺2014.08.19"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   8280
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "��ӭʹ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   855
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim s, Y, h, zzz, xxx As Boolean: Dim m As Integer
 
 
Private Sub Combo1_Click()
    Command1.Enabled = True
      Command1.SetFocus
  If Combo1.Text = Combo1.List(��ͨ�û�) Then
    
    
     xxx = True
     text1.Enabled = False
     Command2.Enabled = False
     Command3.Enabled = False
    Else
    
     text1.Enabled = True
     text1.SetFocus
     
     xxx = False
     text1.Enabled = True
     Command2.Enabled = True
     Command3.Enabled = True
 
    
 End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0  ' ��ֹ��������
End Sub

Private Sub Command1_Click()
    
   
  If xxx = True Then
    Form1.Show
    Form2.Hide
  Else
     
  Call dlmmjc
  
  
  
  End If
 
End Sub



Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  text1.PasswordChar = ""
End Sub

Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  text1.PasswordChar = "*"
  text1.SetFocus
End Sub

Private Sub Command3_Click()
  Form3.Show
  Form2.Hide
  Form3.text1.SetFocus
End Sub



Private Sub Form_Activate()
    Command1.Enabled = False
   
     xxx = True
     text1.Enabled = False
     Command2.Enabled = False
     Command3.Enabled = False
     Combo1.SetFocus
End Sub

Private Sub Form_Click()
  Cls
End Sub







Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    m = m + 1
   If KeyCode = vbKeyS Then s = True
   If KeyCode = vbKeyY Then Y = True
   If KeyCode = vbKeyH Then h = True
   If KeyCode = vbKeyZ Then zzz = True
   If s And zzz Then Form4.Show: Form2.Hide: s = False: zzz = False
   If s And Y And h Then z = "123456": Print "�����ѻ�ԭ": s = False: Y = False: h = False
  
    
   If m > 3 Then m = 0: Cls
    
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyS Then s = False
    If KeyCode = vbKeyY Then Y = False
    If KeyCode = vbKeyH Then h = False
    If KeyCode = vbKeyZ Then zzz = False
End Sub

Private Sub Form_Load()

  
With Combo1
   Combo1.AddItem "��ͨ�û�"
   Combo1.AddItem "����Ա"
End With
 Timer1.Interval = 100
 Label3.Caption = "��ӭʹ��"
  Call jc
  Load Form1
  Load Form3
  Load Form4
  Form1.Timer1.Enabled = False
  Call dr
  z = hy(zz)
  
  
 
 
   
End Sub

  

Private Sub Form_Unload(Cancel As Integer)
  zz = jm(z)
  Call xr
 Unload Form1
 Unload Form2
 Unload Form3
 Unload Form4
End Sub

Private Sub text1_Change()
text1.PasswordChar = "*"

End Sub

Private Sub Text1_Click()
 text1 = ""
 text1.PasswordChar = "*"
 text1.SelStart = 0
End Sub

Private Sub Text1_DblClick()
  text1.Text = ""
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 
 '�Ƿ����ֻ������������
  'ASCII��48��57��ʾ�����̵�0��9��������13Ϊ���س�����8Ϊ���˸�
    If Index = 0 Or Index = 3 Then
        If KeyAscii < 48 And KeyAscii <> 13 And KeyAscii <> 8 Then
            'KeyAscii = 0����ȡ���û�����һ������
            KeyAscii = 0
            MsgBox "�����ַ�������0-9֮�䣡", 48, "��ʾ"
        ElseIf KeyAscii > 57 And KeyAscii <> 13 And KeyAscii <> 8 Then
            KeyAscii = 0
            MsgBox "�����ַ�������0-9֮�䣡", 48, "��ʾ"
        End If
    End If
  
 
 
  If KeyAscii = 13 Then Call dlmmjc
  
   
End Sub

Private Sub Timer1_Timer()
Label3.Left = Label3.Left + 200
If Label3.Left >= Form2.Width Then Label3.Left = -Label3.Width
End Sub
