VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "选择题测试系统"
   ClientHeight    =   8280
   ClientLeft      =   1590
   ClientTop       =   1335
   ClientWidth     =   12450
   Icon            =   "主界面.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   12450
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   10800
      Top             =   600
   End
   Begin VB.CommandButton Command8 
      Caption         =   "测试"
      Enabled         =   0   'False
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
      Left            =   5760
      TabIndex        =   24
      ToolTipText     =   "测试模式与浏览模式切换按钮"
      Top             =   360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "转到"
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
      Left            =   7440
      TabIndex        =   21
      Top             =   6600
      Width           =   975
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
      Left            =   9000
      MaxLength       =   4
      TabIndex        =   20
      Text            =   " "
      Top             =   6570
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "最后一题 "
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
      Left            =   5880
      TabIndex        =   18
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "第一题"
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
      Left            =   2280
      TabIndex        =   17
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "上一题"
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
      Left            =   3480
      TabIndex        =   16
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "下一题"
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
      Left            =   4680
      TabIndex        =   15
      Top             =   6600
      Width           =   1215
   End
   Begin VB.OptionButton Option4 
      Caption         =   "··"
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
      Left            =   360
      TabIndex        =   4
      Top             =   6030
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Option1"
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
      Left            =   360
      TabIndex        =   3
      Top             =   5490
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option1"
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
      Left            =   360
      TabIndex        =   1
      Top             =   4890
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
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
      Left            =   360
      TabIndex        =   2
      Top             =   4170
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   240
      Top             =   7560
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      ConnectMode     =   1
      CursorLocation  =   2
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   100
      BOFAction       =   1
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"主界面.frx":324A
      OLEDBString     =   $"主界面.frx":32D1
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "VB选择题"
      Caption         =   "           数据库"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "楷体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "时间："
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3600
      TabIndex        =   25
      Top             =   1200
      Width           =   990
   End
   Begin VB.Label Label10 
      Caption         =   "答案："
      DataField       =   " "
      DataSource      =   " "
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
      Left            =   6240
      TabIndex        =   23
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "题"
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
      Left            =   9960
      TabIndex        =   22
      Top             =   6600
      Width           =   495
   End
   Begin VB.Label Label8 
      Caption         =   "第："
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
      Left            =   8640
      TabIndex        =   19
      Top             =   6600
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   " "
      DataField       =   "答案"
      DataSource      =   "Adodc1"
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
      Left            =   7560
      TabIndex        =   14
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      DataField       =   "ID"
      DataSource      =   "Adodc1"
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
      Left            =   240
      TabIndex        =   13
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      DataField       =   "选项D"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   840
      TabIndex        =   12
      Top             =   6120
      Width           =   990
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      DataField       =   "选项C"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   840
      TabIndex        =   11
      Top             =   5520
      Width           =   990
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      DataField       =   "选项B"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   840
      TabIndex        =   10
      Top             =   4920
      Width           =   990
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      DataField       =   "选项A"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   840
      TabIndex        =   9
      Top             =   4200
      Width           =   990
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      DataField       =   "题目"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "楷体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   1080
      TabIndex        =   8
      Top             =   1920
      Width           =   10455
   End
   Begin VB.Label Label43 
      Caption         =   "选项C"
      Height          =   375
      Left            =   -3240
      TabIndex        =   7
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label Label33 
      Caption         =   "选项B"
      Height          =   375
      Left            =   -3240
      TabIndex        =   6
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label22 
      Caption         =   "选项A"
      Height          =   375
      Left            =   -3240
      TabIndex        =   5
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "题目"
      Height          =   375
      Left            =   -3240
      TabIndex        =   0
      Top             =   3360
      Width           =   855
   End
   Begin VB.Menu mnuPop 
      Caption         =   "菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuPop2 
         Caption         =   "试题测试"
      End
      Begin VB.Menu mnuPop1 
         Caption         =   "试题浏览"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n, X, xx    As Integer
Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub Command1_Click()
    
   Adodc1.Recordset.MoveNext
  
   If Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveLast
   
End Sub

Private Sub Command2_Click()
 
   Adodc1.Recordset.MovePrevious
  
   If Adodc1.Recordset.BOF Then Adodc1.Recordset.MoveFirst
    
End Sub


Private Sub Command3_Click()
 
  If text1.Text <> "" Then
    lblPosition = Adodc1.Recordset.AbsolutePosition  '判断当前记录位置
    Adodc1.Recordset.Move (Val(text1.Text) - lblPosition)
  End If
  If Val(text1.Text) > Adodc1.MaxRecords Then
    Adodc1.Recordset.MoveLast
   ElseIf Val(text1.Text) < 0 Then
    
    Adodc1.Recordset.MoveFirst
  End If
  
  End Sub

Private Sub Command4_Click()
Adodc1.Recordset.MoveFirst
 
End Sub

Private Sub Command5_Click()
Adodc1.Recordset.MoveLast
End Sub

Private Sub Command6_Click()
    
    MsgBox "你的总分是" & n & "分"
End Sub

 
      

Private Sub Command8_Click()
     xx = -xx
      If xx = -1 Then
      Command8.Caption = "浏览"

      Call dqsj
      n = 0
     Timer1.Enabled = True
      Label12.Visible = True
      Label12.Enabled = True
     Option1.Visible = True
     Option2.Visible = True
      Option3.Visible = True
      Option4.Visible = True
      Option1.Enabled = True
     Option2.Enabled = True
       Option3.Enabled = True
      Option4.Enabled = True
      Command4.Visible = False
      Command2.Visible = False
      Command1.Visible = False
      Command5.Visible = False
      Command3.Visible = False
      Command1.Enabled = False
      Command2.Enabled = False
      Command3.Enabled = False
      Command4.Enabled = False
      Command5.Enabled = False
      Label7.Visible = False
      Label8.Visible = False
     text1.Visible = False
      Label7.Enabled = False
      Label8.Enabled = False
      text1.Enabled = False
       
    Adodc1.Recordset.MoveFirst
     X = 1
    Label9.Visible = False
    Label10.Visible = False
     Label9.Enabled = False
    Label10.Enabled = False
    
     Else
       MsgBox "你的总分是" & n & "分"
      Command8.Caption = "测试"
       
       Call dqsj
       Timer1.Enabled = False
      Label12.Visible = False
      Label12.Enabled = False
       
       Option1.Visible = False
       Option2.Visible = False
       Option3.Visible = False
       Option4.Visible = False
       Option1.Enabled = False
       Option1.Enabled = False
       Option2.Enabled = False
       Option3.Enabled = False
      Option4.Enabled = False
      Command4.Visible = True
      Command2.Visible = True
      Command1.Visible = True
      Command5.Visible = True
      Command3.Visible = True
      Command1.Enabled = True
      Command2.Enabled = True
      Command3.Enabled = True
      Command4.Enabled = True
      Command5.Enabled = True
      
      Label7.Visible = True
      Label8.Visible = True
      text1.Visible = True
      Label7.Enabled = True
      Label8.Enabled = True
      text1.Enabled = True
      Label9.Visible = True
     Label10.Visible = True
     Label9.Enabled = True
     Label10.Enabled = True
      n = 0
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyDown Then Command1_Click
If KeyCode = vbKeyUp Then Command2_Click

End Sub

Private Sub Form_Load()
   
    Call jc
     xx = 1
      Timer1.Enabled = False
      Label12.Visible = False
      Label12.Enabled = False
       
       X = 1: n = 0
      Option1.Value = False
      Option2.Value = False
      Option3.Value = False
      Option4.Value = False
      
     text1.Text = "1"
     
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbKeyRButton Then PopupMenu mnuPop

End Sub

Private Sub Form_Unload(Cancel As Integer)
 Call xr
 Unload Form2: Unload Form3
End Sub

Private Sub mnuPop1_Click()

      MsgBox "你的总分是" & n & "分"
      Command8.Caption = "测试"
       
       Call dqsj
       Timer1.Enabled = False
      Label12.Visible = False
      Label12.Enabled = False
       
       Option1.Visible = False
       Option2.Visible = False
       Option3.Visible = False
       Option4.Visible = False
       Option1.Enabled = False
       Option1.Enabled = False
       Option2.Enabled = False
       Option3.Enabled = False
      Option4.Enabled = False
      Command4.Visible = True
      Command2.Visible = True
      Command1.Visible = True
      Command5.Visible = True
      Command3.Visible = True
      Command1.Enabled = True
      Command2.Enabled = True
      Command3.Enabled = True
      Command4.Enabled = True
      Command5.Enabled = True
      
      Label7.Visible = True
      Label8.Visible = True
      text1.Visible = True
      Label7.Enabled = True
      Label8.Enabled = True
      text1.Enabled = True
      Label9.Visible = True
     Label10.Visible = True
     Label9.Enabled = True
     Label10.Enabled = True
      n = 0
 
End Sub

Private Sub mnuPop2_Click()
   
      Call dqsj
      n = 0
      Timer1.Enabled = True
      Label12.Visible = True
      Label12.Enabled = True
      Option1.Visible = True
      Option2.Visible = True
      Option3.Visible = True
      Option4.Visible = True
      Option1.Enabled = True
      Option2.Enabled = True
      Option3.Enabled = True
      Option4.Enabled = True
      Command4.Visible = False
      Command2.Visible = False
      Command1.Visible = False
      Command5.Visible = False
      Command3.Visible = False
      Command1.Enabled = False
      Command2.Enabled = False
      Command3.Enabled = False
      Command4.Enabled = False
      Command5.Enabled = False
      Label7.Visible = False
      Label8.Visible = False
      text1.Visible = False
      Label7.Enabled = False
      Label8.Enabled = False
      text1.Enabled = False
       
     Adodc1.Recordset.MoveFirst
     X = 1
     Label9.Visible = False
     Label10.Visible = False
     Label9.Enabled = False
     Label10.Enabled = False
     
    
End Sub

Private Sub Option1_Click()
   If Option1.Value = True And Adodc1.Recordset.Fields("答案") = "A" Then n = n + 2
  Adodc1.Recordset.MoveNext
   Option1.Value = False
   If Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveLast: MsgBox "你的总分是" & n & "分": Unload Form1
    
End Sub

Private Sub Option2_Click()
  If X = 1 Then n = n + 2: X = 0
  If Option2.Value = True And Adodc1.Recordset.Fields("答案") = "B" Then n = n + 2
    Adodc1.Recordset.MoveNext
     Option2.Value = False
   If Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveLast: MsgBox "你的总分是" & n & "分": Unload Form1
    
      
End Sub

Private Sub Option3_Click()
  If Option3.Value = True And Adodc1.Recordset.Fields("答案") = "C" Then n = n + 2
   Adodc1.Recordset.MoveNext
    Option3.Value = False
   If Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveLast: MsgBox "你的总分是" & n & "分": Unload Form1
    
End Sub

Private Sub Option4_Click()
 If Option4.Value = True And Adodc1.Recordset.Fields("答案") = "D" Then n = n + 2
   Adodc1.Recordset.MoveNext
    Option4.Value = False
   If Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveLast: MsgBox "你的总分是" & n & "分": Unload Form1
  
End Sub

Private Sub text1_Change()
 If text1.Text = "0" Then MsgBox "输入数字不能为0": text1.Text = ""

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii < 48 And KeyAscii <> 13 And KeyAscii <> 8 Then
    KeyAscii = 0
    MsgBox "输入字符必须为数字"
    ElseIf KeyAscii > 57 And KeyAscii <> 13 And KeyAscii <> 8 Then
    KeyAscii = 0
    MsgBox "输入字符必须为数字"
   
   End If
   
End Sub

Private Sub Option5_Click()
  Option5.Caption = Adodc1.Recordset.Fields("题目")
End Sub

Private Sub Timer1_Timer()
  ss = ss - 1
  If ss = -1 Then ss = 59
  If ss = 59 Then mm = mm - 1 ' 秒耗尽时分钟减少
  If mm < 10 Then Label12.Caption = "时间:" & "0" & Val(mm) & ":" & Val(ss)
  If mm > 10 Then Label12.Caption = "时间:" & Val(mm) & ":" & Val(ss)
  If ss < 10 Then Label12.Caption = "时间:" & Val(mm) & ":" & "0" & Val(ss)
  If mm < 10 And ss < 10 Then Label12.Caption = "时间:" & "0" & Val(mm) & ":" & "0" & Val(ss)   ' 保证分显示都为两位数且在合理范围内
   
  If mm = 0 And ss = 0 Then
   
   xx = -1
    Command8_Click
   Timer1.Enabled = False
  
  End If
End Sub
