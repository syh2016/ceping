VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form main_jbzl_jsr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10230
   ClientLeft      =   7785
   ClientTop       =   870
   ClientWidth     =   11565
   Icon            =   "������������.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10230
   ScaleWidth      =   11565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "������������.frx":324A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "������������.frx":3B24
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "������������.frx":43FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "������������.frx":4CD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "������������.frx":55B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "������������.frx":5E8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "������������.frx":6766
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   1535
      ButtonWidth     =   1032
      ButtonHeight    =   1482
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   " ���� "
            Key             =   "add"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
            Key             =   "save"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ȡ��"
            Key             =   "cancel"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
            Key             =   "close"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame Frame1 
      Height          =   9240
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   20385
      Begin VB.TextBox Text1 
         Height          =   330
         Index           =   6
         Left            =   3840
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   1530
         Index           =   4
         Left            =   1200
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   5970
         Width           =   19005
      End
      Begin VB.TextBox Text1 
         Height          =   1530
         Index           =   5
         Left            =   1200
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   7680
         Width           =   19005
      End
      Begin VB.TextBox Text1 
         Height          =   1530
         Index           =   3
         Left            =   1200
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   4260
         Width           =   19005
      End
      Begin VB.TextBox Text1 
         Height          =   1530
         Index           =   2
         Left            =   1200
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   2550
         Width           =   19005
      End
      Begin VB.TextBox Text1 
         Height          =   1530
         Index           =   1
         Left            =   1200
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   840
         Width           =   19005
      End
      Begin VB.TextBox Text1 
         Height          =   330
         Index           =   0
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "��Ŀ"
         Height          =   165
         Left            =   360
         TabIndex        =   15
         Top             =   1523
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "ѡ��B"
         Height          =   165
         Left            =   360
         TabIndex        =   14
         Top             =   4943
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "ѡ��A"
         Height          =   165
         Left            =   360
         TabIndex        =   13
         Top             =   3233
         Width           =   795
      End
      Begin VB.Label Label5 
         Caption         =   "ѡ��D"
         Height          =   165
         Left            =   360
         TabIndex        =   12
         Top             =   8363
         Width           =   465
      End
      Begin VB.Label Label6 
         Caption         =   "ѡ��C"
         Height          =   165
         Left            =   360
         TabIndex        =   11
         Top             =   6653
         Width           =   750
      End
      Begin VB.Label Label7 
         Caption         =   "��"
         Height          =   165
         Left            =   3240
         TabIndex        =   9
         Top             =   323
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "ID"
         Height          =   165
         Left            =   360
         TabIndex        =   2
         Top             =   323
         Width           =   945
      End
   End
End
Attribute VB_Name = "main_jbzl_jsr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs1 As New ADODB.Recordset
Function Cnn() As ADODB.Connection    '���庯��
  Set Cnn = New ADODB.Connection
  '����һ�����ݿ�����
  Cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\ѡ�������.mdb;Persist Security Info=False"
End Function
'��������Toolbar�ؼ��ϰ�ť״̬�ĺ���
Function ControlState(state As Boolean)
  With Toolbar1
    If state = True Then
       .Buttons(1).Enabled = False
       .Buttons(2).Enabled = True
       For i = 1 To text1.UBound
           text1(i) = ""
           text1(i).Locked = False
       Next i
    Else
       .Buttons(1).Enabled = True
       .Buttons(2).Enabled = False
       For i = 1 To text1.UBound
           text1(i).Locked = True
       Next i
     End If
   End With
End Function
Private Sub Form_Load()
   Unload main_del
  Me.Caption = "�½�����"
  ControlState False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Form1
Unload Form2
Unload Form3
Unload Form4

End Sub

Private Sub Text1_GotFocus(Index As Integer)
  text1(Index).BackColor = &HFFFF00
  text1(Index).SelStart = 0
  text1(Index).SelLength = Len(text1(Index))
End Sub
Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  'If KeyCode = vbKeyReturn And Index < 5 Then Text1(Index + 1).SetFocus
End Sub
Private Sub Text1_LostFocus(Index As Integer)
  text1(Index).BackColor = &H80000005
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
     Case "add"
       ControlState True
       rs1.Open "select * from VBѡ����", Cnn, adOpenKeyset, adLockOptimistic
       If rs1.RecordCount > 0 Then
          rs1.MoveLast
          text1(0) = Format(Val(rs1.Fields("ID")) + 1, "000")
       Else
          text1(0) = "001"
       End If
       rs1.Close
       text1(1).SetFocus
     Case "save"
       On Error GoTo SaveErr
       If text1(0).Text = "" Or text1(1).Text = "" Or text1(2).Text = "" Or text1(3).Text = "" Or text1(4).Text = "" Or text1(5).Text = "" Or text1(6).Text = "" Then
          MsgBox "���ⲻ������", , "ѡ�������ϵͳ"
          Exit Sub
       End If
       rs1.Open "VBѡ����", Cnn, adOpenKeyset, adLockOptimistic
       rs1.AddNew
       For i = 0 To text1.UBound
         rs1.Fields(i) = text1(i)
       Next i
       rs1.Update
       rs1.Close
       MsgBox "���ݱ���ɹ���", , "ѡ�������ϵͳ"
       ControlState False
       Exit Sub
SaveErr:
        MsgBox Err.Description
     Case "cancel"
       For i = 1 To text1.UBound
           text1(i) = ""
           text1(i).Locked = True
       Next i
       ControlState False
     Case "close"
       Load main_del
       main_del.Show
  End Select
End Sub
