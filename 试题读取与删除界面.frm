VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form main_del 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����༭"
   ClientHeight    =   10380
   ClientLeft      =   3675
   ClientTop       =   2025
   ClientWidth     =   17355
   Icon            =   "�����ȡ��ɾ������.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10380
   ScaleWidth      =   17355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "�½�"
      Height          =   390
      Left            =   10920
      TabIndex        =   3
      Top             =   9960
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�˳�"
      Height          =   390
      Left            =   14985
      TabIndex        =   2
      Top             =   9960
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ɾ��"
      Height          =   390
      Left            =   13005
      TabIndex        =   1
      Top             =   9960
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   -1320
      Top             =   3420
      Visible         =   0   'False
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"�����ȡ��ɾ������.frx":324A
      OLEDBString     =   $"�����ȡ��ɾ������.frx":32D1
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "VBѡ����"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "�����ȡ��ɾ������.frx":3358
      Height          =   10575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20175
      _ExtentX        =   35586
      _ExtentY        =   18653
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   22
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "ID"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "��Ŀ"
         Caption         =   "��Ŀ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "ѡ��A"
         Caption         =   "ѡ��A"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "ѡ��B"
         Caption         =   "ѡ��B"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "ѡ��C"
         Caption         =   "ѡ��C"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "ѡ��D"
         Caption         =   "ѡ��D"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "��"
         Caption         =   "��"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         Locked          =   -1  'True
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   6944.882
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2835.213
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2445.166
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2640.189
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2475.213
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   2775.118
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuPop 
      Caption         =   "�˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuPop4 
         Caption         =   "�������"
      End
      Begin VB.Menu mnuPop1 
         Caption         =   "�½�����"
      End
      Begin VB.Menu mnuPop2 
         Caption         =   "ɾ������"
      End
      Begin VB.Menu mnuPop3 
         Caption         =   "����ʱ�趨"
      End
   End
End
Attribute VB_Name = "main_del"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Dim myval As Long
  If Adodc1.Recordset.RecordCount > 0 Then
    myval = MsgBox("�Ƿ�ɾ��ָ����¼��", vbYesNo, "ѡ�������ϵͳ")
    If myval = vbYes Then
      Adodc1.Recordset.Delete
      Adodc1.Recordset.Update
    End If
  Else
    MsgBox "ϵͳû��Ҫɾ�������ݣ�", , "ѡ�������ϵͳ"
  End If
End Sub
Private Sub Command2_Click()
Unload main_jbzl_jsr
  End
End Sub

Private Sub Command3_Click()
    
   main_jbzl_jsr.Show
    
End Sub

Private Sub DataGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbKeyRButton Then PopupMenu mnuPop

End Sub

Private Sub Form_Activate()
 Unload main_jbzl_jsr
 'Adodc1.Recordset.Update
End Sub

Private Sub Form_Load()
Form2.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload main_jbzl_jsr
Unload Form1
Unload Form2
Unload Form3
Unload Form4
End Sub

Private Sub mnuPop1_Click()
main_jbzl_jsr.Show
End Sub

Private Sub mnuPop2_Click()
   Command1_Click
End Sub

Private Sub mnuPop3_Click()
 Form4.Show: Form2.Hide
End Sub
