Attribute VB_Name = "Module1"
 Public zz, a, b, c, d, e, f, a1, b1, c1, d1, e1, f1, X, Y, z, nn, mm, ss   'zz��������   z ����
 
 Public n1, n2, n3, n4 As Boolean
 
   Dim rs1 As New ADODB.Recordset
   Function Cnn() As ADODB.Connection    '���庯��
  Set Cnn = New ADODB.Connection
  '����һ�����ݿ�����
  Cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\ѡ�������.mdb;Persist Security Info=False"
End Function
  
 Sub Main()        '��������
 
  Load Form2
     
  End Sub
  

Function hy(zz)                '���뻹ԭ
   a = Mid(zz, 1, 1)
   b = Mid(zz, 3, 1)
   c = Mid(zz, 5, 1)
   d = Mid(zz, 6, 1)
   e = Mid(zz, 4, 1)
   f = Mid(zz, 2, 1)
  hy = a & b & c & d & e & f
 
End Function
Function jm(z)                 '�������
    a1 = Mid(z, 1, 1)
   b1 = Mid(z, 2, 1)
   c1 = Mid(z, 3, 1)
   d1 = Mid(z, 4, 1)
   e1 = Mid(z, 5, 1)
   f1 = Mid(z, 6, 1)
   jm = a1 & f1 & b1 & e1 & c1 & d1
   
End Function
Sub mmxgjc()     '�����޸ļ��
    
  If Form3.text1.Text = "" Then
       MsgBox "������ԭʼ���룡"
       Form3.Text2.Text = ""
       Form3.Text3.Text = ""
       Form3.text1.SetFocus
   ElseIf Not Form3.text1.Text = "" = True And Form3.Text2.Text = "" Then
      
       MsgBox "�������趨���룡"
        Form3.Text3.Text = ""
        Form3.Text2.SetFocus
   ElseIf Not Form3.text1.Text = "" = True And Len(Form3.Text2.Text) <> 6 Then
      MsgBox "��������λ�����룡"
       Form3.Text2.Text = ""
       Form3.Text3.Text = ""
      Form3.Text2.SetFocus
   ElseIf Form3.Text3.Text = "" Then
        MsgBox "��������֤���룡"
        Form3.Text3.SetFocus
   ElseIf Not Form3.text1.Text = z = True And Not Form3.text1.Text = "" = True And Not Form3.Text2.Text = "" = True Then
      MsgBox "�����������"
      Form3.text1.Text = ""
      Form3.Text2.Text = ""
       Form3.Text3.Text = ""
       Form3.text1.SetFocus
   ElseIf Form3.text1.Text = z And Not Form3.Text2.Text = Form3.Text3.Text = True Then
       MsgBox "�趨��������֤���벻һ�£�"
        Form3.Text2.Text = ""
       Form3.Text3.Text = ""
       Form3.Text2.SetFocus
   ElseIf Form3.text1.Text = z And Form3.Text2.Text = Form3.Text3.Text And Len(Form3.Text2.Text) <> 6 Then
       MsgBox "��������λ�����룡"
        Form3.Text2.Text = ""
       Form3.Text3.Text = ""
   ElseIf Form3.text1.Text = z And Len(Form3.Text2.Text) = 6 Then
        z = Form3.Text2.Text
        MsgBox "�����޸ĳɹ���"
       
        Form3.text1.Text = ""
        Form3.Text2.Text = ""
        Form3.Text3.Text = ""
        Form3.Hide
        Form2.Show
       
  
      
  End If
End Sub

 Sub jc()      '�������Ƿ��ظ�����
     
   If App.PrevInstance = True Then '����ǰһ�汾
    MsgBox "�˳�ʽ�Ѿ���ִ���У�", 48
    End
    
   End If
 
 End Sub
 Sub dr()     '  ��mima.txt����zz����
    rs1.Open "select * from ����", Cnn, adOpenKeyset, adLockOptimistic
    zz = rs1.Fields(1)
    rs1.Update
     rs1.Close
  '  Open App.Path & "\mima.txt" For Input As #1         '�򿪳���·����mima.txt
  ' Line Input #1, zz                                  ' �����ݶ������ZZ��
   'Close #1  ' �ر��ļ�

 End Sub
Sub xr()       ' ��zz����д��mima.txt
    rs1.Open "select * from ����", Cnn, adOpenKeyset, adLockOptimistic
    rs1.Fields(1) = zz
    rs1.Update
    rs1.Close
  
 ' Open App.Path & "\mima.txt" For Output As #1      '��������mima.txt
 ' Print #1, zz                                     'ZZ����д��mima.txt
 ' Close #1 '�ر�mima.txt

 
End Sub
Sub dlmmjc()     '��½������
   If Form2.text1.Text = z Then
   ' Form1.Show
   ' Form2.Hide
     
     Load main_del
    
     main_del.Show
    ElseIf Form2.text1.Text = "" Then
     msg = MsgBox("���������룡", 0, " ��ʾ")
    Form2.text1.SetFocus
   ElseIf Len(Form2.text1.Text) <> 6 Then
   msg = MsgBox("������λ���֣�", 0, " ��ʾ")
    Form2.text1.Text = ""
    Form2.text1.SetFocus
   Else
   
    msg = MsgBox("��������������������룡 ", 0, "��ʾ"): Form2.text1.Text = ""
   Form2.text1.SetFocus
    X = X + 1
    
  End If
   
  If X > 3 Then Unload Form1: Unload Form3: Unload Form2
End Sub
  
Sub xrsj()    'д��ʱ��

    rs1.Open "select * from ʱ��", Cnn, adOpenKeyset, adLockOptimistic
    rs1.Fields(1) = mm
    rs1.Fields(2) = ss
    rs1.Update
    rs1.Close
   
   ' Open App.Path & "\m.txt" For Output As #1      '��������m.txt
  'Print #1, mm                                    'mm����д��mima.txt
 ' Close #1                                          '�ر�m.txt
   '  Open App.Path & "\s.txt" For Output As #1      '��������s.txt
 ' Print #1, ss                                   'ss����д��s.txt
 ' Close #1                                          '�ر�s.txt
End Sub
Sub dqsj()  ' ��ȡʱ��

    rs1.Open "select * from ʱ��", Cnn, adOpenKeyset, adLockOptimistic
    mm = rs1.Fields(1)
    ss = rs1.Fields(2)
    rs1.Update
    rs1.Close

   ' Open App.Path & "\m.txt" For Input As #1         '�򿪳���·����m .txt
  '  Line Input #1, mm                                  ' �����ݶ������mm��
  ' Close #1   ' �ر��ļ�
  ' Open App.Path & "\s.txt" For Input As #1         '�򿪳���·����s .txt
  ' Line Input #1, ss                                  ' �����ݶ������ss��
  ' Close #1  ' �ر��ļ�
End Sub
 

