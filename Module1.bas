Attribute VB_Name = "Module1"
 Public zz, a, b, c, d, e, f, a1, b1, c1, d1, e1, f1, X, Y, z, nn, mm, ss   'zz加密密码   z 密码
 
 Public n1, n2, n3, n4 As Boolean
 
   Dim rs1 As New ADODB.Recordset
   Function Cnn() As ADODB.Connection    '定义函数
  Set Cnn = New ADODB.Connection
  '返回一个数据库连接
  Cnn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\选择题题库.mdb;Persist Security Info=False"
End Function
  
 Sub Main()        '启动程序
 
  Load Form2
     
  End Sub
  

Function hy(zz)                '密码还原
   a = Mid(zz, 1, 1)
   b = Mid(zz, 3, 1)
   c = Mid(zz, 5, 1)
   d = Mid(zz, 6, 1)
   e = Mid(zz, 4, 1)
   f = Mid(zz, 2, 1)
  hy = a & b & c & d & e & f
 
End Function
Function jm(z)                 '密码加密
    a1 = Mid(z, 1, 1)
   b1 = Mid(z, 2, 1)
   c1 = Mid(z, 3, 1)
   d1 = Mid(z, 4, 1)
   e1 = Mid(z, 5, 1)
   f1 = Mid(z, 6, 1)
   jm = a1 & f1 & b1 & e1 & c1 & d1
   
End Function
Sub mmxgjc()     '密码修改检测
    
  If Form3.text1.Text = "" Then
       MsgBox "请输入原始密码！"
       Form3.Text2.Text = ""
       Form3.Text3.Text = ""
       Form3.text1.SetFocus
   ElseIf Not Form3.text1.Text = "" = True And Form3.Text2.Text = "" Then
      
       MsgBox "请输入设定密码！"
        Form3.Text3.Text = ""
        Form3.Text2.SetFocus
   ElseIf Not Form3.text1.Text = "" = True And Len(Form3.Text2.Text) <> 6 Then
      MsgBox "请输入六位数密码！"
       Form3.Text2.Text = ""
       Form3.Text3.Text = ""
      Form3.Text2.SetFocus
   ElseIf Form3.Text3.Text = "" Then
        MsgBox "请输入验证密码！"
        Form3.Text3.SetFocus
   ElseIf Not Form3.text1.Text = z = True And Not Form3.text1.Text = "" = True And Not Form3.Text2.Text = "" = True Then
      MsgBox "密码输入错误"
      Form3.text1.Text = ""
      Form3.Text2.Text = ""
       Form3.Text3.Text = ""
       Form3.text1.SetFocus
   ElseIf Form3.text1.Text = z And Not Form3.Text2.Text = Form3.Text3.Text = True Then
       MsgBox "设定密码与验证密码不一致！"
        Form3.Text2.Text = ""
       Form3.Text3.Text = ""
       Form3.Text2.SetFocus
   ElseIf Form3.text1.Text = z And Form3.Text2.Text = Form3.Text3.Text And Len(Form3.Text2.Text) <> 6 Then
       MsgBox "请输入六位数密码！"
        Form3.Text2.Text = ""
       Form3.Text3.Text = ""
   ElseIf Form3.text1.Text = z And Len(Form3.Text2.Text) = 6 Then
        z = Form3.Text2.Text
        MsgBox "密码修改成功！"
       
        Form3.text1.Text = ""
        Form3.Text2.Text = ""
        Form3.Text3.Text = ""
        Form3.Hide
        Form2.Show
       
  
      
  End If
End Sub

 Sub jc()      '检测程序是否被重复调用
     
   If App.PrevInstance = True Then '检视前一版本
    MsgBox "此程式已经在执行中！", 48
    End
    
   End If
 
 End Sub
 Sub dr()     '  将mima.txt读入zz变量
    rs1.Open "select * from 密码", Cnn, adOpenKeyset, adLockOptimistic
    zz = rs1.Fields(1)
    rs1.Update
     rs1.Close
  '  Open App.Path & "\mima.txt" For Input As #1         '打开程序路径下mima.txt
  ' Line Input #1, zz                                  ' 把内容读入变量ZZ中
   'Close #1  ' 关闭文件

 End Sub
Sub xr()       ' 将zz变量写入mima.txt
    rs1.Open "select * from 密码", Cnn, adOpenKeyset, adLockOptimistic
    rs1.Fields(1) = zz
    rs1.Update
    rs1.Close
  
 ' Open App.Path & "\mima.txt" For Output As #1      '建立并打开mima.txt
 ' Print #1, zz                                     'ZZ变量写入mima.txt
 ' Close #1 '关闭mima.txt

 
End Sub
Sub dlmmjc()     '登陆密码检测
   If Form2.text1.Text = z Then
   ' Form1.Show
   ' Form2.Hide
     
     Load main_del
    
     main_del.Show
    ElseIf Form2.text1.Text = "" Then
     msg = MsgBox("请输入密码！", 0, " 提示")
    Form2.text1.SetFocus
   ElseIf Len(Form2.text1.Text) <> 6 Then
   msg = MsgBox("请输六位数字！", 0, " 提示")
    Form2.text1.Text = ""
    Form2.text1.SetFocus
   Else
   
    msg = MsgBox("密码输入错误，请重新输入！ ", 0, "提示"): Form2.text1.Text = ""
   Form2.text1.SetFocus
    X = X + 1
    
  End If
   
  If X > 3 Then Unload Form1: Unload Form3: Unload Form2
End Sub
  
Sub xrsj()    '写入时间

    rs1.Open "select * from 时间", Cnn, adOpenKeyset, adLockOptimistic
    rs1.Fields(1) = mm
    rs1.Fields(2) = ss
    rs1.Update
    rs1.Close
   
   ' Open App.Path & "\m.txt" For Output As #1      '建立并打开m.txt
  'Print #1, mm                                    'mm变量写入mima.txt
 ' Close #1                                          '关闭m.txt
   '  Open App.Path & "\s.txt" For Output As #1      '建立并打开s.txt
 ' Print #1, ss                                   'ss变量写入s.txt
 ' Close #1                                          '关闭s.txt
End Sub
Sub dqsj()  ' 读取时间

    rs1.Open "select * from 时间", Cnn, adOpenKeyset, adLockOptimistic
    mm = rs1.Fields(1)
    ss = rs1.Fields(2)
    rs1.Update
    rs1.Close

   ' Open App.Path & "\m.txt" For Input As #1         '打开程序路径下m .txt
  '  Line Input #1, mm                                  ' 把内容读入变量mm中
  ' Close #1   ' 关闭文件
  ' Open App.Path & "\s.txt" For Input As #1         '打开程序路径下s .txt
  ' Line Input #1, ss                                  ' 把内容读入变量ss中
  ' Close #1  ' 关闭文件
End Sub
 

