<!--#include file="connsql.ASP"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�����ʼ�ϵͳ</title>
</head>

<body>
<%
         '2006/07/11:�޸�wmLogChk.aspʹ���ܼ��ɵ�����ϵͳ�У��Կ���ϵͳ�����
         If Trim(Request("passwd"))="" Then
             Response.Write "<DIV ALIGN=CENTER>" 
             Response.Write "<p>&nbsp;</p>"
             Response.Write "<h3>�����������������!</h3>"
             Response.Write "<p>&nbsp;</p>"
             Response.Write "<p><a href='wmLogin.asp'>"
             Response.Write "����</a></p>"
             Response.Write "</DIV>"
             Response.END          
         Else
            Response.Write "<h3>���!</h3>" & trim(Request("passwd"))
            session("kaohao") = trim(Request("kaohao"))
            session("passwd") = trim(Request("passwd"))

         End If
            
         DIM stuEnabled 
          
          '��鿼���Ƿ���� 
'          lcSql = "select * from student where stu_no = '" & session("kaohao") & "'"  
'          MYRS.Open lcSql,MYcn
'          If MYRS.Eof = True Then
'             Response.Write "<DIV ALIGN=CENTER>" 
'             Response.Write "<p>&nbsp;</p>"
'             Response.Write "<h3>�Ҳ�����������" & session("kaohao") & "�Ŀ���!</h3>"
'             Response.Write "<p>&nbsp;</p>"
'             Response.Write "<p><a href='wmLogin.asp'>"
'             Response.Write "����</a></p>"
'             Response.Write "</DIV>"
'             Response.END
'          Else
'             strEnabled = MYRS("stu_enabled")
'          End If
'          MYRS.Close
          
          '��鿼��״̬�Ƿ������ڿ���
'         If strEnabled = 0 Then
'             Response.Write "<DIV ALIGN=CENTER>" 
'             Response.Write "<p>&nbsp;</p>"
'             Response.Write "<h3>������δ��ʼ���Ѿ�����!</h3>"
'             Response.Write "<p>&nbsp;</p>"
'             Response.Write "<p><a href='wmLogin.asp'>"
'             Response.Write "����</a></p>"
'             Response.Write "</DIV>"
'             Response.END             
'          End If

          '��鿼������������Ƿ���webmail���� 
          lcSql = "select * from question where qst_no = '" & session("passwd") & "'"  
          MYRS.Open lcSql,MYcn
          If MYRS.Eof = True Then
             Response.Write "<DIV ALIGN=CENTER>" 
             Response.Write "<p>&nbsp;</p>"
             Response.Write "<h3>����" & session("passwd") & "������!</h3>"
             Response.Write "<p>&nbsp;</p>"
             Response.Write "<p><a href='wmLogin.asp'>"
             Response.Write "����</a></p>"
             Response.Write "</DIV>"
             Response.END
          Else
             If MYRS("qst_type") <> 35 Then
                Response.Write "<DIV ALIGN=CENTER>" 
                Response.Write "<p>&nbsp;</p>"
                Response.Write "<h3>�������Ͳ���ȷ!</h3>"
                Response.Write "<p>&nbsp;</p>"
                Response.Write "<p><a href='wmLogin.asp'>"
                Response.Write "����</a></p>"
                Response.Write "</DIV>"
                Response.END       
             End If
          End If
          MYRS.Close

          '���ÿ����·��Ծ��ж�Ӧ�Ŀ��������Ƿ����
'          lcSql = "select * from webmail where qst_no='" & session("passwd") & "'"
'          MYRS.Open lcSql,MYcn
'          If MYRS.Eof = True Then
'             Response.Write "<DIV ALIGN=CENTER>" 
'             Response.Write "<p>&nbsp;</p>"
'             Response.Write "<h3>����" & session("kaohao") & "���Ծ���û�������" & session("passwd") & "�Ŀ���!</h3>"
'             Response.Write "<p>&nbsp;</p>"
'             Response.Write "<p><a href='wmLogin.asp'>"
'             Response.Write "����</a></p>"
'             Response.Write "</DIV>"
'             Response.END
'          End If
'          MYRS.Close
          Response.Redirect("wmMain.asp")
          
%>

</body>
</html>