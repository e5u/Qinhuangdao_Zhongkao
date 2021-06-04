<!--#include file="connsql.ASP"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>电子邮件系统</title>
</head>

<body>
<%
         '2006/07/11:修改wmLogChk.asp使其能集成到考试系统中，对考试系统作检查
         If Trim(Request("passwd"))="" Then
             Response.Write "<DIV ALIGN=CENTER>" 
             Response.Write "<p>&nbsp;</p>"
             Response.Write "<h3>请在密码栏输入题号!</h3>"
             Response.Write "<p>&nbsp;</p>"
             Response.Write "<p><a href='wmLogin.asp'>"
             Response.Write "返回</a></p>"
             Response.Write "</DIV>"
             Response.END          
         Else
            Response.Write "<h3>题号!</h3>" & trim(Request("passwd"))
            session("kaohao") = trim(Request("kaohao"))
            session("passwd") = trim(Request("passwd"))

         End If
            
         DIM stuEnabled 
          
          '检查考生是否存在 
'          lcSql = "select * from student where stu_no = '" & session("kaohao") & "'"  
'          MYRS.Open lcSql,MYcn
'          If MYRS.Eof = True Then
'             Response.Write "<DIV ALIGN=CENTER>" 
'             Response.Write "<p>&nbsp;</p>"
'             Response.Write "<h3>找不到考生号是" & session("kaohao") & "的考生!</h3>"
'             Response.Write "<p>&nbsp;</p>"
'             Response.Write "<p><a href='wmLogin.asp'>"
'             Response.Write "返回</a></p>"
'             Response.Write "</DIV>"
'             Response.END
'          Else
'             strEnabled = MYRS("stu_enabled")
'          End If
'          MYRS.Close
          
          '检查考生状态是否是正在考试
'         If strEnabled = 0 Then
'             Response.Write "<DIV ALIGN=CENTER>" 
'             Response.Write "<p>&nbsp;</p>"
'             Response.Write "<h3>考试尚未开始或已经结束!</h3>"
'             Response.Write "<p>&nbsp;</p>"
'             Response.Write "<p><a href='wmLogin.asp'>"
'             Response.Write "返回</a></p>"
'             Response.Write "</DIV>"
'             Response.END             
'          End If

          '检查考题的试题类型是否是webmail试题 
          lcSql = "select * from question where qst_no = '" & session("passwd") & "'"  
          MYRS.Open lcSql,MYcn
          If MYRS.Eof = True Then
             Response.Write "<DIV ALIGN=CENTER>" 
             Response.Write "<p>&nbsp;</p>"
             Response.Write "<h3>试题" & session("passwd") & "不存在!</h3>"
             Response.Write "<p>&nbsp;</p>"
             Response.Write "<p><a href='wmLogin.asp'>"
             Response.Write "返回</a></p>"
             Response.Write "</DIV>"
             Response.END
          Else
             If MYRS("qst_type") <> 35 Then
                Response.Write "<DIV ALIGN=CENTER>" 
                Response.Write "<p>&nbsp;</p>"
                Response.Write "<h3>试题类型不正确!</h3>"
                Response.Write "<p>&nbsp;</p>"
                Response.Write "<p><a href='wmLogin.asp'>"
                Response.Write "返回</a></p>"
                Response.Write "</DIV>"
                Response.END       
             End If
          End If
          MYRS.Close

          '检查该考生下发试卷中对应的考试试题是否存在
'          lcSql = "select * from webmail where qst_no='" & session("passwd") & "'"
'          MYRS.Open lcSql,MYcn
'          If MYRS.Eof = True Then
'             Response.Write "<DIV ALIGN=CENTER>" 
'             Response.Write "<p>&nbsp;</p>"
'             Response.Write "<h3>考号" & session("kaohao") & "的试卷中没有题号是" & session("passwd") & "的考题!</h3>"
'             Response.Write "<p>&nbsp;</p>"
'             Response.Write "<p><a href='wmLogin.asp'>"
'             Response.Write "返回</a></p>"
'             Response.Write "</DIV>"
'             Response.END
'          End If
'          MYRS.Close
          Response.Redirect("wmMain.asp")
          
%>

</body>
</html>