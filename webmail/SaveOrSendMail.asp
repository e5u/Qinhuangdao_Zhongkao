<!--#include file="connsql.ASP"-->
<%
Err.Clear
On Error Resume Next

'2007/03/12.if request.Form("BSave")="存草稿" then
If request.Form("ActionType")="2" then  
'response.write "err2=" & err.description & "action=[" & request.Form("ActionID") & "]"

   If request.Form("MailID")=0 Or request.Form("ActionID")>0 Then
      MYRS.open"select * from webmail where 1=2"
      MYRS.addNew
   Else
      MYRS.open"select * from webmail where mail_id=" & request.Form("MailID")
   End If  
   MYRS("mail_from")="kaoshi@webmail.com"
   MYRS("mail_to")=request.Form("To")
   MYRS("mail_alsoto")=request.Form("Alsoto")
   MYRS("mail_subject")=request.Form("Subject")
   MYRS("mail_attachment")=request.form("Attachment")
   MYRS("mail_content")=request.Form("Content")
   MYRS("mail_type")=2
   MYRS("mail_datetime")="" & now() 
   MYRS("stu_no")=session("kaohao")
   MYRS("qst_no")=session("passwd")
   If Session("Exm_No")<> "" Then
      MYRS("exm_no") = Session("Exm_No")
   End If
   MYRS.update
   MYRS.close

   If Err.Number=0 Then
      message = "成功将信存入草稿箱！"
   Else
      message = Err.Description
   End If
'2007/03/12.elseif request.Form("BSend")="发  送" then
ElseIf request.Form("ActionType")="1" then
   If request.form("to")<>"" then
      If request.Form("MailID")=0 Or request.Form("ActionID")>0 Then
         MYRS.open"select * from webmail where 1=2"
         MYRS.addNew
      Else
         MYRS.open"select * from webmail where mail_id=" & request.Form("MailID")
      End If  
      MYRS("mail_from")="kaoshi@webmail.com"
      MYRS("mail_to")=request.Form("To")
      MYRS("mail_alsoto")=request.Form("Alsoto")
      MYRS("mail_subject")=request.Form("Subject")
      MYRS("mail_attachment")=request.form("Attachment")
      MYRS("mail_content")=request.Form("Content")
      MYRS("mail_type")=3
      MYRS("mail_datetime")="" & now() 
      MYRS("stu_no")=session("kaohao")
      MYRS("qst_no")=session("passwd")
      If Session("Exm_No")<> "" Then
         MYRS("exm_no") = Session("Exm_No")
      End If
      MYRS.update
      MYRS.close

      If Err.Number=0 Then
         message = "邮件发送成功！"
      Else
         message = Err.Description
      End If
   End If
End If
%>

<html>

<head>

</head>

<body>
<form method="POST" action="Maillist.asp" >  
<table border="0" width="100%" height="420">
  <tr>
    <td width="100%" height="416">
      <table border="0" width="100%" height="147" cellspacing="0" cellpadding="0">
        <tr>
          <td width="22%" height="33" align="center"></td>
          <td width="44%" height="33" align="center" bgcolor="#589cdb">
            <p align="left"><b><font size="2" color="#FFFFFF">[系统提示]</font></b></td>
          <td width="34%" height="33" align="center"></td>
        </tr>
        <tr>
          <td width="22%" height="71" align="center"></td>
          <td width="44%" height="71" align="center" bgcolor="#589cdb"><b><font size="2" color="#FFFFFF"><% response.Write message %></font></b></td>
          <td width="34%" height="71" align="center"></td>
        </tr>
        <tr>
          <td width="22%" height="35" align="center"></td>
          <td width="44%" height="35" align="center" bgcolor="#589cdb"><b><a href="MailList.asp?id=1">返回</a></b></td>
          <td width="34%" height="35" align="center"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>

</form>
</body>

</html>

