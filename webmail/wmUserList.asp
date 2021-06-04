<!--#include file="connsql.ASP"-->
<html>
<head>
<title>168电子邮件系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/css.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="1027" background="images/mail2_02.gif"><img src="images/mail2_01.gif" width="748" height="55" alt=""></td>
    <td width="13"><img src="images/mail2_03.gif" width="13" height="55" alt=""></td>
  </tr>
</table>

<table width="100%"  border="0" cellspacing="0" cellpadding="0" height="516">
  <tr>
    <td width="17" height="516"></td>
    <td width="176" valign="top" background="images/mail2_05.gif" height="516">
        　
    </td>
    <td width="9" bgcolor="#91C8F1" height="516">　</td>
    <td width="824" valign="top" height="516">
<%
  'Set MyRS = SERVER.CREATEOBJECT("ADODB.RECORDSET")
  'MyRS.ACTIVECONNECTION = MYCN
  'MyRS.CURSORLOCATION   = adUseServer
  'MyRS.CURSORTYPE       = adOpenKeySet
  'MyRS.LOCKTYPE         = adLockOptimistic  'adLockBatchOptimistic 'adLockOptimistic  'adLockBatchOptimistic  '  

  MyRS.Open "SELECT * FROM WEBMAILUSER WHERE STU_NO='" & session("Stu_No") & "' and qst_no='" & session("qst_No") & "' "
  Response.write "<p>&nbsp;</p>"
  If MyRs.Eof Then
     Response.write "&nbsp;你已申请的邮箱数：0"
  Else
     Response.write "&nbsp;你已申请的邮箱数：" & MyRS.RecordCount
  End If   
  Response.write "<hr>"
  Do While Not MyRS.Eof 
     Response.Write "<br>&nbsp;帐号："  & MyRS.Fields("user_no")
     Response.Write "<br>&nbsp;用户名："  & MyRS.Fields("user_name")
     Response.Write "<br>&nbsp;密码："  & MyRS.Fields("user_pwd")
     Response.Write "<br>&nbsp;提示问题："  & MyRS.Fields("user_ask")
     Response.Write "<br>&nbsp;提示答案："  & MyRS.Fields("user_asw")
     Response.Write "<hr>"
     
     MyRS.MoveNext
  Loop   
  MyRS.Close
  MyCn.Close
  Set MyCn = Nothing
  Set MyRS = Nothing
%>
  <p align="center"><a href="wmReg.htm">申请邮箱</a>&nbsp;<a href="wmLogin2.asp">返回</a></p>   
          </TD>
        </TR>
</TABLE>
   <input name="frmLocalIP" type="hidden" value="" />
   <input name="frmWsNo" type="hidden" value="" />
   <input name="frmTesterNo" type="hidden" value="" />        
   <input name="kaohao" id="kaohao"  type="hidden" value="210601165001" />
<table width="100%" height="16"  border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="24" height="13" background="images/mail2_201.gif"></td>
    <td width="991" height="13" background="images/mail2_202.gif"></td>
    <td width="25" height="13" background="images/mail2_203.gif"></td>
  </tr>
</table>
</body>
</html>

