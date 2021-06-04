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
    <td width="824" valign="top" height="516"><%
  Response.write "<p>&nbsp;</p>"
  'Set MyRS = SERVER.CREATEOBJECT("ADODB.RECORDSET")
  'MyRS.ACTIVECONNECTION = MYCN
  'MyRS.CURSORLOCATION   = adUseServer
  'MyRS.CURSORTYPE       = adOpenKeySet
  'MyRS.LOCKTYPE         = adLockOptimistic  'adLockBatchOptimistic 'adLockOptimistic  'adLockBatchOptimistic  '  

  MyRS.Open "SELECT * FROM WEBMAILUSER WHERE STU_NO='" & session("Stu_No") & "' and qst_no='" & session("qst_No") & "' and user_no='" & Request.form("user_no") & "'"
  If MyRS.Eof Then
     MyRS.Addnew
  End If

  MyRS.Fields("stu_no")    = session("stu_no")     
  MyRS.Fields("qst_no")    = session("qst_No")     
  MyRS.Fields("user_no")   = Request.form("user_no")
  MyRS.Fields("user_pwd")  = Request.form("user_pwd")
  MyRS.Fields("user_name") = Request.form("user_name")
  MyRS.Fields("user_ask")  = Request.form("user_ask")
  MyRS.Fields("user_asw")  = Request.form("user_asw")
      
  '2007/03/09:如果Session("Exm_No")不空，则保存
  If Session("Exm_No")<> "" Then
     MyRS.Fields("Exm_No")  = Session("Exm_No")
  End If
      
  MyRS.Update
  If Err.Number=0 Then
     Response.write "<h3 align=center>邮箱注册成功!</h>"
     Response.write "<p>&nbsp;</p>"
     Response.write "<a href='wmUserList.asp'>返回</a>"
  Else
     Response.write "<h3 align=center>注册信息保存出错：" & Err.Description
     Response.write "<p>&nbsp;</p>"
     Response.write "<input type=button name='b1' value='返回' onclick='history.back()'>"
  End If    
  
  MyRS.Close
  MyCn.Close
  Set MyCn = Nothing
  Set MyRS = Nothing
%>
  <p align="center">　</p>   
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

