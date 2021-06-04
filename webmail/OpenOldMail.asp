<!--#include file="connsql.ASP"-->	
<head>
<link rel="stylesheet" type="text/css" href="mystyle.css">
<title></title>
</head>
<% dim MailID
   MailID = request("uid")
   sql  = "select * from webmail where mail_id="& MailID
   MYRS.Open sql,MYcn
   MYRS("mail_isnew") = 0
   MYRS.Update
%>
<body background="images/bluebg.gif">

<table border="0" width="87%" cellspacing="0" cellpadding="0" height="579" bgcolor="#F0F0F0">
  <tr>
    <td width="100%" height="1" colspan="8">
    </td>
  </tr>
  <tr>

    <td width="3%" height="31" bgcolor="#F0F0F0">
        <form method="POST" action="WriteMail.asp" >
         <p align="right">
         <input alt="submit" type="image" value="回  复" src="images/a008.gif" width=77 width=31 name="Re" >&nbsp;   
         <input name="Action_ID" type="hidden" value="1">    
         <input name="Mail_ID"   type="hidden" value=<%=MYRS("mail_id")%>>     
         </form>   </td>  
  
    <td width="11%" height="31" bgcolor="#F0F0F0">  
        <form method="POST" action="WriteMail.asp" >  
         <p align="center"> 
         <input alt="submit" type="image" value="转  发" src="images/a009.gif" width=77 width=31 name="Alsoto" > 
         <input name="Action_ID" type="hidden" value="2">     
         <input name="Mail_ID"   type="hidden" value=<%=MYRS("mail_id")%>>   
         </form>   </td>  
  
    <td width="10%" height="31" colspan="2" bgcolor="#F0F0F0">  
    </td>  
  
    <td width="66%" height="31" colspan="4" bgcolor="#F0F0F0">  
    </td>  
  </tr>  
  <tr>  
    <td width="112%" height="3" colspan="8" bgcolor="#F0F0F0">  
    </td>  
  </tr>  
  <tr>  
    <td width="100%" height="519" valign="top" colspan="8" bgcolor="#F0F0F0">  
      <table border="0" width="99%" height="391" cellpadding="2" bgcolor="#F0F0F0">  
        <tr>  
          <td width="14%" height="22" align="center" valign="middle" bordercolor="#F0F0F0" bgcolor="#F0F0F0"><b><font size="2">日&nbsp;              
            期：</font></b></td>  
          <td width="96%" height="22" bordercolor="#C0C0C0" bgcolor="#FFFFFF"><%=MYRS("mail_datetime")%></td>  
        </tr>  
        <tr>  
          <td width="5%" height="22" align="center" valign="middle" bordercolor="#F0F0F0" bgcolor="#F0F0F0"><b><font size="2">发件人：</font></b></td>  
          <td width="96%" height="22" bordercolor="#C0C0C0" bgcolor="#FFFFFF"><%=MYRS("mail_from")%></td>  
        </tr>  
        <tr>  
          <td width="5%" height="22" align="center" valign="middle" bordercolor="#F0F0F0" bgcolor="#F0F0F0"><b><font size="2">收件人：</font></b></td>  
          <td width="96%" height="22" bordercolor="#C0C0C0" bgcolor="#FFFFFF"><%=MYRS("mail_to")%></td>  
        </tr>  
        <tr>  
          <td width="5%" height="22" align="center" valign="middle" bordercolor="#F0F0F0" bgcolor="#F0F0F0"><b><font size="2">主&nbsp;              
            题：</font></b></td>  
          <td width="96%" height="22" bordercolor="#C0C0C0" bgcolor="#FFFFFF"><%=MYRS("mail_subject")%></td>  
        </tr>  
        <tr>  
          <td width="5%" height="22" align="center" valign="middle" bordercolor="#F0F0F0" bgcolor="#F0F0F0"><b><font size="2">附&nbsp;              
            件：</font></b></td>  
          <td width="96%" height="22" bordercolor="#C0C0C0" bgcolor="#FFFFFF"><a href="attachments/<%=MYRS("mail_attachment")%>"> <%=MYRS("mail_attachment")%></a></td>   
        </tr>  
        <tr>  
          <td width="5%" align="center" height="264" valign="top" bgcolor="#F0F0F0"><b><font size="2">内&nbsp;              
            容：</font></b></td>  
          <td width="95%" height="264" valign="top" bgcolor="#FFFFFF"><%=MYRS("mail_content")%></td>  
        </tr>  
  
      </table>  
    </td>  
  </tr>  
  <tr>  
    <td width="13%" height="27">  
      <p align="center"></td>  
    <td width="7%" height="27">  
      <p align="center">  
      </p>  
    </td>  
    <td width="11%" height="27">  
      <p align="center">&nbsp;</p></td>  
    <td width="1%" height="27">  
      <p align="center">&nbsp;</p></td>  
    <td width="44%" height="27"></td>  
    <td width="12%" height="27"></td>  
    <td width="12%" height="27"></td>  
    <td width="12%" height="27"></td>  
  </tr>  
</table>  
<input name="Mail_ID" type="hidden" value=<%=MYRS("mail_id")%>>  
  
  
