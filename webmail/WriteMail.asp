<!--#include file="connsql.ASP"--><head>
<link rel="stylesheet" type="text/css" href="mystyle.css">
<title></title>
</head>
<% dim Action_ID
   dim MailID

   If request("uid")<>"" Then
      MailID = request("uid")
   ElseIf request.Form("Mail_ID")<>"" Then
      MailID=request.Form("Mail_ID")
   Else
      MailID = 0
   End If

   If request("Action_ID")<>"" Then
      Action_ID = request("Action_ID")
   ElseIf Request.Form("Action_ID")<>"" Then
      Action_ID = Request.Form("Action_ID")
   Else
      Action_ID = 0
   End If

'2007/03/12:见上段
'if request("Newmail")="写信" then
'   MailID=0
'   ActionID=0
'elseif request.Form("Re")="回  复" then
'   MailID=request.Form("Mail_ID")
'   Action_ID=1
'elseif request.Form("Alsoto")="转  发" then
'   MailID=request.Form("Mail_ID")
'   Action_ID=2
'end if
'response.write "actionID=[" & Action_ID & "]mainID=[" & MailID & "]" & request("uid")

MYRS.open"select * from webmail where mail_id=" & MailID %>
<form method="POST" action="saveorsendmail.asp" name="form"> 
<body background="images/bluebg.gif">

<table border="0" height="463" bgcolor="#F0F0F0" width="573">
  <tr>
    <td width="565" height="35">
      <table border="0" width="100%" height="27">
        <tr>
          <td width="35%" height="23">
            <p align="center">
            <IMG src="images/a010.gif" name="BSend" id="BSend" width="77" height="31" onclick="BSendClick()" style="cursor:hand">&nbsp;           
            <IMG src="images/a011.gif" name="BSave" id="BSave" width="77" height="31" onclick="BSaveClick()" style="cursor:hand"></td>
          <td width="8%" height="23"></td>
          <td width="11%" height="23"></td>
          <td width="7%" height="23"><input name="ActionID" type="hidden" value="<% =Action_ID %>" ></td>
          <td width="8%" height="23"><input name="MailID" type="hidden" Value="<% =MailID %>" ></td>
          <td width="10%" height="23"></td>
          <td width="13%" height="23"></td>
          <td width="9%" height="23"></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td width="565" height="387" align="left" valign="top">
	
      <table border="0" width="558" height="345" cellspacing="1" bordercolor="#F0F0F0">
        <tr>
          <td width="58" height="18" align="center"><b>发 给</b></td>    
          <td height="18" width="486"><input type="text" name="To" size="71" value= <% Call WriteTo(MailID,Action_ID)%>></td>
        </tr>
        <tr>
         <td width="58" height="18" align="center"><b>抄 送</b></td>
          <td height="18" width="486"><input type="text" name="Alsoto" size="71" value= <% Call WriteAlsoto(MailID,Action_ID)%>></td>
        </tr>
        <tr>
          <td width="58" height="18" align="center"><b>主 题</b></td>
          <td height="18" width="486"><input type="text" name="Subject" size="71" value= "<% Call WriteSubject(MailID,Action_ID)%>"></td>
        </tr>
        <tr>
          <td width="58" height="30" align="center"><b>附 件</b></td>
          <td height="30" width="486">
<input type="button" name="btnFindFile" value="添 加"> 
<input type="text"   name="Attachment"  value="<% Call WriteAttachment(MailID,Action_ID)%>" size=62 readonly>     
		  </td>
        </tr>
        <tr>
          <td width="58" height="280" align="center" valign="top"><b>正 文</b></td>
          <td height="280" valign="top" width="486">
		  <textarea name="content" cols="59" rows="11" ><% Call WriteContent(MailID,Action_ID)%></textarea>
		  <input type="hidden" name="ActionType" value="1">                                                        
		  </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td width="565" height="29">
      <table border="0" width="100%">
        <tr>
          <td width="12%" height="100%">
            <p align="right"></td>
          <td width="12%" height="100%"></td>
          <td width="12%" height="100%"></td>
          <td width="12%" height="100%"></td>
          <td width="13%" height="100%"></td>
          <td width="13%" height="100%"></td>
          <td width="13%" height="100%"></td>
          <td width="13%" height="100%"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>

</form>    
<form name='formfile'>
   <input type="file"   name="file1"       value="" size=1 style="visibility:hidden"> 
</form>


<%    
function  WriteTo(MailID,ActionID)
  if MailID=0 and ActionID=0 then
     response.Write("")
  elseif MailID>0 and ActionID=0 then
     response.Write(MYRS("mail_to"))
  elseif MailID>0 and ActionID=1 then 
      response.Write(MYRS("mail_from")) 
  elseif MailID>0 and ActionID=2 then 
       response.Write(MYRS("mail_to")) 
  end if 
  end function  
%> 
  
<%    
 function WriteAlsoTo(MailID,ActionID)
 if MailID=0 then
     response.Write("")
  elseif ActionID=0 then
     response.Write(MYRS("mail_alsoto")) 
  elseif ActionID=1 then 
      response.Write("") 
  elseif  ActionID=2 then 
       response.Write("") 
  end if 
  end function  
 %> 
 
 
 <%function WriteSubject(MailID,ActionID)    
  if MailID=0 and ActionID=0 then
     response.Write("")
  elseif MailID>0 and ActionID=0 then
     response.Write(MYRS("mail_subject")) 
  elseif MailID>0 and ActionID=1 then 
     response.Write("re:" & MYRS("mail_subject")) 
  elseif MailID>0 and  ActionID=2 then 
     'response.Write(MYRS("mail_subject")) 
     response.Write("fw:" & MYRS("mail_subject")) 
  end if 
  end function  
 %> 
 
 
  <%function WriteAttachment(MailID,ActionID)    
  if MailID=0 then
     response.Write("")
  elseif ActionID=0 then
     response.Write(MYRS("mail_attachment")) 
  elseif ActionID=1 then 
     'response.Write(MYRS("mail_attachment")) 
     response.Write("") 
  elseif  ActionID=2 then 
     response.Write(MYRS("mail_attachment")) 
  end if 
  end function  
 %> 
 
 
<%function WriteContent(MailID,ActionID)    
 if MailID=0 then
     response.Write("")
  elseif ActionID=0 then
     response.Write(MYRS("mail_content")) 
  elseif ActionID=1 then 
      response.Write("") 
  elseif  ActionID=2 then 
       response.Write(MYRS("mail_content")) 
  end if 
  end function  
 %>
 
 
<SCRIPT LANGUAGE=vbscript>            
  Sub BSendClick()
    If Trim(Form.To.Value)="" Then
       Msgbox "请填写邮件收件人", vbInformation+vbOkOnly, "发送"
    ElseIf Trim(Form.Subject.Value)="" Then
       Msgbox "请填写邮件主题", vbInformation+vbOkOnly, "发送"
    Else   
       Form.ActionType.Value = 1
       Form.Submit
    End If   
  End Sub  

  Sub BSaveClick()
    If Trim(Form.To.Value)="" Then
       Msgbox "请填写邮件收件人", vbInformation+vbOkOnly, "发送"
    ElseIf Trim(Form.Subject.Value)="" Then
       Msgbox "请填写邮件主题", vbInformation+vbOkOnly, "发送"
    Else   
       Form.ActionType.Value = 2
       Form.Submit
    End If   
  End Sub  

  Sub btnFindFile_OnClick()
    formfile.file1.Click 
    Form.Attachment.Value = formfile.file1.Value
  End Sub
</SCRIPT>      



























