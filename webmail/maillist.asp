
<!--#include file="DispMailList.ASP"-->

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" type="text/css" href="mystyle.css">
<title></title>
</head>
<body background="images/bluebg.gif">


<form id="form" name="form" method="post" onSubmit="" action="maillist.asp">

<%	
	mailidtodelete = request.form("MailIDToDeleted")
	mailidtomove = request.form("MailIDToMove")
	actionid = request.Form("ActionId")

	if right(mailid,1)="," then
	   mailid=left(mailid,len(mailid)-1)
	end if

	if request("id")=1  or request("id") =2 or request("id")=3 or request("id")=4 then
	   session("MailType") = request("id")
	   session("MailBoxToMoveTo") = request("id")
	end if   

	MYRS.OPEN "select * from webmail WHERE mail_id IN ("& mailid &")"	
    if  actionid = 1  then
        if mailidtodelete<>"" then  'ɾ���ʼ�
				if  session("MailType") =4  then
			     	MYRS.CLOSE
			    	mysql = "DELETE  FROM webmail WHERE mail_id IN ("& mailidtodelete &")"
				    MYcn.EXECUTE mysql
				    MYRS.MoveFirst
    				Showdata
	 			else 
	  	
	    			MYRS.CLOSE
		    		MYcn.EXECUTE "UPDATE webmail set mail_type= 4 WHERE mail_id IN (" & mailidtodelete & ")"
					
				end if
		end if	
	end if

	if  actionid = 1  then
			if mailidtomove<> "" then
				MYRS.CLOSE
				MYcn.EXECUTE "UPDATE webmail set mail_type=" & Request.Form("FolderMoveTo") &  " WHERE mail_id IN (" & mailidtomove & ")"
		    end if
	end if   
%>
    
	<input name="MailIDToDeleted" type="hidden" value="" />
	<input name="MailIDToMove" type="hidden" value="" />
	<input name="MailBoxToMoveTo" type="hidden" value="" />
	<input name="ActionId" type="hidden" value="2" />
	
<table border="0" width="100%" height="172">
  <tr>
    <td width="10%" height="30">
        
    
    
    </td>
  </tr>
  <tr>
    <td width="36%" height="30" colspan="7"> <% CALL DispMailTitle(session("MailType")) %></td>
    <!--<td width="18%" height="30" colspan="2">(��������)</td> -->
    <td width="21%" height="30" colspan="2">&nbsp;</td>
  </tr>
  <tr>
    <td width="100%" height="76" colspan="13" valign="top">
      <table width="100%" height="55" border="0" cellspacing="0" style="border-top-style: solid; border-bottom-style: solid" cellpadding="0">
        <tr>
          <td width="23%" height="30" bordercolor="#589cdb" bgcolor="#589cdb"><div align="center"><font size="2"><b> ������</b></font></div></td>
          <td width="47%" height="30" bordercolor="#589cdb" bgcolor="#589cdb"><div align="center"><font size="2"><b>����</b></font></div></td>
          <td width="16%" height="30" bordercolor="#589cdb" bgcolor="#589cdb"><font size="2"><b>����</b></font></td>
        </tr>
        
        <tr>
          <td height="25" colspan="3" bordercolor="#589cdb" bgcolor="#F0F0F0"><% 
		  IF session("MailType") <> "" THEN

		  	CALL DispMailList(session("MailType")) 
		 END IF
			%>
			</td>
        </tr>
       
        <tr>
          <td width="67%" height="25" bordercolor="#589cdb" bgcolor="#589cdb" colspan="2">&nbsp;�ƶ�����<select size="1" name="FolderMoveTo">
<% If session("MailType")=1 Then %>          
              <option value="2">�ݸ���</option>
              <option value="3">�ѷ���</option>
              <option value="4">��ɾ��</option>
<% ElseIf session("MailType")=2 Then %>          
              <option value="1">�ռ���</option>
              <option value="3">�ѷ���</option>
              <option value="4">��ɾ��</option>
<% ElseIf session("MailType")=3 Then %>          
              <option value="1">�ռ���</option>
              <option value="2">�ݸ���</option>
              <option value="4">��ɾ��</option>
<% ElseIf session("MailType")=4 Then %>          
              <option value="1">�ռ���</option>
              <option value="2">�ݸ���</option>
              <option value="3">�ѷ���</option>
<% End If %>
            </select></td>
          <td width="16%" height="25" bordercolor="#589cdb" bgcolor="#589cdb"><% 
		   IF	session("MailType") <> "" THEN
		  CALL DispMailButton(session("MailType"))

		  	END IF
		   %></td>
        </tr>

      </table>    </td>
  </tr>
  <tr>
    <td width="100%" height="14" colspan="13">
      <p align="center">
      </td>
  </tr>

  
</form>
</table>

<SCRIPT LANGUAGE=VBscript>                                                           
<!-- 
  function selectchk()
     On Error Resume Next
     str2 = ""
     form.MailIDToDeleted.value= ""
     len1 = document.form.elements.length -1 
     for i=0 to len1
        lcName = document.form.elements(i).name
        if IsNumeric(Right(lcName,len(lcName)-1)) Then
           if document.form.elements(i).checked= true then
              str = Right(lcName,len(lcName)-1)
              str2 = str &"," + str2
              form.MailIDToDeleted.value= str2
           end if   
        end if
     next
     if str2="" then
        selectchk = false
     else
        form.MailIDToDeleted.value= Left(str2, Len(str2)-1)
        selectchk = true
     end if
  end function

  sub cmddel_onclick()    
     if not selectchk() then
        msgbox "������ѡ��Ҫɾ�����ʼ�", vbExclamation +vbOkOnly, "�ʼ�" 
        exit sub
     end if
     If msgbox("���Ҫɾ����?", vbQuestion+vbOKCancel, "ȷ��")=vbOK Then
        form.ActionId.value= 1
     else
        form.ActionId.value= 2
     end if
  end sub	
	
  sub cmdreh_onclick()
     Location.Reload
  end sub
	
  sub cmdmov_onclick()
     if not selectchk() then
        msgbox "������ѡ��Ҫ�ƶ����ʼ�", vbExclamation +vbOkOnly, "�ʼ�" 
        exit sub
     end if
     If msgbox("���Ҫ�ƶ���?", vbQuestion+vbOKCancel, "ȷ��")=vbOK Then
        form.MailIDToMove.value= form.MailIDToDeleted.value
        form.MailIDToDeleted.value= ""
        form.ActionId.value= 1
     else
        Location.Reload
        form.ActionId.value= 2
     end if
  end sub

  function form_onsubmit()
     If form.ActionId.value= 1 Then
        form_onsubmit = true
     Else
        form_onsubmit = false
     End if   
  end function
  
  sub cmdbak_onclick()
     history.back
  end sub
!-->                                                           
</SCRIPT>  
</body>
</html>
