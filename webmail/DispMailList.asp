<!--#include file="connsql.ASP"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ޱ����ĵ�</title>
</head>
<body>
<%	
	function DispMailButton(pnMailType)
		'pnmailType=request("id")
		SELECT CASE pnMailType
			CASE 1
				'Response.WRITE "<INPUT type=submit name='cmdbak'    id='cmdbak'    value='����'>"
				Response.WRITE "<INPUT type=submit name='cmddel'    id='cmddel'    value='ɾ��'>"
				Response.WRITE "<INPUT type=submit name='select'    id='cmdmov'    value='�ƶ�'>"
				Response.WRITE "<INPUT type=submit name='cmdreh'    id='cmdreh'    value='ˢ��'>"
			CASE 2
				'Response.WRITE "<INPUT type=submit name='cmdbak'    id='cmdbak'    value='����'>"
				Response.WRITE "<INPUT type=submit name='cmddel'    id='cmddel'    value='ɾ��'>"
				Response.WRITE "<INPUT type=submit name='select'    id='cmdmov'    value='�ƶ�'>"
				Response.WRITE "<INPUT type=submit name='cmdreh'    id='cmdreh'    value='ˢ��'>"
			CASE 3
				'Response.WRITE "<INPUT type=submit name='cmdbak'    id='cmdbak'    value='����'>"
				Response.WRITE "<INPUT type=submit name='cmddel'    id='cmddel'    value='ɾ��'>"
				Response.WRITE "<INPUT type=submit name='select'    id='cmdmov'    value='�ƶ�'>"
				Response.WRITE "<INPUT type=submit name='cmdreh'    id='cmdreh'    value='ˢ��'>"
			CASE 4
				'Response.WRITE "<INPUT type=submit name='cmdbak'    id='cmdbak'    value='����'>"
				Response.WRITE "<INPUT type=submit name='cmddel'    id='cmddel'    value='ɾ��'>"
				Response.WRITE "<INPUT type=submit name='select'    id='cmdmov'    value='�ƶ�'>"
				Response.WRITE "<INPUT type=submit name='cmdreh'    id='cmdreh'    value='ˢ��'>"
		END SELECT
	end function

	function DispMailTitle(pnMailType)
		'pnmailType=request("id")
		dim uid,qid
		uid=session("kaohao")
		qid=session("passwd")	
		SELECT CASE(pnMailType)
			CASE 1
				MYRS.Open "select * from webmail where mail_Type=1 and stu_no='"& session("kaohao") &"' and qst_no='"& session("passwd") & "'"

				Response.WRITE "�ռ���: �ʼ�����Ϊ" & MYRS.RecordCount
				MYRS.Close
				
				MYRS.Open "select * from webmail where mail_isnew=1 and mail_Type=1 and stu_no='"& session("kaohao") &"' and qst_no='"& session("passwd") & "'"

				Response.WRITE "&nbsp;&nbsp;&nbsp;&nbsp;δ���ʼ�:" & MYRS.RecordCount
				MYRS.Close
				
			CASE 2
				MYRS.Open "select * from webmail where mail_Type=2 and stu_no='"& session("kaohao") &"' and qst_no='"& session("passwd") & "'"

				Response.WRITE "�ݸ�: �ʼ�����Ϊ" & MYRS.RecordCount
				MYRS.Close
				
				
				
			CASE 3
				MYRS.Open "select * from webmail where mail_Type=3 and stu_no='"& session("kaohao") &"' and qst_no='"& session("passwd") & "'"

				Response.WRITE "�ѷ���: �ʼ�����Ϊ" & MYRS.RecordCount
				MYRS.Close
				
				
				
				
			CASE 4
				MYRS.Open "select * from webmail where mail_Type=4 and stu_no='"& session("kaohao") &"' and qst_no='"& session("passwd") & "'"

				Response.WRITE "��ɾ��: �ʼ�����Ϊ" & MYRS.RecordCount
				MYRS.Close
			
				
				
				
	END SELECT
	end function

	function DispMailList(pnMailType)
		dim m_id
	        dim uid,qid
		uid=session("kaohao")
		qid=session("passwd")
		'pnmailType=request("id")
		SELECT CASE pnMailType
			CASE 1
				
				MYRS.Open "select * from webmail where mail_Type=1 and stu_no='"& session("kaohao") &"' and qst_no='"& session("passwd") & "'"
			%>


<table width="100%"  border="0" cellspacing="0" cellpadding="0" style="border-top-style: dotted; border-bottom-style: dotted; padding-right: 0; padding-top: 1; padding-bottom: 1">
			<%DO UNTIL MYRS.EOF%><tr>
		
		<td width="3%" height="25" bgcolor="#F0F0F0"><input type="checkbox" name="C<% =MYRS("mail_id") %> " value="checkbox"></td>
        
          <td width="20%" height="25" bgcolor="#F0F0F0"><% Response.WRITE MYRS("mail_from")%></td>
          <td width="50%" height="25" bgcolor="#F0F0F0">
<% 
 lcAttachment = Trim(MYRS("mail_attachment"))
 If Not IsNull(lcAttachment) And lcAttachment<>"" Then
    response.write "<img src='images/attachment.gif'>" 
 Else
    response.write "&nbsp;" 
 End If 
%>
<% response.Write("<a href='OpenOldMail.asp?uid=" & MYRS("mail_id") & "'>" & MYRS("mail_subject") & "</a>")%></td>
          <td width="16%" height="25" bgcolor="#F0F0F0"><% Response.WRITE MYRS("mail_datetime")%></td>
       

		  </tr><%
					'Response.WRITE MYRS("mail_from")
					m_id		   = MYRS.fields("mail_ID")
					session("uid") = MYRS.fields("mail_ID")
					'response.Write("<a href='maillist.asp?uid=m_id'>" & MYRS("mail_subject") & "</a>")
					'Response.WRITE MYRS("mail_datetime")
					midd 		   = MYRS.fields("mail_ID")
					session("uid") = MYRS.fields("mail_ID")
					MYRS.MoveNext
					loop
				MYRS.Close	%>
				</tr>
</table>			
		<%	CASE 2
				MYRS.Open "select * from webmail where mail_Type=2 and stu_no='"& session("kaohao") &"' and qst_no='"& session("passwd") & "'"
								%>


<table  width="100%"  border="0" cellspacing="0" cellpadding="0" style="border-top-style: dotted; border-bottom-style: dotted">
			<%DO UNTIL MYRS.EOF%><tr>
					
		<td width="3%" height="25" bgcolor="#F0F0F0"><input type="checkbox" name="C<% =MYRS("mail_id") %> " value="checkbox"></td>
         
          <td width="20%" height="25" bgcolor="#F0F0F0"><% Response.WRITE MYRS("mail_from")%></td>
          <td width="50%" height="25" bgcolor="#F0F0F0">
<% 
 lcAttachment = Trim(MYRS("mail_attachment"))
 If Not IsNull(lcAttachment) And lcAttachment<>"" Then
    response.write "<img src='images/attachment.gif'>" 
 Else
    response.write "&nbsp;" 
 End If 
%>
<% response.Write("<a href='WriteMail.asp?uid=" & MYRS("mail_id") & "'>" & MYRS("mail_subject") & "</a>")%></td>
          <td width="16%" height="25" bgcolor="#F0F0F0"><% Response.WRITE MYRS("mail_datetime")%></td>
          
	
		  </tr><%
					'Response.WRITE MYRS("mail_from")
					m_id		   = MYRS.fields("mail_ID")
					session("uid") = MYRS.fields("mail_ID")
					'response.Write("<a href='maillist.asp?uid=m_id'>" & MYRS("mail_subject") & "</a>")
					'Response.WRITE MYRS("mail_datetime")
					midd 		   = MYRS.fields("mail_ID")
					session("uid") = MYRS.fields("mail_ID")
					MYRS.MoveNext
					loop
				MYRS.Close	%>
				</tr>
</table>	
		<%	CASE 3
				
				MYRS.Open "select * from webmail where mail_Type=3 and stu_no='"& session("kaohao") &"' and qst_no='"& session("passwd") & "'"
									%>


<table  width="100%" border="0" cellspacing="0" cellpadding="0" style="border-top-style: dotted; border-bottom-style: dotted">
			<%DO UNTIL MYRS.EOF%><tr>
				
		<td width="3%" height="25" bgcolor="#F0F0F0"><input type="checkbox" name="C<% =MYRS("mail_id") %> " value="checkbox"></td>
         
          <td width="20%" height="25" bgcolor="#F0F0F0"><% Response.WRITE MYRS("mail_from")%></td>
          <td width="50%" height="25" bgcolor="#F0F0F0">
<% 
 lcAttachment = Trim(MYRS("mail_attachment"))
 If Not IsNull(lcAttachment) And lcAttachment<>"" Then
    response.write "<img src='images/attachment.gif'>" 
 Else
    response.write "&nbsp;" 
 End If 
%>
<% response.Write("<a href='OpenOldMail.asp?uid=" & MYRS("mail_id") & "'>" & MYRS("mail_subject") & "</a>")%></td>
          <td width="16%" height="25" bgcolor="#F0F0F0"><% Response.WRITE MYRS("mail_datetime")%></td>
        
	
		  </tr><%
					'Response.WRITE MYRS("mail_from")
					m_id		   = MYRS.fields("mail_ID")
					session("uid") = MYRS.fields("mail_ID")
					'response.Write("<a href='maillist.asp?uid=m_id'>" & MYRS("mail_subject") & "</a>")
					'Response.WRITE MYRS("mail_datetime")
					midd 		   = MYRS.fields("mail_ID")
					session("uid") = MYRS.fields("mail_ID")
					MYRS.MoveNext
					loop
				MYRS.Close	%>
				</tr>
</table>	
				
		<%	CASE 4
				MYRS.Open "select * from webmail where mail_Type=4 and stu_no='"& session("kaohao") &"' and qst_no='"& session("passwd") & "'"
				
					%>


<table width="100%" border="0" cellspacing="0" cellpadding="0" style="border-top-style: dotted; border-bottom-style: dotted">
			<%DO UNTIL MYRS.EOF%><tr>
				
		<td width="3%" height="25" bgcolor="#F0F0F0"><input type="checkbox" name="C<% =MYRS("mail_id") %> " & <% =MYRS("mail_id") %> value="checkbox"></td>
          
          <td width="20%" height="25" bgcolor="#F0F0F0"><% Response.WRITE MYRS("mail_from")%></td>
          <td width="50%" height="25" bgcolor="#F0F0F0">
<% 
 lcAttachment = Trim(MYRS("mail_attachment"))
 If Not IsNull(lcAttachment) And lcAttachment<>"" Then
    response.write "<img src='images/attachment.gif'>" 
 Else
    response.write "&nbsp;" 
 End If 
%>
<% response.Write("<a href='OpenOldMail.asp?uid=" & MYRS("mail_id") & "'>" & MYRS("mail_subject") & "</a>")%></td>
          <td width="16%" height="25" bgcolor="#F0F0F0"><% Response.WRITE MYRS("mail_datetime")%></td>
          
		  
		  </tr><%
					'Response.WRITE MYRS("mail_from")
					m_id		   = MYRS.fields("mail_ID")
					session("uid") = MYRS.fields("mail_ID")
					'response.Write("<a href='maillist.asp?uid=m_id'>" & MYRS("mail_subject") & "</a>")
					'Response.WRITE MYRS("mail_datetime")
					midd 		   = MYRS.fields("mail_ID")
					session("uid") = MYRS.fields("mail_ID")
					MYRS.MoveNext
					loop
				MYRS.Close	%>
				</tr>
</table>	
		<%END SELECT
	end function
%> 
</body>
</html>
















