<%@ Language=VBScript %>
<!--#include file="connsql.ASP"-->
<html>
<head>
<title>168�����ʼ�ϵͳ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="images/css.css" rel="stylesheet" type="text/css">
</head>
<%
   If session("kaohao") = "" then
   	  Response.Write "<div align=center>"
      Response.Write "<p>&nbsp;</p><p>&nbsp;</p>"
      Response.Write "<H3>����ʧЧ,���뿪��ʱ��̫����,�����µ�¼</H3>"
      Response.Write "<a href='default.htm' name='link' target='_top'></a>"
      Response.Write "<p>&nbsp;</p>"
      Response.Write "<input type='button' value='���µ�¼' name='cancel' id=cancel>"
      Response.Write "</div>"
      
      Response.Write "<SCRIPT LANGUAGE=VBscript> " & Chr(13) & Chr(10)
      Response.Write "<!-- " & Chr(13) & Chr(10)
      Response.Write "Sub cancel_onclick " & Chr(13) & Chr(10)
      Response.Write "link.click() "  & Chr(13) & Chr(10)
      Response.Write "End Sub " & Chr(13) & Chr(10)
      Response.Write "!-->" & Chr(13) & Chr(10)
      Response.Write "</SCRIPT>" & Chr(13) & Chr(10)
      Response.End
   End If
%>
<%
   MYRS.Open "select * from webmail where mail_type=1 and stu_no='"& session("kaohao") &"' and qst_no='"& session("passwd") & "'"
   Session("MailsNum")= MYRS.RecordCount
   MYRS.Close		

   MYRS.Open "select * from webmail where mail_Type=2 and stu_no='"& session("kaohao") &"' and qst_no='"& session("passwd") & "'"
   Session("MailsForModiNum")= MYRS.RecordCount
   MYRS.Close		
   
   MYRS.Open "select * from webmail where mail_Type=3 and stu_no='"& session("kaohao") &"' and qst_no='"& session("passwd") & "'"
   Session("MailsForSendNum")= MYRS.RecordCount
   MYRS.Close		

   MYRS.Open "select * from webmail where mail_Type=4 and stu_no='"& session("kaohao") &"' and qst_no='"& session("passwd") & "'"
   Session("MailsForDeleteNum")= MYRS.RecordCount
   MYRS.Close		
   
   MYRS.Open "select * from webmail where mail_type=1 and mail_isnew=1 and stu_no='"& session("kaohao") &"' and qst_no='"& session("passwd") & "'"
   Session("MailsForReadNum")=MYRS.RecordCount
   MYRS.Close

   MYRS.Open "select * from webmail where stu_no='"& session("kaohao") &"' and qst_no='"& session("passwd") & "'"
   Session("MailsNumALL")= MYRS.RecordCount
   MYRS.Close		
%>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="824" valign="top"><form name="form1" method="post" action="">
      <table width="100%" border="0" cellspacing="0" cellpadding="5">
        <tr>
          <td bgcolor="#ECFFFE" class="unnamed1">���ã�<% =session("stu_no") %> ����鿴���֡������ռ��乲��<% =Session("MailsNum") %>���ʼ�������<% =Session("MailsForReadNum") %>��δ�� �����ļ���</td>                               
        </tr>                             
      </table>                             
      <table width="100%" border="0" cellspacing="0" cellpadding="0">                             
        <tr>                             
          <td height="1" bgcolor="#91C8F1" class="unnamed1"></td>                             
        </tr>                             
      </table>                             
      <table width="100%" border="0" cellspacing="0" cellpadding="5">                             
        <tr>                             
          <td class="unnamed1">����<% = Session("MailsNumALL") %>���ʼ���<% =Session("MailsForReadNum") %>��δ�����ܿռ�3G���Ѿ�ʹ��2K��0.01%��                     
          </td>                             
        </tr>                             
        <tr>                             
          <td class="unnamed1">                             
          </td>                             
        </tr>                             
      </table>                             
      <table width="98%" border="0" align="center" cellpadding="0" cellspacing="0">                             
        <tr>                             
          <td height="30" background="images/dot08.gif" class="unnamed1">                             
            <div align="center">�ļ���</div></td>                             
          <td background="images/dot08.gif" class="unnamed1"><div align="center">δ���ʼ�</div></td>                             
          <td background="images/dot08.gif" class="unnamed1"><div align="center">�ܷ���</div></td>                             
          <td background="images/dot08.gif" class="unnamed1"><div align="center">�ռ��С</div></td>                             
          <td background="images/dot08.gif"><p align="center" class="unnamed1">����</p></td>                             
        </tr>                             
        <tr>                             
          <td height="25" class="unnamed1"><div align="center">�ռ���</div></td>                             
          <td class="unnamed1"><div align="center"><% =Session("MailsForReadNum") %></div></td>                             
          <td class="unnamed1"><div align="center"><% =Session("MailsNum") %></div></td>                             
          <td class="unnamed1"><div align="center">                 
          <%                 
          If Session("MailsNum")= 0 Then                   
             response.write "0K"                             
          Else                  
             response.write "2K"                  
          End If                 
          %>                  
          </div></td>                 
          <td><div align="center" class="unnamed1">���</div></td>                             
        </tr>                             
        <tr>                             
          <td height="25" class="unnamed1"><div align="center">�ݸ���</div></td>                             
          <td class="unnamed1"><div align="center">0</div></td>                             
          <td class="unnamed1"><div align="center"><% =Session("MailsForModiNum") %></div></td>                             
          <td class="unnamed1"><div align="center">      
          <%                 
          If Session("MailsForModiNum")= 0 Then                   
             response.write "0K"                             
          Else                  
             response.write "2K"                  
          End If                 
          %>                 
          </div></td>                             
          <td><div align="center" class="unnamed1">���</div></td>                             
        </tr>                             
        <tr>                             
          <td height="25" class="unnamed1"><div align="center">�ѷ���</div></td>                             
          <td class="unnamed1"><div align="center">0</div></td>                             
          <td class="unnamed1"><div align="center"><%=Session("MailsForSendNum")%></div></td>                             
          <td class="unnamed1"><div align="center">      
          <%                 
          If Session("MailsForSendNum")= 0 Then                   
             response.write "0K"                             
          Else                  
             response.write "2K"                  
          End If                 
          %>                 
          </div></td>                             
          <td><div align="center" class="unnamed1">���</div></td>                             
        </tr>                             
        <tr>                             
          <td height="25" class="unnamed1"><div align="center">��ɾ��</div></td>                             
          <td class="unnamed1"><div align="center">0</div></td>                             
          <td class="unnamed1"><div align="center"><%=Session("MailsForDeleteNum")%></div></td>                             
          <td class="unnamed1"><div align="center">      
          <%                 
          If Session("MailsForDeleteNum")= 0 Then                   
             response.write "0K"                             
          Else                  
             response.write "2K"                  
          End If                 
          %>                 
          </div></td>                             
          <td><div align="center" class="unnamed1">���</div></td>                             
        </tr>                             
        <tr>                             
          <td height="25" class="unnamed1"><div align="center">����ʼ�</div></td>                             
          <td class="unnamed1"><div align="center">0</div></td>                             
          <td class="unnamed1"><div align="center">0</div></td>                             
          <td class="unnamed1"><div align="center">0K</div></td>                             
          <td><div align="center" class="unnamed1">���</div></td>                             
        </tr>                             
        <tr>                             
          <td height="25" class="unnamed1"><div align="center">�����ʼ�</div></td>                             
          <td class="unnamed1"><div align="center">0</div></td>                             
          <td class="unnamed1"><div align="center">0</div></td>                             
          <td class="unnamed1"><div align="center">0K</div></td>                             
          <td><div align="center" class="unnamed1">���</div></td>                             
        </tr>                             
      </table>                             
      <table width="100%" border="0" cellspacing="0" cellpadding="0">                             
        <tr>                             
          <td height="30"> ��                             
              <input type="hidden" name="Submit" value="�½��ļ���">                             
              <input type="hidden" value="��������ռ���" name="B1">                             
</td>                             
        </tr>                             
      </table>                             
    </form></td>                             
    <td width="13"></td>                             
  </tr>                             
</table>                             
</body>                             
</html>