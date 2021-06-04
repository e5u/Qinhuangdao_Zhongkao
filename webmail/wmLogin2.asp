<!--#include file="connsql.ASP"-->

<HTML><HEAD><TITLE>电子邮件系统</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<SCRIPT languange="javascript">
function openHelp()
{
	window.open("outlook_help.htm", "win4", "height=250,width=400");
}
function openStat()
{
	window.open("stat_domain.php", "win5", "height=550,width=800,left=0,top=0");
}
</SCRIPT>
<LINK href="images/css.css" rel=stylesheet>
<META content="Microsoft FrontPage 6.0" name=GENERATOR><style type="text/css">
<!--
body {
	background-color: #3B6D9E;
}
-->
</style>
<link href="images/css.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY leftMargin=0 background=images/bluebg.gif topMargin=0>

<%
  
 '2007/03/09:webmail表和webmailuser表中新增exm_no列
'Err.clear
' On error resume next
 
' lcSql = "select exm_no from webmail"  
 'MYRS.Open lcSql,MYcn
 'MYRS.Close
 'If err.number<>0 Then
' 	err.clear
' 	MYcn.execute "alter table webmail add exm_no char(20) null"
' End If
 
 'lcSql = "select exm_no from webmailuser"  
 'MYRS.Open lcSql,MYcn
 'MYRS.Close
 'If err.number<>0 Then
' 	err.clear
 '	MYcn.execute "alter table webmailuser add exm_no char(20) null"
 'End If
		
	'更新webmail题库
		 sub addnewwmail()
		 MYcn.Execute("insert into webmail select  b.* ,a.exm_no from question a , webmail0  b where a.qst_no=b.qst_no and   exm_no = '" & SESSION("EXM_NO") & "'")
		 end sub
				 
		 '取出现在练习rsd_id 为rstid
		 lSql = "select top 1 * from result where exm_no='"&SESSION("EXM_NO")&"' order by rst_id desc"
		 MYRS.Open lSql,MYcn
		 rstid= MYRS("rst_id")
		 MYRS.Close
		 
		 '取出已做练习的rsd_id 为 lastrstid
		 lcSql = "select * from result1 where  exm_no='"&SESSION("EXM_NO") & "'"   
         MYRS.Open lcSql,MYcn
		 
		 if MYRS.eof then
			 MYRS.Close
			 MYcn.Execute("insert into result1(rst_id,stu_no,qst_no,exm_no) values ("&rstid&", 'USER', '" & SESSION("QST_NO") & "','"&SESSION("EXM_NO")&"')") 
		     call addnewwmail
		 else
		    lastrstid=MYRS("rst_id")
		     MYRS.Close
		     if rstid-lastrstid>0 then  '现在练习与上次练习rstid不同，说明是重新发题
			 	 '更新lsatrstid为现在练习rstid
		     MYcn.Execute("update result1 set rst_id="&rstid &" where exm_no = '" & SESSION("EXM_NO") & "'") 
			 '删除旧题
			 'response.write("完成更新")
			 MYcn.Execute("delete * from webmail where exm_no = '" & SESSION("EXM_NO") & "'")
             MYcn.Execute("delete * from WebMailUser where exm_no = '" & SESSION("EXM_NO") & "'")
             call addnewwmail
			 end if
		 end if
		 response.Write("  当前题号："&SESSION("QST_NO")&"-"&SESSION("EXM_NO"))
%>

<form action="wmLogchk.asp" id="FORM" method="post" name="Form"> 

<TABLE height="100%" width="100%">
  <TBODY>
  <TR>
    <TD vAlign=center align=middle>
      <table width="643" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#205486">
        <tr>
          <td><table width="643" border="0" cellpadding="0" cellspacing="5" bgcolor="#D8EDFE">
              <tr>
                <td><table width="643" border="0" cellpadding="0" cellspacing="1" bgcolor="#205486">
                    <tr>
                      <td><table width="643" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td colspan="2"><img src="images/mail_03.gif" width="643" height="163" alt=""></td>
                          </tr>
                          <tr>
                            <td width="237"><img src="images/mail_05.gif" width="237" height="240" alt=""></td>
                            <td width="406" background="images/mail_06.gif"><table width="65%"  border="0" cellspacing="0" cellpadding="0">
                                <tr>
                                  <td height="30"><div align="center"><span class="unnamed2">用户名：</span>
                                          <INPUT name="kaohao_nouse"  class=field id="kaohao_nouse" size=13  readonly maxLength=20 value="<%=SESSION("Stu_No") %>"> 
                                  </div></td>
                                </tr>
                                <tr>
                                  <td height="30"><div align="center"><span class="unnamed2">密　码：</span>
                                          <INPUT name="passwd_nouse"  class=field id="passwd_nouse" size=13  maxLength=20> 
                                  </div></td>
                                </tr>
                                <tr>
                                  <td height="25"><div align="center" class="unnamed2">[<a href="wmUserList.asp">申请邮箱</a>] 
                                      　　[忘记密码] </div></td>
                                </tr>
                                <tr>
                                  <td height="25"><div align="center">
                                      <input type="submit" name="Submit" value="登  录">
                                  </div></td>
                                </tr>
                            </table></td>
                          </tr>
                      </table></td>
                    </tr>
                </table></td>
              </tr>
          </table></td>
        </tr>
      </table>      </TD>
  </TR></TBODY></TABLE>
   <input name="frmLocalIP"  type="hidden" value="" >
   <input name="frmWsNo"     type="hidden" value="" >
   <input name="frmTesterNo" type="hidden" value="" >        

<INPUT type="hidden" name="kaohao"  class=field id="kaohao" size=13  maxLength=20 value="<% =SESSION("Stu_No") %>"> 
<INPUT type="hidden" name="passwd"  class=field id="passwd" size=13  maxLength=20 value="<% =SESSION("Qst_No") %>"> 
   </FORM>

</BODY></HTML>
<SCRIPT  LANGUAGE=vbscript>    
<!--    
Function form_onsubmit()                                                     
  'lbRTV = False                                         
  'If Trim(form.kaohao.value)="" Then                                         
  '   MsgBox "登录名不能为空!", vbOkOnly+vbExclamation, "数据检查"                           
  'ElseIf  Trim(form.passwd.value)="" Then                                         
  '   MsgBox "密码不能为空!", vbOkOnly+vbExclamation, "数据检查" 
  'Else                           
  '   lbRTV = True                           
  'End If
  lbRTV = True                                         
  form_onsubmit = lbRTV         
End Function                          
-->    
</SCRIPT>  
