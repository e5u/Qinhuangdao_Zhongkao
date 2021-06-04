<%@ Language=VBScript %>
<html>
<head>
<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>电子邮件系统</title>
<link rel="stylesheet" type="text/css" href="mystyle.css">
</head>
<%
   If session("kaohao") = "" then
      Response.Write "<div align=center>"
      Response.Write "<p>&nbsp;</p><p>&nbsp;</p>"
      Response.Write "<H3>连接失效,您离开的时间太长了,请重新登录</H3>"
      Response.Write "<a href='default.htm' name='link' target='_top'></a>"
      Response.Write "<p>&nbsp;</p>"
      Response.Write "<input type='button' value='重新登录' name='cancel' id=cancel>"
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

<frameset rows="56,*,64" framespacing="0" border="0" frameborder="0">
  <frame name="banner" scrolling="no" noresize target="contents" src="wmhead.htm">
  <frameset cols="213,*">
    <frame name="contents" target="main" src="wmMenu.htm" scrolling="no" noresize>
    <frame name="main" src="welcome.asp" scrolling="auto">
  </frameset>
  <frame name="bottom" scrolling="no" noresize target="contents" src="wmBottom.htm">
  <noframes>
  <body>

  <p>此网页使用了框架，但您的浏览器不支持框架。</p>

  </body>
  </noframes>
</frameset>
</html>
