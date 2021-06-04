
<%  
  On Error Resume Next
  Dim MYcn,MYRS
 Set MYcn = Server.CreateObject("ADODB.Connection")
 MYcn.open GetDBStr
 If Err.Number<>0 Then
    Response.Write "<p align='center'><h3>数据库连接出错！" & Err.Description & "</h></p>"
    Response.End
 End If
	 
 SET MYRS = SERVER.CREATEOBJECT("ADODB.RECORDSET")
 MYRS.ActiveConnection = MYcn
 MYRS.CursorLocation = 3
 MYRS.CursorType = 3
 MYRS.LockType = 2

 Function GetDBStr()
   GetDBStr = SESSION("DBCONNECTSTRING")
 End Function
%>
