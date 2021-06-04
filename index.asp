<%SESSION("QST_NO")=trim(request("qstno"))
  SESSION("EXM_NO")=trim(request("exmno"))
   lcAppName = "QLIB"
  SESSION("DBCONNECTSTRING") = GetDBStrByApp(lcAppName)
     Dim MYcn
     Set MYcn = Server.CreateObject("ADODB.Connection")
     MYcn.open SESSION("DBCONNECTSTRING") 
        SESSION("STU_NO") = Trim("user")
        Response.Redirect "webmail/wmLogin2.asp"

  Function GetDBStrByApp(pcApp)
     If UCase(pcApp) = "QLIB" Then
        Dim lcDbPath
        lcDbPath = LCase(Server.Mappath("weblogin.asp")) 
        lcDbPath = Replace(lcase(lcDbPath),  "\weblogin.asp", "")
        lcDbPath = Left(lcDbPath, InstrRev(lcDbPath, "\")) 
        lcDbPath = lcDbPath & "question.mdb"
        GetDBStrByApp = "provider=microsoft.jet.oledb.4.0;data source=question.mdb"   
     End If
  End Function
%>
