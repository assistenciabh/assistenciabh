<%
set connDB = server.CreateObject ("ADODB.Connection")
Sub OpenCn(Cn)
	 if ConnDB.state = 0 Then 		
	connDB.Open "Driver={SQL Server};Server=beta\si;Uid=1b[07=a5;Pwd=1313;Database=master;Network=DBMSSOCN;APP=Conexão Via Site. CodUser: "  & Session("coduser") & " IdUsuario: " & Session("IDUsuario") & " IP: " & Request.ServerVariables("REMOTE_ADDR") & " URL: " & Request.ServerVariables("URL") & " Query STR:" & formulario & replace(Request.querystring, ";", "")
		End if
end sub
  if ConnDB.state = 0 Then 		

        set connDB = server.CreateObject ("ADODB.Connection")
	connDB.Open "Driver={SQL Server};Server=beta\si;Uid=1b[07=a5;Pwd=1313;Database=master;Network=DBMSSOCN;APP=Conexão Via Site. CodUser: "  & Session("coduser") & " IdUsuario: " & Session("IDUsuario") & " IP: " & Request.ServerVariables("REMOTE_ADDR") & " URL: " & Request.ServerVariables("URL") & " Query STR:" & formulario & replace(Request.querystring, ";", "")
   End if	



%>