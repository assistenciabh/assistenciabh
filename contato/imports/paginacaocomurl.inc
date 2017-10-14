<%
dim intPageSize 'Tamanho da página de dados em nº de registros
dim intPageCount 'Numero de páginas disponiveis 
dim intPageCurrent 'Pagina atual
dim intPaginaLixo
Dim i
Dim strHidden
Dim strQuery

sub conectar(strSQL,intPagina)
	

	if len(trim(intPagina)) <> 0 then
		if IsNumeric(intPagina) then
			intPageSize = cInt(intPagina)
		else
			intPageSize = 10
		end if
	else
		intPageSize = 10
	end if
	
	'Verifica se foi passado o numero da pagina atual
	If recupera("page") = "" Then
		intPageCurrent = 1
	Else
		if recupera("co") = "c" then
			intPageCurrent = 1
		else
			intPageCurrent = CInt(recupera("page"))
		end if
	End If
	set rs = nothing
	Set rs = Server.CreateObject("ADODB.RecordSet")
	'Define o tamanho da pagina em quantidade de registros
	rs.PageSize = CInt(intPageSize)
	rs.CursorLocation = 3
	
	rs.Open strSQL, connDB,3,1
	if not rs.eof then	
		'Verifica o numero de paginas criadas pelo recordset
		intPageCount = rs.PageCount
		'Move para a pagina selecionada
		rs.AbsolutePage = intPageCurrent
	end if
end sub

sub paginar(strpagina, strvar, strQuery, strform, url)
	'rs.movefirst
	'if not rs.eof then
Select case rs.eof
	case rs.bof <> rs.eof	
	%>
	<script language="javascript">
		function submeter(pag){
		    pag = pag + '<%=url%>';			
			document.forms['<%= strform%>'].action = pag;
			document.forms['<%= strform%>'].submit();		
		}
	</script>
	<table width="50%" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td align="center">
				<%
				If intPageCount > 10 then
					If intPageCurrent <> 1 Then
					%>
						<a href="#" onclick="JavaScript:submeter('<%= strpagina%>?page=1<%= strQuery%>');">
						<img src="ico/pag01.gif" border="0" alt="Primeira Página">
						</a>
					<%
					Else
						Response.Write "<img src=ico/pag01d.gif' border='0' alt='Primeira Página'>"
					End If
				Else
					Response.Write "<img src='ico/pag01d.gif' border='0' alt='Primeira Página'>"
				End If
				%>
			</td>
			<td align="center">
				<%
				'response.write(intPageCount&" - "&intPageCurrent&" - "&(cInt(intPageCurrent) <> 1 ))
				If cInt(intPageCount) > 10 then
'					response.write("Aqui")
					If cInt(intPageCurrent) <> 1 Then
						if cInt(intPageCurrent) < 11 then
						%>
							<a href="#" onclick="JavaScript:submeter('<%= strpagina%>?page=<%= intPageCurrent - 1%><%= strQuery%>');">
							<img src="ico/pag02.gif" border="0" alt="Página Anterior">
							</a>
						<%
						else
						%>
							<a href="#" onclick="JavaScript:submeter('<%= strpagina%>?page=<%= intPageCurrent - 10%><%= strQuery%>');">
							<img src="ico/pag02.gif" border="0" alt="Página Anterior">
							</a>
						<%
						end if
					Else
						Response.Write "<img src='ico/pag02d.gif' border='0' alt='Página Anterior'>"
					End If
				Else
					Response.Write "<img src='ico/pag02d.gif' border='0' alt='Página Anterior'>&nbsp;"
				End if
				%>
			</td>
				<%
				intPaginaLixo = intPageCurrent
				while ((right(cstr(intPaginaLixo),1) <> 1) and (right(cstr(intPaginaLixo),1) <> 11))'11
					intPaginaLixo = cint(intPaginaLixo) - 1
				Wend
				for i = intPaginaLixo to (intPaginaLixo+9)
				%>
					<%if cInt(i) = cInt(intPageCurrent) then%>
						<td align="center">
							<font size="1" face="Verdana, Arial, Helvetica, sans-serif">
							<strong><font color="#0000CC"><%= i%></font></strong>
							</font>
						</td>
					<%else%>
						<td align="center">
							<font size="1" face="Verdana, Arial, Helvetica, sans-serif">
							<a href="#" onclick="JavaScript:submeter('<%= strpagina%>?page=<%= i%><%= strQuery%>');">
							<%= i%>
							</A>
							</font>
						</td>
					<%end if%>
					<%
					if cint(i) = cint(intPageCount) then
						exit for
					end if
					%>
				<%next%>
				<%= strvar%>
				<td align="center">
				<%
				If intPageCount > 10 then
					If intPageCurrent < intPageCount Then
					%>
						<a href="#" onclick="JavaScript:submeter('<%= strpagina%>?page=<%= i%><%= strQuery%>');">
						<img src="ico/pag03.gif" border="0" alt="Próxima Página">
						</a>
					<%
					Else
						Response.Write "<img src='ico/pag03d.gif' border='0' alt='Próxima Página'>"
					End If
				Else
					Response.Write "<img src='ico/pag03d.gif' border='0' alt='Próxima Página'>"
				End if
				%>
				</td>
			<td align="center">
				<%'response.write intPageCount
				If intPageCount > 10 then
					If intPageCurrent < intPageCount Then
						%>
							<a href="#" onclick="JavaScript:submeter('<%= strpagina%>?page=<%= intPageCount%><%= strQuery%>');">
							<img src="ico/pag04.gif" border="0" alt="Última Página">
							</a>
						<%
					Else
						Response.Write "<img src='ico/pag04d.gif' border='0' alt='Última Página'>"
					End If
				Else
					Response.Write "<img src='ico/pag04d.gif' border='0' alt='Última Página'>"
				End if
				%>
			</td>
		</tr>
	</table>
	<%
	'end if
end select
%>
<input type="hidden" value="<%= intPageCurrent%>" name="pagina_atual">
<%
end sub
%>