<!--#include file="imports/funcoesmmq.asp"-->
<%
response.charset = "iso-8859-1"
nome = recupera("nome")
telefone = recupera("telefone")
fromemail = recupera ("email")
assunto = recupera ("assunto")
mensagem = recupera ("mensagem")
toemail =  recupera ("toemail")

Set Mail = Server.CreateObject("Persits.MailSender")

Mail.Host = "smtp.gmail.com"
Mail.Port = 465
Mail.Username = "ferreirabdiego@gmaill.com"
Mail.Password = "Fl0r3st@"
Mail.From = fromemail
Mail.FromName = nome
Mail.AddAddress "ferreirabdiego@gmaill.com"
Mail.Subject = assunto
Mail.Body = "Nome: "&nome&"<br>Telefone: "&telefone&"<br>Email: "&fromemail&"<br><br>Mensagem: "&mensagem 
Mail.IsHTML = True

On Error Resume Next 
Mail.Send 
If Err <> 0 Then 
Response.Write "Ocorreu o seguinte erro: " & Err.Description 
else 
response.write "Email enviado com sucesso"
End If


%>





