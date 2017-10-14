
<%
'Seta Configurações Regionais para Portugues(Brazil)
Session.LCID = 1046 

function Nulo(val,tipo)
      if tipo = "1" then	'Apenas String 
			if val <> "" AND val <> "-1" AND LCase(val) <> "null" then
				 val = "'"&val&"'"
			else
				 val = "null"	 
			end if
	  else 		' tipo = 2 Numeros
	  		if val <> "" AND val <> "-1" AND LCase(val) <> "null"  then
				 val = val
			else
				 val = "null"	 
			end if
	  end if		
	  Nulo = val 
end function
function FnChecked(Campo)
        if Ucase(Trim(Campo)) = "ON" Then 
           FnChecked =  1
        Else
            FnChecked = 0
        End if
End Function

function ValidaIM(IM)
	if ISNUMERIC(IM) then 
		for i = 1 to  len(IM)
			sumIM = CINT(sumIM) + CINT(mid(IM,i,1))			
		next
		if sumIM = 0 then
			ValidaIM = "ISENTO"
		else
			ValidaIM = IM
		end if
	elseif UCASE(IM) = "ISENTA" then
		   ValidaIM = "ISENTO"		
	else 
		   ValidaIM = IM  
	end if
end function

'ULR Decode valores
Function URLDecode(sConvert)
    Dim aSplit
    Dim sOutput
    Dim I
    If IsNull(sConvert) Then
       URLDecode = ""
       Exit Function
    End If

    ' convert all pluses to spaces
    sOutput = REPLACE(sConvert, "+", " ")

    ' next convert %hexdigits to the character
    aSplit = Split(sOutput, "%")

    If IsArray(aSplit) Then
      sOutput = aSplit(0)
      For I = 0 to UBound(aSplit) - 1
        sOutput = sOutput & _
          Chr("&H" & Left(aSplit(i + 1), 2)) &_
          Right(aSplit(i + 1), Len(aSplit(i + 1)) - 2)
      Next
    End If

    URLDecode = sOutput
End Function

'Fecha recordSet e Conexao BD 
Sub FechaCn(Cn)
	If cn.state <> 0 then cn.Close()	
	set Cn = Nothing
End Sub

'recupera Data do Banco de Dados
function dbdate()
	RsData = ConnDB.Execute ("SET DATEFORMAT DMY SELECT CONVERT(VARCHAR,GETDATE(),103) AS DATA")
	dbdate = CDate(RsData("DATA"))
end function

'recupera Hora do Banco de Dados
function dbtime()
	RsData = ConnDB.Execute ("SET DATEFORMAT DMY SELECT CONVERT(VARCHAR,GETDATE(),108) AS DATA")
	dbtime = RsData("DATA")
end function

'formata data com 0
function formatadata(sdata)
if sdata<>"" then
	data = cdate(sdata)
	dia = day(data)
	mes = month(data)
	ano = year(data)
	if dia < 10 then
		dia = "0" & CStr(dia)
	end if
	
	if mes < 10 then 
	mes = "0" & CStr(mes)
	end if
		
	formatadata = CStr(dia) + "/" + CStr(mes) + "/" + CStr(ano)
end if
end function

Function MesExt(mes)
	if mes<> "" then
		mes = cint(mes)
		select case mes
			case 1
				MesExt = "Janeiro"
			case 2
				MesExt = "Fevereiro"
			case 3
				MesExt = "Março"
			case 4
				MesExt = "Abril"
			case 5
				MesExt = "Maio"
			case 6
				MesExt = "Junho"
			case 7
				MesExt = "Julho"
			case 8
				MesExt = "Agosto"
			case 9
				MesExt = "Setembro"
			case 10
				MesExt = "Outubro"
			case 11
				MesExt = "Novembro"
			case 12
				MesExt = "Dezembro"
		end select
	end if
End Function

Function recuperacompleto(Nome)

array_split_item = ("/*,*/,@@,char,nchar,varchar,nvarchar,begin,cast,create,cursor,declare,delete,drop,end,exec,execute, fetch,insert,kill,open,select,sys,sysobjects,syscolumns,alter,update,script")

array_separadores = ("-,=,*,!,@,$,%,&,>,<,(,),|,{,},[,],#, ,;,:,.,',+,/,\,§,"",")

array_split_item = Split(array_split_item, ",")

array_separadores = Split(array_separadores, ",")

recuperacompleto = request(Nome)
recuperacompleto = Replace(recuperacompleto, "'", " ", 1)
recuperacompleto = Replace(recuperacompleto, "--", " ", 1)
recuperacompleto = Replace(recuperacompleto, Chr(37) + ">", " ", 1)
recuperacompleto = Replace(recuperacompleto, "<" + Chr(37), " ", 1)
recuperacompleto = Replace(recuperacompleto, "chr(37)", " ", 1)
recuperacompleto = Replace(recuperacompleto, "declare", " ", 1)

For i = 0 To UBound(array_split_item)
    If InStr(recuperacompleto, array_split_item(i)) Then
        If array_riscos = "" Then
            array_riscos = array_split_item(i)
        Else
            array_riscos = array_riscos & "," & array_split_item(i)
        End If
    End If
Next
 
 if array_riscos <>"" then
	For i = 0 To UBound(array_separadores)
		If InStr(recuperacompleto, array_separadores(i)) Then
			If array_separadores_uso="" Then
				array_separadores_uso = array_separadores(i)
			Else
				array_separadores_uso = array_separadores_uso & "," & array_separadores(i)
			End If
		End If
	Next
	 
	 
	array_riscos = Split(array_riscos, ",")
	array_separadores_uso = Split(array_separadores_uso, ",")

  'end if
erro = 0

 'if array_riscos <>"" then
	If Ubound(array_riscos) >-1 Or Ubound(array_separadores_uso) >-1 Then
		   
			For i = 0 To UBound(array_riscos)
				If erro = 0 Then
					If Ubound(array_separadores_uso) >-1 Then
						For j = 0 To UBound(array_separadores_uso)
							array_nome = recuperacompleto
							array_nome = Split(array_nome, array_separadores_uso(j))
							for k = 0 to ubound(array_nome)
							If array_nome(k) = array_riscos(i) Then
								response.Write "<script>alert('Não são aceitas palavras que possuem: " & array_riscos(i) & "');    </script>"
								erro = 1
							End If
							next
						Next
					'Else
'						If Recupera = array_riscos(i) Then
'							response.Write "<script>alert('Não são aceitas palavras que possuem: " & array_riscos(i) & "');    <script>"
'							erro = 1
'						End If
					End If
				End If
			Next
	End If
end if
    If erro = 1 Then
            recuperacompleto = ""
            %>
                <script>
                    javascript:history.go(-1);
                </script>
            <%
    End If
End Function


'utilizado no lugar do request
Function Recupera(Nome)
'array_split_item = ("/*,*/,@@,char,nchar,varchar,nvarchar,begin,cast,create,cursor,declare,delete,drop, end,exec,execute, fetch,insert,kill,open,select,sys,sysobjects,syscolumns,alter,update,<script")


array_split_item = ("/*,*/,@@,char,nchar,varchar,nvarchar,begin,cast,create,cursor,declare,delete,drop,end,exec,execute, fetch,insert,kill,open,select,sys,sysobjects,syscolumns,alter,update,script")

array_split_item = split(array_split_item, ",")

	erro = 0
	Recupera=trim(request(Nome))
	Recupera=Replace(Recupera,"'"," ",1)
	Recupera=Replace(Recupera,"--"," ",1)
	Recupera=Replace(Recupera,chr(37) + ">"," ",1)
	Recupera=Replace(Recupera,"<" + chr(37)," ",1)
	Recupera=Replace(Recupera,"chr(37)"," ",1)
	Recupera=Replace(Recupera,"declare"," ",1)     
    
	if session("idusuario") = "" then
		for ww = 0 to UBound(array_split_item)            
			if instr(Recupera, array_split_item(ww)) > 0 then
				erro = 1
			end if
		next	
	end if
	if erro = 1 then
		'if session("idusuario") = "" then
			Recupera = ""
			'response.write "Algum caracter digitado não é válido."
			%>
				<script>
					alert('Algum dado digitado não é aceito.');
				</script>
                
                <script>
					window.location.href = 'http://www.mastermaq.com.br';
				</script>
			<%
			
		'end if
	end if
end function


Function recuperacompletoAJAX(Nome)

array_split_item = ("/*,*/,@@,char,nchar,varchar,nvarchar,begin,cast,create,cursor,declare,delete,drop,end,exec,execute, fetch,insert,kill,open,select,sys,sysobjects,syscolumns,alter,update,script")

array_separadores = ("-,=,*,!,@,$,%,&,>,<,(,),|,{,},[,],#, ,;,:,.,',+,/,\,§,"",")

array_split_item = Split(array_split_item, ",")

array_separadores = Split(array_separadores, ",")

recuperacompletoAJAX = request(Nome)
recuperacompletoAJAX = Replace(recuperacompletoAJAX, "'", " ", 1)
recuperacompletoAJAX = Replace(recuperacompletoAJAX, "--", " ", 1)
recuperacompletoAJAX = Replace(recuperacompletoAJAX, Chr(37) + ">", " ", 1)
recuperacompletoAJAX = Replace(recuperacompletoAJAX, "<" + Chr(37), " ", 1)
recuperacompletoAJAX = Replace(recuperacompletoAJAX, "chr(37)", " ", 1)
recuperacompletoAJAX = Replace(recuperacompletoAJAX, "declare", " ", 1)

For i = 0 To UBound(array_split_item)
    If InStr(recuperacompletoAJAX, array_split_item(i)) Then
        If array_riscos = "" Then
            array_riscos = array_split_item(i)
        Else
            array_riscos = array_riscos & "," & array_split_item(i)
        End If
    End If
Next
 
 if array_riscos <>"" then
	For i = 0 To UBound(array_separadores)
		If InStr(recuperacompletoAJAX, array_separadores(i)) Then
			If array_separadores_uso="" Then
				array_separadores_uso = array_separadores(i)
			Else
				array_separadores_uso = array_separadores_uso & "," & array_separadores(i)
			End If
		End If
	Next
	 
	 
	array_riscos = Split(array_riscos, ",")
	array_separadores_uso = Split(array_separadores_uso, ",")

  'end if
erro = 0

 'if array_riscos <>"" then
	If Ubound(array_riscos) >-1 Or Ubound(array_separadores_uso) >-1 Then
		   
			For i = 0 To UBound(array_riscos)
				If erro = 0 Then
					If Ubound(array_separadores_uso) >-1 Then
						For j = 0 To UBound(array_separadores_uso)
							array_nome = recuperacompletoAJAX
							array_nome = Split(array_nome, array_separadores_uso(j))
							for k = 0 to ubound(array_nome)
							If array_nome(k) = array_riscos(i) Then								
								response.Write "Não são aceitas palavras que possuem: " & array_riscos(i):response.end 							
								erro = 1
							End If
							next
						Next

					End If
				End If
			Next
	End If
end if
    If erro = 1 Then
            recuperacompletoAJAX = ""	           
    End If
End Function


'utilizado no lugar do request
Function RecuperaAJAX(Nome)
'array_split_item = ("–,;,/*,*/,@@,char,nchar,varchar,nvarchar,begin,cast,create,cursor,declare,delete,drop, end,exec,execute, fetch,insert,kill,open,select,sys,sysobjects,syscolumns,alter,update,<script,<script>, ‘")
array_split_item = ("/*,*/,@@,char,nchar,varchar,nvarchar,begin,cast,create,cursor,declare,delete,drop,end,exec,execute, fetch,insert,kill,open,select,sys,sysobjects,syscolumns,alter,update,script")
array_split_item = split(array_split_item, ",")

	erro = 0
	RecuperaAJAX=request(Nome)
	RecuperaAJAX=Replace(RecuperaAJAX,"'"," ",1)
	RecuperaAJAX=Replace(RecuperaAJAX,"--"," ",1)
	RecuperaAJAX=Replace(RecuperaAJAX,chr(37) + ">"," ",1)
	RecuperaAJAX=Replace(RecuperaAJAX,"<" + chr(37)," ",1)
	RecuperaAJAX=Replace(RecuperaAJAX,"chr(37)"," ",1)
	RecuperaAJAX=Replace(RecuperaAJAX,"declare"," ",1)
	if session("idusuario") = "" then
		for ww = 0 to ubound(array_split_item)
			if instr(RecuperaAJAX, array_split_item(ww)) > 0 then				
				erro = 1
			end if
		next	
	end if
	if erro = 1 then		
			RecuperaAJAX = ""
			response.write "Algum dado digitado não é aceito.":response.end	
	end if
end function


Function TrocaPonto(valor)
	for i = 1 to len(valor)
		if mid(valor,i,1) = "." then 
			valor=left(valor,i-1) & "," & mid(valor,i+1,100)
		elseif mid(valor,i,1) = "," then 
			valor=left(valor,i-1) & "." & mid(valor,i+1,100)
		elseif mid(valor,i,1) = "" then 
			valor=left(valor,i-1) & "," & mid(valor,i+1,100)
		end if
	next
	trocaponto=valor
end function


function ZeroLeft(sNumber,iZeros)
	if Len(sNumber) >= iZeros then ZeroLeft=sNumber:exit function
	iLen=iZeros-Len(sNumber)
	ZeroLeft = String(iLen,"0")& sNumber
end function

'|///////////////////////////////////////////////////////////|
'| |
'| Funcao para calcular cpf |
'| |
'|///////////////////////////////////////////////////////////|
function CalculaCPF(CNPJ)
Dim RecebeCPF, Numero(11), soma, resultado1, resultado2
RecebeCPF = CNPJ
'Retirar todos os caracteres que nao sejam 0-9
s="" 
for x=1 to len(RecebeCPF)
ch=mid(RecebeCPF,x,1)
if asc(ch)>=48 and asc(ch)<=57 then
s=s & ch
end if
next
RecebeCPF = s

if len(RecebeCPF) <> 11 then
	CalculaCPF = false
elseif RecebeCPF = "00000000000" then
	CalculaCPF = false
else

Numero(1) = Cint(Mid(RecebeCPF,1,1))
Numero(2) = Cint(Mid(RecebeCPF,2,1))
Numero(3) = Cint(Mid(RecebeCPF,3,1))
Numero(4) = Cint(Mid(RecebeCPF,4,1))
Numero(5) = Cint(Mid(RecebeCPF,5,1))
Numero(6) = CInt(Mid(RecebeCPF,6,1))
Numero(7) = Cint(Mid(RecebeCPF,7,1))
Numero(8) = Cint(Mid(RecebeCPF,8,1))
Numero(9) = Cint(Mid(RecebeCPF,9,1))
Numero(10) = Cint(Mid(RecebeCPF,10,1))
Numero(11) = Cint(Mid(RecebeCPF,11,1))

soma = 10 * Numero(1) + 9 * Numero(2) + 8 * Numero(3) + 7 * Numero(4) + 6 * Numero(5) + 5 * Numero(6) + 4 * Numero(7) + 3 * Numero(8) + 2 * Numero(9)

soma = soma -(11 * (int(soma / 11)))

if soma = 0 or soma = 1 then
resultado1 = 0
else
resultado1 = 11 - soma
end if

if resultado1 = Numero(10) then

soma = Numero(1) * 11 + Numero(2) * 10 + Numero(3) * 9 + Numero(4) * 8 + Numero(5) * 7 + Numero(6) * 6 + Numero(7) * 5 + Numero(8) * 4 + Numero(9) * 3 + Numero(10) * 2

soma = soma -(11 * (int(soma / 11)))

if soma = 0 or soma = 1 then
resultado2 = 0
else
resultado2 = 11 - soma
end if

if resultado2 = Numero(11) then
	CalculaCPF = true
else
	CalculaCPF = false
end if
else 
	CalculaCPF = false
end if
end if

end function

'|///////////////////////////////////////////////////////////|
'| |
'| Funcao para calcular CNPJ |
'| |
'|///////////////////////////////////////////////////////////|

function CalculaCNPJ(CNPJ) 
Dim RecebeCNPJ, Numero(14), soma, resultado1, resultado2
RecebeCNPJ = CNPJ

s="" 
for x=1 to len(RecebeCNPJ)
ch=mid(RecebeCNPJ,x,1)
if asc(ch)>=48 and asc(ch)<=57 then
s=s & ch
end if
next
RecebeCNPJ = s

if len(RecebeCNPJ) <> 14 then
	CalculaCNPJ = false
elseif RecebeCNPJ = "00000000000000" then
	CalculaCNPJ = false
else

Numero(1) = Cint(Mid(RecebeCNPJ,1,1))
Numero(2) = Cint(Mid(RecebeCNPJ,2,1))
Numero(3) = Cint(Mid(RecebeCNPJ,3,1))
Numero(4) = Cint(Mid(RecebeCNPJ,4,1))
Numero(5) = Cint(Mid(RecebeCNPJ,5,1))
Numero(6) = CInt(Mid(RecebeCNPJ,6,1))
Numero(7) = Cint(Mid(RecebeCNPJ,7,1))
Numero(8) = Cint(Mid(RecebeCNPJ,8,1))
Numero(9) = Cint(Mid(RecebeCNPJ,9,1))
Numero(10) = Cint(Mid(RecebeCNPJ,10,1))
Numero(11) = Cint(Mid(RecebeCNPJ,11,1))
Numero(12) = Cint(Mid(RecebeCNPJ,12,1))
Numero(13) = Cint(Mid(RecebeCNPJ,13,1))
Numero(14) = Cint(Mid(RecebeCNPJ,14,1))

soma = Numero(1) * 5 + Numero(2) * 4 + Numero(3) * 3 + Numero(4) * 2 + Numero(5) * 9 + Numero(6) * 8 + Numero(7) * 7 + Numero(8) * 6 + Numero(9) * 5 + Numero(10) * 4 + Numero(11) * 3 + Numero(12) * 2

soma = soma -(11 * (int(soma / 11)))

if soma = 0 or soma = 1 then
resultado1 = 0
else
resultado1 = 11 - soma
end if
if resultado1 = Numero(13) then
soma = Numero(1) * 6 + Numero(2) * 5 + Numero(3) * 4 + Numero(4) * 3 + Numero(5) * 2 + Numero(6) * 9 + Numero(7) * 8 + Numero(8) * 7 + Numero(9) * 6 + Numero(10) * 5 + Numero(11) * 4 + Numero(12) * 3 + Numero(13) * 2
soma = soma - (11 * (int(soma/11)))
if soma = 0 or soma = 1 then
resultado2 = 0
else
resultado2 = 11 - soma
end if
if resultado2 = Numero(14) then
	CalculaCNPJ = true
else
	CalculaCNPJ = false
end if
else
	CalculaCNPJ = false
end if
end if
end function


function formatamoedav(num)
	if num <>"" then
	num = replace(num,",","")
	formatamoedav = replace(num,".",",")
	end if
end function

function formatamoedap(num)
	if num <>"" then
	num = replace(num,".","")
	formatamoedap = replace(num,",",".")
	end if
end function

function retiraacentos(texto)
 if texto <>"" then
	
	
	texto = replace(texto,"á","a")
	texto = replace(texto,"é","e")
	texto = replace(texto,"í","i")
	texto = replace(texto,"ó","o")
	texto = replace(texto,"ú","u")
	
	texto = replace(texto,"â","a")
	texto = replace(texto,"ê","e")
	texto = replace(texto,"î","i")
	texto = replace(texto,"ô","o")
	texto = replace(texto,"û","u")
	
	texto = replace(texto,"à","a")
	texto = replace(texto,"è","e")
	texto = replace(texto,"ì","i")
	texto = replace(texto,"ò","o")
	texto = replace(texto,"ù","u")
	
	texto = replace(texto,"ã","a")
	texto = replace(texto,"õ","o")
	
	texto = replace(texto,"ç","c")
	
	retiraacentos = texto
	end if
end function



	function ClearString(sString)
    		For i = 1 To Len(sString)
			
			If (Asc(Mid(sString, i, 1)) - 48) >= 0 And (Asc(Mid(sString, i, 1)) - 48) < 10 then
				sRet = sRet & Mid(sString, i, 1)
			else
				sRet = sRet & ""
			end if

				'response.write Mid(sString, i, 1) & "-" 

        		If Asc(Mid(sString, i, 1)) >= 65 And Asc(Mid(sString, i, 1)) <= 90 Then
            			sRet = sRet & Trim(CStr(Asc(Mid(sString, i, 1)) - 55))
        		End If
    		Next
    		ClearString = sRet
	end function

FUNCTION RemoveFormatacao(sTexto)
	if sTexto <> "" then
		for i = 1 to len(sTexto)
			a=mid(sTexto,i,1)
			if a >="0" and a <="9" then
				RemoveFormatacao=RemoveFormatacao & a		
			end if
		next		
	end if
END FUNCTION

FUNCTION Cript(Senha)

	if senha <> "" then
		Dim sText        
		Dim MAT(10)
		
		Dim sCriptograph
		Dim i  , x 
		
		i = 0
		x = 0 
		
		sText = senha
			
		MAT(1) = "M"
		MAT(2) = "A"
		MAT(3) = "Q"
			
		For i = 1 To Len(sText)
		x = Asc(Mid(sText, i, 1)) + Asc(MAT((i Mod 3) + 1))
		If x > 254 Then
			x = Asc(Mid(sText, i, 1)) - 23
			sCriptograph = sCriptograph & Chr(255)
			sCriptograph = sCriptograph & Chr(x)
		Else
		sCriptograph = sCriptograph & Chr(x)
		End If
			
		Next
		
		Cript = sCriptograph
	end if

end function

Function De_Criptograph(Senha)
	if senha <> "" then 
		Dim sAX, Y
		Dim sText        
		Dim MAT(10)
		
		Dim sCriptograph
		Dim i  , x 
		
		MAT(1) = "M"
		MAT(2) = "A"
		MAT(3) = "Q"
		
		sText = senha
		
		i = 1
		Y = 1
		
			While i <= Len(sText)
			  If Asc(Mid(sText, i, 1)) = 255 Then
			   i = i + 1
				sCriptograph = sCriptograph & Chr(Asc(Mid(sText, i, 1)) + 23)
				Y = Y + 1
				Else
				 x = Asc(Mid(sText, i, 1)) - Asc(MAT((Y Mod 3) + 1))
				 'response.Write x
				 sCriptograph = sCriptograph & Chr(x)
				 Y = Y + 1
			   End If
			  
			  i = i + 1
			Wend
		 'response.Write("<script>alert('"&sCriptograph&"');<script>")
		 De_Criptograph = sCriptograph
 	end if
End Function 


Function SeparaHora(data)
		pos=instr(1,data," ")
		SeparaHora=mid(data,pos+1)
end Function
	
Function EmailValido(ByVal strEmail)

	Dim regEx
	Dim ResultadoHum
	Dim ResultadoDois 
	Dim ResultadoTres
    Set regEx = New RegExp            ' Cria o Objeto Expressão
    regEx.IgnoreCase = True         ' Sensitivo ou não
    regEx.Global = True             ' Não sei exatamente o que faz 
    
    ' Caracteres Excluidos
    regEx.Pattern    = "[^@\-\.\w]|^[_@\.\-]|[\._\-]{2}|[@\.]{2}|(@)[^@]*\1"
    ResultadoHum    = RegEx.Test(strEmail)
    ' Caracteres validos
    regEx.Pattern    = "@[\w\-]+\."        
    ResultadoDois    = RegEx.Test(strEmail)
    ' Caracteres de fim
    regEx.Pattern    = "\.[a-zA-Z]{2,3}$"  
    ResultadoTres    = RegEx.Test(strEmail)
    Set regEx = Nothing
    
    If Not (ResultadoHum) And ResultadoDois And ResultadoTres Then
        EmailValido = True
    Else
        EmailValido = False
    End If

End Function



'------------------------------------------------------------------------------------------------
' Função: Extenso
' Converte um Valor Numérico entre 1 e 9.999.999.999.999,99 para seu Extenso em Reais
'	- Valores devem usar "vírgula" para Casas Decimais e "ponto" ou NÃO para Milhar
'	- Valores com mais de 2 Casas Decimais são arredondados (Ex. 2,496) = 2,50
'
' Recebe	: Valor 
' Retorna	: String com Extenso
'
' Faça um em sua rotina principal e nela use:
'	response.write Extenso(<ValorparaConversao>)
'
'------------------------------------------------------------------------------------------------

Dim x_Centavos, x_I, x_J, x_K, x_Numero, x_QtdCentenas, x_TotCentenas, x_TxtExtenso( 900 ), x_TxtMoeda( 6 ), x_ValCentena( 6 ), x_Valor, x_ValSoma

' Matrizes de textos
x_TxtMoeda( 1 ) = "rea"
x_TxtMoeda( 2 ) = "mil"
x_TxtMoeda( 3 ) = "milh"
x_TxtMoeda( 4 ) = "bilh"
x_TxtMoeda( 5 ) = "trilh"

x_TxtExtenso( 1 ) = "um"
x_TxtExtenso( 2 ) = "dois"
x_TxtExtenso( 3 ) = "tres"
x_TxtExtenso( 4 ) = "quatro"
x_TxtExtenso( 5 ) = "cinco"
x_TxtExtenso( 6 ) = "seis"
x_TxtExtenso( 7 ) = "sete"
x_TxtExtenso( 8 ) = "oito"
x_TxtExtenso( 9 ) = "nove"
x_TxtExtenso( 10 ) = "dez"
x_TxtExtenso( 11 ) = "onze"
x_TxtExtenso( 12 ) = "doze"
x_TxtExtenso( 13 ) = "treze"
x_TxtExtenso( 14 ) = "quatorze"
x_TxtExtenso( 15 ) = "quinze"
x_TxtExtenso( 16 ) = "dezesseis"
x_TxtExtenso( 17 ) = "dezessete"
x_TxtExtenso( 18 ) = "dezoito"
x_TxtExtenso( 19 ) = "dezenove"
x_TxtExtenso( 20 ) = "vinte"
x_TxtExtenso( 30 ) = "trinta"
x_TxtExtenso( 40 ) = "quarenta"
x_TxtExtenso( 50 ) = "cinquenta"
x_TxtExtenso( 60 ) = "sessenta"
x_TxtExtenso( 70 ) = "setenta"
x_TxtExtenso( 80 ) = "oitenta"
x_TxtExtenso( 90 ) = "noventa"
x_TxtExtenso( 100 ) = "cento"
x_TxtExtenso( 200 ) = "duzentos"
x_TxtExtenso( 300 ) = "trezentos"
x_TxtExtenso( 400 ) = "quatrocentos"
x_TxtExtenso( 500 ) = "quinhentos"
x_TxtExtenso( 600 ) = "seiscentos"
x_TxtExtenso( 700 ) = "setecentos"
x_TxtExtenso( 800 ) = "oitocentos"
x_TxtExtenso( 900 ) = "novecentos"

' Função Principal de Conversão de Valores em Extenso
Function Extenso( x_Numero )
x_Numero = FormatNumber( x_Numero , 2 )
x_Centavos = right( x_Numero , 2 )
x_ValCentena( 0 ) = 0
x_QtdCentenas = int( ( len( x_Numero ) + 1 ) / 4 )

For x_I = 1 to x_QtdCentenas
	x_ValCentena( x_I ) = "" 
Next
'
x_I = 1
x_J = 1
For x_K = len( x_Numero ) - 3 to 1 step -1
	x_ValCentena( x_J ) = mid( x_Numero , x_K , 1 ) & x_ValCentena( x_J )
	if x_I / 3 = int( x_I / 3 ) then
		x_J = x_J + 1
		x_K = x_K - 1
	end if
	x_I = x_I + 1
next
x_TotCentenas = 0
Extenso = "" 		
For x_I = x_QtdCentenas to 1 step -1
	
	x_TotCentenas = x_TotCentenas + int( x_ValCentena( x_I ) )

	if int( x_ValCentena( x_I ) ) <> 0 or ( int( x_ValCentena( x_I ) ) = 0 and x_I = 1 )then
		if int( x_ValCentena( x_I ) = 0 and int( x_ValCentena( x_I + 1 ) ) = 0 and x_I = 1 )then
			Extenso = Extenso & ExtCentena( x_ValCentena( x_I ) , x_TotCentenas ) & " de " & x_TxtMoeda( x_I )
		else
			Extenso = Extenso & ExtCentena( x_ValCentena( x_I ) , x_TotCentenas ) & " " & x_TxtMoeda( x_I )
		end if
		if int( x_ValCentena( x_I ) ) <> 1 or ( x_I = 1 and x_TotCentenas <> 1 ) then
			Select Case x_I
				Case 1
					Extenso = Extenso & "is"
				Case 3, 4, 5
					Extenso = Extenso & "ões"
			End Select		
		else
			Select Case x_I
				Case 1
					Extenso = Extenso & "l"
				Case 3, 4, 5
					Extenso = Extenso & "ão"
			End Select		
		end if
	end if
	if int( x_ValCentena( x_I - 1 ) ) = 0 then
		Extenso = Extenso
	else
		if ( int( x_ValCentena( x_I + 1 ) ) = 0 and ( x_I + 1 ) <= x_QtdCentenas ) or ( x_I = 2 ) then
			Extenso = Extenso & " e "
		else
			Extenso = Extenso & ", "
		end if
	end if		
next

if x_Centavos > 0 then
	if int( x_Centavos ) = 1 then
			Extenso = Extenso & " e " & ExtDezena( x_Centavos ) & " centavo"
	else
			Extenso = Extenso &  " e " & ExtDezena( x_Centavos ) & " centavos"
	end if
end if
Extenso = UCase( Left( Extenso , 1 ) )&right( Extenso , len( Extenso ) - 1 )
End Function

function ExtDezena( x_Valor )
' Retorna o Valor por Extenso referente à DEZENA recebida
ExtDezena = ""
if int( x_Valor ) > 0 then
	if int( x_Valor ) < 20 then
		ExtDezena = x_TxtExtenso( int( x_Valor ) )
	else
		ExtDezena = x_TxtExtenso( int( int( x_Valor ) / 10 ) * 10 )
		if ( int( x_Valor ) / 10 ) - int( int( x_Valor ) / 10 ) <> 0 then
			ExtDezena = ExtDezena & " e " & x_TxtExtenso( int( right( x_Valor , 1 ) ) )
		end if
	end if
end if
End Function

function ExtCentena( x_Valor, x_ValSoma )
ExtCentena = ""

if int( x_Valor ) > 0 then

	if int( x_Valor ) = 100 then
		ExtCentena = "cem"
	else
		if int( x_Valor ) < 20 then
			if int( x_Valor ) = 1 then
				If x_ValSoma - int( x_Valor ) = 0 then
					ExtCentena = "hum"
				else
					ExtCentena = x_TxtExtenso( int( x_Valor ) )
				end if
			else
					ExtCentena = x_TxtExtenso( int( x_Valor ) )
			end if
		else
			if int( x_Valor ) < 100 then
				ExtCentena = ExtDezena( right( x_Valor , 2 ) )
			else				
				ExtCentena = x_TxtExtenso( int( int( x_Valor ) / 100 )*100 )
				if ( int( x_Valor ) / 100 ) - int( int( x_Valor ) / 100 ) <> 0 then
					ExtCentena = ExtCentena & " e " & ExtDezena( right( x_Valor , 2 ) )
				end if
			end if
		end if
	end if
end if
End Function




Function EnCrypt(strCryptThis)

'Dim strChar, iKeyChar, iStringChar, i,g_Key

g_Key = mid("dscalkwc3ro2qr3t436/$%/$%&N5cdfcadscnsijsdifjsdoifjosdijf0ui09iu09sdi09i09sidf09sidf09si",1,Len(strCryptThis))

for i = 1 to Len(strCryptThis)

iKeyChar = Asc(mid(g_Key,i,1))

iStringChar = Asc(mid(strCryptThis,i,1))

' *** uncomment below to encrypt with addition,

' iCryptChar = iStringChar + iKeyChar

iCryptChar = iKeyChar Xor iStringChar

iCryptCharHex = Hex(iCryptChar)

iCryptCharHexStr = cstr(iCryptCharHex)

if len(iCryptCharHexStr)=1 then iCryptCharHexStr = "0" & iCryptCharHexStr

strEncrypted = strEncrypted & iCryptCharHexStr

next

EnCrypt = strEncrypted

End Function
 

Function DeCrypt(strEncrypted)

		Dim strChar, iKeyChar, iStringChar, i,g_Key
		
		g_Key = mid("dscalkwc3ro2qr3t436/$%/$%&N5cdfcadscnsijsdifjsdoifjosdijf0ui09iu09sdi09i09sidf09sidf09si",1,Len(strEncrypted)/2)
		
		for i = 1 to Len(strEncrypted) /2
		
		iKeyChar = (Asc(mid(g_Key,i,1)))
		
		iStringChar2 = mid(strEncrypted,(i*2)-1,2)
		
		iStringChar = CLng("&H" & iStringChar2)
		
		 
		
		' *** uncomment below to decrypt with subtraction
		
		' iDeCryptChar = iStringChar - iKeyChar
		
		iDeCryptChar = iKeyChar Xor iStringChar
		
		strDecrypted = strDecrypted & chr(iDeCryptChar)
		
		next
		
		DeCrypt = strDecrypted

End Function






%>

	
