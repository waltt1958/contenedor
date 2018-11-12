<html>
<head>
</head>

<body>

<%
primero	= trim(request.form("primerNUM"))
segundo = trim(request.form("segundoNUM"))
cp= request.form("select")

if ((not IsNumeric(trim(primero)) or not isNumeric(trim(segundo))) or (len(primero) <> 10 or len(segundo) <> 10)) then

response.redirect("error.asp?target=_self")

else
%>
<!--#include virtual="/conectar.asp"-->
<%

set rs=server.createobject("ADODB.Recordset") 
sql= "select * from RANGOS"
rs.open sql, conectarOEP 

menor = primero
mayor = segundo

If primero > segundo Then
    
	mayor = primero
	menor = segundo
	
End If

do while not rs.EOF

if (menor >= rs("rango_desde") and menor =< rs("rango_hasta")) or (mayor >= rs("rango_desde") and mayor =< rs("rango_hasta")) then

%>
<!--#include virtual="/desconectar.asp"-->
<%
	
	session("repetido")= "si"
	response.redirect ("error.asp?target=_self")
	
end if

rs.MoveNext

loop

conectarOEP.execute "INSERT INTO RANGOS(CP,rango_desde, rango_hasta) VALUES('"&CP&"','"&menor&"', '"&mayor&"')"

end if

%>

<!--#include virtual="/desconectar.asp"-->
<%
response.redirect ("cargaRango.asp?target=_self")
%>
</body>
</html>