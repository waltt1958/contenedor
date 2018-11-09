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
response.write (menor)
response.write (mayor)

do while not rs.EOF

if (menor = rs("rango_desde")) or (mayor = rs("rango_hasta")) then
	
	session("repetido")= "si"
	response.redirect ("error.asp?target=_self")
	
end if

rs.MoveNext

loop

end if


%>

<!--#include virtual="/desconectar.asp"-->
</body>
</html>