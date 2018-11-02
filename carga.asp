<html>
<head>
</head>
CARGAR LA BBDD Y LUEGO VOLVER A cargaRango.asp PARA SEGUIR CARGANDO
<body>



<%
primero	= trim(request.form("primerNUM"))
segundo = trim(request.form("segundoNUM"))

if ((not IsNumeric(trim(primero)) or not isNumeric(trim(segundo))) or (len(primero) <> 10 or len(segundo) <> 10)) then

response.redirect("error.asp?target=_self")

else
%>
<!--#include virtual="/conectar.asp"-->
<%







end if
%>

<!--#include virtual="/desconectar.asp"-->
</body>
</html>