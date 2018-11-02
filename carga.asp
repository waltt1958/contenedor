<html>
<head>
</head>
CARGAR LA BBDD Y LUEGO VOLVER A cargaRango.asp PARA SEGUIR CARGANDO
<body>

<!--#include virtual="/conectar.asp"-->

<%
primero	= request.form("primerNUM")
segundo = request.form("segundoNUM")

if (not IsNumeric(trim(primero)) or not isNumeric(trim(segundo))) then

response.redirect("cargaRango.asp?target=_self")



end if

'usar trim() dado que elimina los espacios en blanco al principio y final
'TAMBIEN CONTROLAR QUE SEAN 10 LOS DIGITOS, EL ESPACIO NO LO CUENTA COMO TEXTO




%>

<!--#include virtual="/desconectar.asp"-->
</body>
</html>