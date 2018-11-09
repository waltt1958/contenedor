<html>
<head>

<meta charset="utf-8">
<link rel="stylesheet" title="estilos.css" type="text/css" href="estilos.css">
<title>CONTENEDORES ROSARIO</title>

</head>

<body>
<H5>Hoy es: <%=weekdayname(weekday(date()))%>, <%=date%></H5>
<h1>SECTOR INGRESOS Y RENDICIONES DE D.E.</h1>
<hr size= 6 color="black"></hr>

<% if session("repetido") <> "si" then %>
<h2 align="center"><b> VUELVA A CARGAR EL ULTIMO RANGO DESDE Y HASTA DADO QUE NO FUE TOMADO POR ERROR EN LA CANTIDAD DE DIGITOS (deben ser 10 numeros) 
O PORQUE HABIA LETRAS<h2>

<% 
		else
%>
<h2 align="center"><b> ALGUNO DE LOS RANGOS ESTA REPETIDO. CONTROLE Y VUELVA A CARGAR<h2>

<% end if %>
<a href="cargaRango.asp" target="_self"><input type="button" class="button" name="volver" value="VOLVER"></a>

</body>
</html>