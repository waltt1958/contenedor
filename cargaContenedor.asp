<HTML>
<HEAD>

<meta charset="utf-8">
<link rel="stylesheet" title="estilos.css" type="text/css" href="estilos.css">
<title>CONTENEDORES ROSARIO</title>

</HEAD>

<body onload="maximizar()">
<br>

<H5>Hoy es: <%=weekdayname(weekday(date()))%>, <%=date%></H5>
<h1>SECTOR INGRESOS Y RENDICIONES DE D.E.</h1>

<hr size= 6 color="black"></hr>

<br>
<h3> GENERACION DE CONTENEDOR</H3>
<br>

<table align="center">
<td align="center"><input type="button" class="button" name="volver" onclick=location.href='index.asp' value="VOLVER"></td>
<tr align="center">
<td></td>
<td></td>
</tr>
</table>



<SCRIPT Language="javascript" type="text/javascript">

function maximizar() {

window.moveTo(0,0);

window.resizeTo(screen.width,screen.height);
}
</SCRIPT>


</body>

</HTML>