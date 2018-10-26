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
<br>
<hr size= 6 color="black"></hr>

<br>



<br>

<hr size= 6 color="black"></hr>

<br>
<br>

<table align="center">
<tr align="center">
<td><input type="button" class="button" name="iniciar" onclick=location.href='cargaArchivo.asp' value="CARGAR RANGOS"></td>
<td><input type="button" class="button" name="contenedor" onclick=location.href='cargaArchivo.asp' value="GENERAR CONTENEDOR"></td>
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