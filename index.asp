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

<h3> ELIJA LA OPCIÓN QUE CORRESPONDA</H3>
<br>
<form action="cargaRango.asp" name="formulario" id="formulario" method="post">

<input type="hidden" name="oculto">

<table cellspacing="8" align="center">

<tr>
<td align="center" colspan="2">
<input type="radio" name="diez" id="diez" value="diez" checked>
<label for="diez"><b>CB 10 DIGITOS:</b> </label>
</td>
</tr>

<tr>
<td align="center" colspan="2">
<input type="radio" name="diez" id="doce" value="doce">
<label for="doce"><b>CB 12 DIGITOS: </b></label>
</td>
</tr>

<tr align="center">

<td><input type="button" class="button" name="iniciar" onclick=location.href='cargaRango.asp' value="CARGAR RANGOS"></td>
<td><input type="button" class="button" name="contenedor" onclick=location.href='cargaContenedor.asp' value="GENERAR CONTENEDOR"></td>


</tr>
</table>
</form>

<%
session("repetido")= "no"
%>




<SCRIPT Language="javascript" type="text/javascript">

function maximizar() {

window.moveTo(0,0);

window.resizeTo(screen.width,screen.height);
}
</SCRIPT>


</body>

</HTML>