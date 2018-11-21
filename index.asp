<HTML>
<HEAD>

<meta charset="utf-8">
<link rel="stylesheet" title="estilos.css" type="text/css" href="estilos.css">
<title>CONTENEDORES ROSARIO</title>

</HEAD>

<body onload="maximizar()">

<H5>Hoy es: <%=weekdayname(weekday(date()))%>, <%=date%></H5>
<h1>SECTOR INGRESOS Y RENDICIONES DE D.E.</h1>
<h3>GENERACION DE CONTENEDORES</h3>
<hr size= 6 color="black"></hr>

<br>

<h3> ELIJA LA OPCIÓN QUE CORRESPONDA</H3>
<br>

<table cellspacing="8" align="center">

<tr align="center">

<td><input type="button" class="button" name="iniciar" onclick=location.href='cargaRango.asp' value="CARGAR RANGOS"></td>
<td><input type="button" class="button" name="contenedor" onclick=location.href='cargaContenedor.asp' value="GENERAR CONTENEDOR"></td>

</tr>
</table>
<script>
<%
texto= "txt"
bbdd= "mdb"
clasico= "asp"
forma= "css"
imagen = "png"

Set objFSO = server.CreateObject ("Scripting.FileSystemObject")
set objFolder=objFSO.GetFolder("c:\inetpub\wwwroot\contenedor\")

for each objFile in objFolder.files

Select case objFSO.GetExtensionName(objFile)
case bbdd
case clasico
case forma
case imagen
case else

objFile.delete
end select

next


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