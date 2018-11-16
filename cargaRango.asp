<HTML>
<HEAD>

<meta charset="utf-8">
<link rel="stylesheet" title="estilos.css" type="text/css" href="estilos.css">
<title>CONTENEDORES ROSARIO</title>

</HEAD>


<body >

<!--#include virtual="/conectar.asp"-->

<H5>Hoy es: <%=weekdayname(weekday(date()))%>, <%=date%></H5>
<h1>SECTOR INGRESOS Y RENDICIONES DE D.E.</h1>

<hr size= 6 color="black"></hr>

<h3> CARGA DE RANGOS</H3>
<form action="carga.asp" method="POST" id="cargador" name="cargador">
<table align="center" border="2">

<tr>

<td>
<b>SELECCIONE LA SUCURSAL</b> 
</td>

<td>
<Select NAME="select" style="width: 300px;"  SIZE=1 autofocus>

<%


session ("repetido")= "no"
Set rs = Server.CreateObject("ADODB.recordset")
sql= "select * from CP_SUCURSALES order by SUCURSAL"
rs.open sql, conectarOEP


Do While not rs.EOF


response.write ("<option value='" & rs("CP")&"'>" & rs("SIGLA") & " - " & rs("SUCURSAL") & "</option>" )

rs.MoveNext

Loop
rs.Close

%>
</Select>

</td>

</tr>

<tr>

<td>
<b>INGRESE EL 1er CODIGO DE BARRAS</b>
</td>

<td>
<input type="text" name="primerNUM" id="primerNUM" maxlength="10" required>
</td>

</tr>

<tr>

<td>
<b>INGRESE EL 2do CODIGO DE BARRAS</b>
</td>

<td>
<input type="text" name="segundoNUM" id="segundoNUM" maxlength="10" onblur="enviar()" required>
</td>

</tr>
</table>

</form>
<br>
<br>

<table align="center">
<tr align="center">
<td align="center"><input type="button" class="button" name="volver" onclick=location.href='index.asp' value="VOLVER"></td>
<td align="center"><input type="button" class="button" name="contenedor" onclick=location.href='cargaContenedor.asp' value="GENERAR CONTENEDOR"></td>
</tr>
</table>

<%
session("repetido")= "no"
response.write(request.form("diez"))
response.write("SSSSSSSSSSSSSSSSSSSSSSSSSSSSSS")
%>

<SCRIPT Language="javascript" type="text/javascript">
function enviar() {

var primero = document.cargador.primerNUM.value.length;
var segundo = document.cargador.segundoNUM.value.length;

if (( primero != 10) || (segundo != 10)) {
	
	alert ("Verique CB ingresado. Tiene que tener 10 digitos");
	document.cargador.primerNUM.focus();
	}
	
	else {
	document.cargador.submit()
	}
	}


</SCRIPT>

<!--#include virtual="/desconectar.asp"-->

</body>

</HTML>