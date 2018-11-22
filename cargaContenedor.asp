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

<!--#include virtual="/conectar.asp"-->
<%

set rsRangos=server.createobject("ADODB.Recordset") 
sql= "select * from RANGOS"

Set rsContador = server.createobject("ADODB.Recordset")
sqlcuenta= "select * from RANGOS"
rsContador.open sqlcuenta, conectarOEP

if (rsContador.EOF = false and rsContador.BOF = false) then

rsRangos.open sql, conectarOEP 


do while not rsRangos.EOF

	menor = rsRangos("rango_desde")
	mayor = rsRangos("rango_hasta")

	CP= rsRangos("CP")

	 do while menor < mayor
		incluir = 0 & Cstr(menor)
		conectarOEP.execute "INSERT INTO contenedor(CUIP, CP) VALUES('"&incluir&"','"&CP&"')"
		menor = menor + 1
	 loop 
	
	menor= 0
	mayor = 0

	rsRangos.MoveNext

loop

rsRangos.close

set rsCuenta = server.createobject("ADODB.Recordset")	
set rsArchivo = server.createobject("ADODB.Recordset")
sqlcta = "select count (*) AS cuenta from contenedor"
sqlarc ="select * from contenedor"

rsCuenta.open sqlcta, conectarOEP
rsArchivo.open sqlarc, conectarOEP

espacio10 = "          "
espacio12= "        "
actual = now()

nombre = "Contenedor dia " & day(actual) & "-" & month(actual) & "-" & year(actual) & "  "& hour(actual) & "-" & Minute(actual) & "-" & Second(actual) & ".txt"

Set fso = Server.CreateObject ("Scripting.FileSystemObject")

Set arcTEXTO = fso.CreateTextFile(server.mappath(nombre), true)

mostrar = len(rsARCHIVO("CUIP"))

if mostrar = 10 then

	do while not rsARCHIVO.EOF

		texto = rsARCHIVO.Fields("CUIP") & espacio10 & rsARCHIVO.Fields("CP")
		arcTEXTO.WriteLine(texto)
		rsARCHIVO.MoveNext

	loop

else

	do while not rsARCHIVO.EOF

		texto = rsARCHIVO.Fields("CUIP") & espacio12 & rsARCHIVO.Fields("CP")
		arcTEXTO.WriteLine(texto)
		rsARCHIVO.MoveNext

	loop
	
end if

Set fso = nothing
Set arcTEXTO = nothing



session("registros")= rsCuenta("cuenta")

session("nombreArc")= nombre

rsCuenta.close

rsArchivo.close



%>

<!--#include virtual="/desconectar.asp"-->

<table align="center" style="font-size:20px" border="3" cellspacing=0 bordercolor="black" width="55%" height="10%">
<tr>

<td align="center" bgcolor="#E6E6FA"><b><u>Se generó el archivo: <%=response.write(nombre) %> y contiene <%= response.write(session("registros")) %></u></b></td>

</tr>
</table>
<br>
<br>

<table align="center">
<tr>
<td>
<input type="button" class="button" name="genera" onclick=location.href='bajaArchivo.asp' value="BAJAR ARCHIVO">
</td>
</tr>

<tr>
<td height="10">
               
</td>
</tr>
</table>

<table align="center">
<tr>
<td align="center">
<input type="button" class="button" name="inicial" onclick=location.href='index.asp' value="VOLVER AL INICIO">
</td>
</tr>
</table>



<%
else

	response.redirect ("cargaRango.asp?target=_self")


end if

%>
</body>

</HTML>