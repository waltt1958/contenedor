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
<!--#include virtual="/conectar.asp"-->
<%

set rsRangos=server.createobject("ADODB.Recordset") 
sql= "select * from RANGOS"
rsRangos.open sql, conectarOEP 

do while not rsRangos.EOF

	menor = rsRangos("rango_desde")
	mayor = rsRangos("rango_hasta")

	CP= rsRangos("CP")

	 do while menor < mayor
		incluir = Cstr(menor)
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

actual = now()

nombre = "Contenedor dia " & day(actual) & "-" & month(actual) & "-" & year(actual) & "  "& hour(actual) & "-" & Minute(actual) & "-" & Second(actual) & ".txt"

Set fso = Server.CreateObject ("Scripting.FileSystemObject")

Set arcTEXTO = fso.CreateTextFile(server.mappath(nombre), true)

do while not rsARCHIVO.EOF

	texto = rsARCHIVO.Fields("CUIP") & &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp & rsARCHIVO.Fields("CP")
	arcTEXTO.WriteLine(texto)
	rsARCHIVO.MoveNext

loop
	
Set fso = nothing
Set arcTEXTO = nothing



session("registros")= rsCuenta("cuenta")

session("nombreArc")= nombre

rsCuenta.close

rsArchivo.close



%>

<!--#include virtual="/desconectar.asp"-->


<table align="center">
<tr>
<td>
<a href="/bajaArchivo.asp" target="_self"><input type="button" name="descarga" value="DESCARGAR ARCHIVO" style="FONT-SIZE: 20pt; border: 5px solid; [b]FONT-FAMILY: Verdana, boldt[/b];
BACKGROUND-COLOR: #C0C0C0"></a>
</td>
</tr>

</table>

</body>

</HTML>