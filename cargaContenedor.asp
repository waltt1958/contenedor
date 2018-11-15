<HTML>
<HEAD>

<title>CONTENEDORES ROSARIO</title>

</HEAD>

<body>

<!--#include virtual="/conectar.asp"-->
<%

set rsRangos=server.createobject("ADODB.Recordset") 
sql= "select * from RANGOS"
rsRangos.open sql, conectarOEP 


do while not rsRangos.EOF

menor = rsRangos("rango_desde") + 1
mayor = rsRangos("rango_hasta")
CP= rsRangos("CP")

	 do while menor < mayor
		conectarOEP.execute "INSERT INTO contenedor(CUIP, CP) VALUES('"&menor&"','"&CP&"')"
		menor = menor + 1
	 loop 
	
menor= 0
mayor = 0

rsRangos.MoveNext

loop

rsRangos.close

set rsCuenta = server.createobject("ADODB.Recordset")	
set rsArchivo = ser.createobject("ADODB.Recordset")
sqlcta = "select count (*) AS cuenta from contenedor"
sqlarc ="select * from contenedor"

rsCuenta.open sqlcta, conectarOEP
rsArchivo.open sqlarc, conectarOEP

actual = now()

nombre = "Contenedor dia " & day(actual) & "-" & month(actual) & "-" & year(actual) & "  "& hour(actual) & "-" & Minute(actual) & "-" 
& Second(actual) & ".txt"

Set fso = Server.CreateObject ("Scripting.FileSystemObject")

Set arcTEXTO = fso.CreateTextFile(server.mappath(nombre), true)

 texto1 = rsARCHIVO.Fields(0).name & "|" & rsARCHIVO.Fields(7).name & "|" & rsARCHIVO.Fields(1).name & "|" & rsARCHIVO.Fields(8).name & "|" & rsARCHIVO.Fields(9).name & "|" & _
  rsARCHIVO.Fields(10).name & "|" & rsARCHIVO.Fields(2).name & "|" & rsARCHIVO.Fields(3).name & "|" & rsARCHIVO.Fields(4).name & "|" & rsARCHIVO.Fields(11).name & "|" & _
  rsARCHIVO.Fields(12).name & "|" & rsARCHIVO.Fields(13).name & "|" & rsARCHIVO.Fields(14).name & "|" & rsARCHIVO.Fields(15).name & "|" & rsARCHIVO.Fields(16).name _
  & "|" & rsARCHIVO.Fields(17).name & "|" & rsARCHIVO.Fields(18).name & "|" & rsARCHIVO.Fields(19).name & "|" & rsARCHIVO.Fields(20).name & "|" & _
  rsARCHIVO.Fields(21).name & "|" & rsARCHIVO.Fields(22).name & "|" & rsARCHIVO.Fields(23).name & "|" & rsARCHIVO.Fields(24).name & "|" & _
  rsARCHIVO.Fields(25).name & "|" & rsARCHIVO.Fields(26).name & "|" & rsARCHIVO.Fields(27).name & "|" & rsARCHIVO.Fields(5).name & "|" & rsARCHIVO.Fields(28).name _
  & "|" & rsARCHIVO.Fields(6).name & "|" & rsARCHIVO.Fields(29).name & "|" & rsARCHIVO.Fields(30).name & "|" & rsARCHIVO.Fields(31).name




' USAR ESTA VARIABLE DE SESSION EN LA PAGINA QUE BAJA EL ARCHIVO PARA INDICAR LA CANTIDAD DE REGISTROS
'OJO RESTRINGIR EL INGRESO A LA PAGINA, AGREGAR VARIABLE SESSION

session("registros")= rsCuenta("cuenta")

rsCuenta.close
rsArchivo.close




%>

<!--#include virtual="/desconectar.asp"-->

</body>

</HTML>