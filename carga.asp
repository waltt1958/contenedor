<html>
<head>
</head>

<body>


<%
primero	= trim(request.form("primerNUM"))
segundo = trim(request.form("segundoNUM"))
cp= request.form("select")

if ((not IsNumeric(trim(primero)) or not isNumeric(trim(segundo))) or (len(primero) <> 10 or len(segundo) <> 10)) then

response.redirect("error.asp?target=_self")

else
%>
<!--#include virtual="/conectar.asp"-->
<%

codigo= request.form("select")

set rs=server.createobject("ADODB.Recordset") 
sql= "select * from RANGOS"
rs.open sql, conectarOEP 


PRIMERO VERIFICAR QUE NO VAYA A ESTAR DUPLICADO EL NUMERO CARGADO
Y LUEGO SE LO AGREGA A LA BBDD

If primero > segundo Then
    mayor = primero
	menor = segundo

End If



if rs!rango_desde = menor or rs!rango_hasta = mayor then 
	
session("repetido")= "si"




rs.AddNew

        Do While Val(menor) < Val(mayor)
                    
            rs3.AddNew
            If Left([menor], 1) = 0 Then
            rs3!cuip = menor
            Else
            rs3!cuip = 0 & menor
            End If
            rs3!suc = rs1!SUCURSAL
            rs3.Update
           
           menor = Val(menor) + 1
         
        Loop
        
        rs1.MoveNext
        mayor = ""
        menor = ""
        
    Loop


%>

<!--#include virtual="/desconectar.asp"-->
</body>
</html>