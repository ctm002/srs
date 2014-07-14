<!--#include file="consultas.asp"-->
<%
    session_usuario_nombre = Session("sesion_usuario_nombres")
    session_usuario_id = Session("sesion_usuario_id")
    session_usuario_tipo = Session("sesion_usuario_tipo")
    id_usuario_creador = Request.QueryString("id_usuario_creador")

    global_id=0
	global_desde=""
	global_hasta=""
	global_titulo=""
	global_hora=""
	global_usuarios = ""

    'Recuperamos las salas del piso
    global_piso = Iff(Request.QueryString("piso") = "" , "9" , Request.QueryString("piso"))
    Dim salas: set salas = getSalasByIdPiso(global_piso)
    Numero_Salas = salas.Count
    
    fecha = Request.QueryString("fecha")
    if fecha = "" then
        xday=cstr(day(date))
        xmonth=cstr(month(date))
        xyear=cstr(year(date))
    else
        xday=cstr(day(fecha))
        xmonth=cstr(month(fecha))
        xyear=cstr(year(fecha))
    end if
    
    xfecha = xyear + "-" + xmonth + "-" + xday
    matriz = array("","Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre")
    matriz_semana = array("DOMINGO","LUNES","MARTES","MIERCOLES","JUEVES","VIERNES","SABADO")

'Llamado al procedimiento Main
Main

function Iff(condicion, verdadero, falso)
    if (condicion) then 
        Iff= verdadero
    else
        Iff = falso
    end if 
end function

function ingresa_cero(xnumero)
	numero=trim(cstr(xnumero))
	largo=len(numero)
	salida=string(2-largo,"0") + numero
	ingresa_cero=salida
end function

Sub Main()
    dia_semana=matriz_semana(weekday(xfecha,2) mod 7)
%>
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=Utf-8" />
    <title>Detalle Dia</title>
    <script type="text/javascript">
        function Redirect(url) {
            document.location = url;
        }
    </script>
    <style type="text/css">
        .fondo_amarillo
        {
            background-color: #FFFF66;
            font-weight: bold;
            font-family: Verdana;
            font-size: 8pt;
            color: #000000;
        }
        .fondo_blanco
        {
            color: #FFFFFF;
            font-weight: bold;
            font-family: Verdana;
            font-size: 8pt;
        }
    </style>
</head>
<body bgcolor="#CCCCCC" topmargin="1" leftmargin="5" rightmargin="1" bottommargin="1"
    marginwidth="1" marginheight="1">
    <div align="center" id="subcontent2" style="position: absolute; display: none; left: 250;
        top: 250">
        <div style="border: 1px solid black; background-color: lightyellow; width: 200px;
            height: 80px; padding: 2px">
            <p align="center">
                <font face="Verdana" size="2" color="#003399"><b>
                    <br />
                    <br />
                    Buscando reservas
                    <br />
                    espere un momento ...<br />
                    <img src="./img/espera.gif" alt=""/></b></font></p>
        </div>
    </div>
    <div align="center">
        <center>
            <font color="#cccccc" size="1"><span lang="es">.</span></font>
            <table id="table94" style="border-collapse: collapse" bordercolor="#517dbf" cellspacing="0"
                width="100%" border="1">
                <tr>
                    <td valign="top" width="100%" bgcolor="#ffffff" height="250">
                        <div align="center">

                            <!--Formulario de ingreso-->
                            <form name="doublecombo" method="post" action="grabar_edit_reserva.asp">
                            <input type="hidden" value="" name="id" />
                            <center>
                                <table id="table95" style="border-collapse: collapse" bordercolor="#111111" cellspacing="0"
                                    cellpadding="0" width="99%" border="0">
                                    <tr>
                                        <td align="middle" width="100%">
                                        </td>
                                    </tr>
                                    <tr>
                                        <td align="middle" width="100%">
                                            <font face="Verdana" color="#000000" size="1">**Reunion propia</font><img src="./img/cita.gif" alt="" /><font
                                                face="Verdana" color="#000000" size="1">*Sala Ocupada</font><img src="./img/ocupado.gif" alt="" />
                                        </td>
                                    </tr>
                                    <tr>
                                        <td width="25%">
                                            <div align="center">
                                                <table id="table101" style="border-collapse: collapse" bordercolor="#111111" cellspacing="0"
                                                    cellpadding="4" width="100%" bgcolor="#cccccc" border="1">
                                                    <tr>
                                                        <td align="left" width="60%" bgcolor="#517dbf">
                                                            <b><i><font face="Verdana" color="#FFFF00" size="2">
                                                                <% Response.write dia_semana %>&nbsp<% response.write xday %>&nbsp de &nbsp<% response.write matriz(xmonth)%> &nbsp de &nbsp<% response.write xyear %>
                                                            </font></i></b>
                                                        </td>
                                                        <td align="left" width="40%" bgcolor="#517dbf">
                                                            <b><i><font face="Verdana" color="#ffffff" size="2">
                                                                <% response.write session_usuario_nombre %></font></i></b>
                                                        </td>
                                                    </tr>
                                                </table>
                                            </div>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td width="25%">
                                            <div align="center">
                                                <table id="table104" style="border-collapse: collapse" bordercolor="#111111" cellspacing="0"
                                                    cellpadding="4" width="100%" bgcolor="#cccccc" border="1">
                                                    <tr>
                                                        <td align="middle" width="35" bgcolor="#517dbf">
                                                            <i><font face="Verdana" size="2" color="#FFFFFF">Hora</font></i>
                                                        </td>
                                                        <% For Each key in salas 
                                                           Dim fila : set fila = salas.item(key)
                                                        %>
                                                        <td align="middle" background="./img/barra_menu.jpg" bgcolor="#517dbf">
                                                            <i><font face="Verdana" size="2" color="#FFFFFF">
                                                                <%
                                                                    salida = "Sala " & fila.item("nro") & "<br/>" & fila.item("nombre")
                                                                    Response.Charset = "UTF-8"
                                                                    Response.Write salida
                                                                %>
                                                            </font></i>
                                                        </td>
                                                        <% Next %>
                                                    </tr>
                                                    <% 
                                                    color1="#ebf0f9"
                                                    color2="#FFFFFF"

                                                    'Horario de atencion de las salas 7:00 a 22:00 hrs
                                                    for i = 7 to 22
                                                        if color= color1 then
                                                            color=color2
                                                        else
                                                            color=color1
                                                        end if
                                                    %>
                                                    <tr>
                                                        <td bgcolor="#517dbf" width="35">
                                                            <font face="Verdana" size="2" color="#FFFFFF">
                                                                <% response.write ingresa_cero(i)%>:00</font>
                                                        </td>
                                                        <% For each key in salas 
                                                            dim link_celda
                                                            dim dato_celda

                                                            dim f : set f = salas.item(key)
                                                            call getReservaByIdSalaAndHora(f.item("id"), i, link_celda, dato_celda,session_usuario_id)
                                                        %>

                                                        <td style="cursor : hand;" bgcolor="<% response.write color%>" onmouseover="this.className='fondo_amarillo'" 
                                                            onmouseout="this.className='fondo_blanco'" class="fondo_blanco" onclick="javascript:Redirect('<%=link_celda%>');">
                                                            <%=dato_celda%>
                                                        </td>
                                                        
                                                        <% next %>
                                                     </tr>
                                                     <% next %>
                                                </table>
                                            </div>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td width="100%">
                                            &nbsp;
                                        </td>
                                    </tr>
                                </table>
                            </center>
                            </form>
                        </div>
                    </td>
                </tr>
            </table>
        </center>
    </div>
    <font size="1">
        <% response.write Request.ServerVariables("REMOTE_ADDR")%>
    </font>
</body>
</html>
<%
end sub
%>