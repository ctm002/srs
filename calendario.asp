<!--#include file="parametros_globales.asp"-->
<!--#include file="consultas.asp"-->
<%
    set fp_rs = Server.CreateObject("ADODB.Recordset")
    sesion_usuario_tipo = Session("sesion_usuario_tipo")
    sesion_usuario_id = Session("sesion_usuario_id")
    xcorrelativo= 1984 

Main
function ingresa_cero(xnumero)
	numero=trim(cstr(xnumero))
	largo=len(numero)
	salida=string(2-largo,"0") + numero
	ingresa_cero=salida
end function

function busca_festivo(xday,xmes,xyear)
    busca_festivo=0
    fecha_busqueda=xyear+ ingresa_cero(xmes) +ingresa_cero(xday)
    fp_sQry="Select * from "+ GLOBAL_DB_CARIOLA + " TS_FESTIVOS where Fecha='" + fecha_busqueda + "'"
    fp_rs.Open fp_sQry, GLOBAL_DSN    
    if fp_rs.eof and fp_rs.bof then
        salida=0
    else
        salida=1
    end if

    fp_rs.close
    busca_festivo=salida
end function

function busca_reunion(xday,xmes,xyear)
    fp_sQry="Select * from "+ GLOBAL_DB_ADM + GLOBAL_TABLA_AGENDA +" where datepart(dd,fecha)=" + cstr(xday) +" and datepart(mm,fecha)=" + cstr(xmes) +" and datepart(YY,fecha)="+ cstr(xyear) +"  and IdUsuarioAutor=" + cstr(xcorrelativo)
    fp_rs.Open fp_sQry,GLOBAL_DSN
    
    if fp_rs.eof and fp_rs.bof then
        salida=0
    else
        salida=1
    end if

    fp_rs.close
    busca_reunion=salida
end function

Sub Main
    matriz=array("","Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre")
	Global_Mes=Request.Form("XMES")
	Global_Year=Request.Form("XYEAR")

	if Global_Mes="" then Global_Mes=Request.Querystring("XMES")
	if Global_Year="" then Global_Year=Request.Querystring("XYEAR")

	if Global_Mes="" then Global_Mes=cstr(month(date))
	if Global_Year="" then Global_Year=cstr(year(date))

    '''El formato de fecha debe ser cambiado dependiendo del Idioma del sistema operativo
	xfecha=Global_Mes+"/"+"01/"+Global_Year

    '''El formato de fecha debe ser cambiado dependiendo del Idioma del sistema operativo
	xfirst=weekday(Cdate(xfecha))-1

	if xfirst=0 then
        xfirst=7	
    end if	
%>
<html lang="es-cl">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
    <meta http-equiv="Content-Language" content="es" />
<!--    <meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />-->

    <link rel="stylesheet" href="./jquery/css/jquery.ui.all.css" />
	<script src="./jquery/jquery-1.10.2.js" type="text/javascript"></script>
	<script src="./jquery/ui/jquery.ui.core.js" type="text/javascript"></script>
	<script src="./jquery/ui/jquery.ui.widget.js" type="text/javascript"></script>
	<script src="./jquery/ui/jquery.ui.datepicker.js" type="text/javascript"></script>
    <link rel="stylesheet" href="./css/demos.css" />
    <script type="text/javascript">
        $.datepicker.regional['es'] = {
            closeText: 'Cerrar',
            prevText: 'Ant',
            nextText: 'Sig',
            currentText: 'Hoy',
            monthNames: ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'],
            monthNamesShort: ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'],
            dayNames: ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'],
            dayNamesShort: ['Dom', 'Lun', 'Mar', 'Mié', 'Juv', 'Vie', 'Sáb'],
            dayNamesMin: ['Do', 'Lu', 'Ma', 'Mi', 'Ju', 'Vi', 'Sá'],
            weekHeader: 'Sm',
            dateFormat: 'dd/mm/yy',
            firstDay: 1,
            isRTL: false,
            showMonthAfterYear: false,
            yearSuffix: ''
        };
        $.datepicker.setDefaults($.datepicker.regional['es']);

        function cambiar_fecha(xyear, xmes, xdia) {
            var espera = top.trabajo.document.getElementById("subcontent2");
            if (espera != null) {
                espera.style.display = "block";
            }

            Direccion = "ver_dia.asp?xyear=" + xyear + "&xmes=" + xmes + "&xdia=" + xdia;
            parent.trabajo.location.href = Direccion;
        }

        function onClick() {
            var url;
            var fecha = $('#datepicker').val();
            var cboPisos = document.getElementById("cboPisos");
            var strNroPiso = cboPisos.value;
            var cboAbogados = document.getElementById("cboAbogados");
            if (cboAbogados != null) {
                url = "ver_dia.asp?piso=" + strNroPiso + "&fecha=" + fecha + "&id_usuario_creador=" + cboAbogados.value;
            } else {
                var id_usuario_creador = "<%=sesion_usuario_id%>";
                url = "ver_dia.asp?piso=" + strNroPiso + "&fecha=" + fecha + "&id_usuario_creador=" + id_usuario_creador;
            }
            parent.trabajo.location.href = url;
        }

        $(function () {
            $("#datepicker").datepicker({
                onSelect: function (){
                    onClick();
                }
            });
        });

        function onChange()
        {
            onClick();
        }
	</script>

    <style type="text/css">
        .cell_over
        {
            background-color: #FFFF99;
        }
        .cell_out
        {
            background-color: #FFFFFF;
        }
        .cell_over1
        {
            background-color: #C9E4FC;
        }
        .cell_out1
        {
            background-color: #FFFFCC;
        }
        .cell_out2
        {
            background-color: #FFFFCC;
        }
    </style>
    <title>Calendario</title>
</head>
<body>
<table>
    <tr>
        <td>
            <div id="datepicker"></div>
        </td>
    </tr>
    <tr>
        <td>
            <select id="cboPisos" style="width:193px" onchange="onChange();">
                <option value="9" selected="selected">Piso 19</option>
                <option value="10">Piso 24</option>
            </select>
        </td>
    </tr>
</table>
</body>
</html>
<%end sub%>