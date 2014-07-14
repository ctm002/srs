<% 
    idSesionUsuario = Session("sesion_usuario_id")
%>
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=windows-1252"/>
    
    <script type="text/javascript">
        function cerrar() {
            top.document.location = "default.asp";
        }

        function abre_ventana() {
            xdireccion = "calendario_propio.asp?correlativo=<% response.write xcorrelativo%>";
            top.frames[2].document.location = xdireccion;
        }

        function abre_hoy() {
            var e = top.frames[0].document.getElementById("cboPisos");
            var idPiso = e.value;
            if (idPiso != "") {
                var url = "ver_dia.asp?piso=" + idPiso;
                top.frames[2].document.location = url;
            }
        }

        function abre_reserva(idUsuario) {
            var idPiso = top.frames[0].document.getElementById("cboPisos").value;
            if (idPiso != "") {
                var cboAbogados = top.frames[0].document.getElementById("cboAbogados");
                var idUsuario;
                if (typeof cboAbogados === "undefined") {
                    idUsuario = cboAgogados.value;
                } else {
                    idUsuario = "<%=idSesionUsuario%>";
                }
                
                var url = "crear.asp?piso=" + idPiso;
                top.frames[2].document.location = url;
            }
        }
    </script>
    <meta name="ProgId" content="FrontPage.Editor.Document" />
    <title>Sistema de Reserva</title>
</head>
<body bgcolor="#45496F" topmargin="0" leftmargin="8" rightmargin="0" bottommargin="0"
    marginwidth="0" marginheight="0">
    <table border="0" align="center" width="100%" id="table2" cellpadding="3">
        <tr>
            <td background="./img/barra_menu.jpg" height="10">
                <font size="2" color="#FFFFFF"><a style="color: #FFFFFF; font-family: Verdana; font-size: 10pt"
                        href="javascript:abre_hoy();""><span style="font-family: Verdana; text-decoration: none">
                        Hoy</span></a></font>
            </td>
        </tr>
        <tr>
            <td background="./img/barra_menu.jpg" height="10">
                <font size="2" color="#FFFFFF"><a style="color: #FFFFFF; font-family: Verdana; font-size: 10pt"
                    href="javascript:abre_reserva();"><span style="font-family: Verdana; text-decoration: none">
                    Reservar Sala</span></a></font>
            </td>
        </tr>
        <tr>
            <td background="./img/barra_menu.jpg">
                <a style="color: #FFFFFF; font-family: Verdana; font-size: 10pt" href="javascript:abre_ventana();">
                    <font size="2" color="#FFFFFF"><span style="font-family: Verdana; text-decoration: none">
                        Mis Reuniones</span></font></a>
            </td>
        </tr>
        <tr>
            <td background="./img/barra_menu.jpg">
                <font size="2" color="#FFFFFF"><a style="color: #FFFFFF; font-family: Verdana; font-size: 10pt"
                    href="javascript:cerrar();"><span style="font-family: Verdana; text-decoration: none">
                        Salir</span></a></font>
            </td>
        </tr>
    </table>
</body>
</html>
