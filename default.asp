    <!--#include file="consultas.asp"-->
    <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
    <html>
    <head>
        <meta http-equiv="Content-Language" content="es" />
        <meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
        <title>.:Login Reserva Salas:.<%Response.write now %></title>
    </head>
    <body bgcolor="#CCCCCC">
        <%
            nameUser = ""
            if Request.QueryString("USR") <> "" then
                nameUser = Request.QueryString("USR")
            end if
            
            if Request.Form("USR") <> "" then
                nameUser = Request.Form("USR")
            end if
            
            dim xNameUser : xNameUser =  Request.ServerVariables("LOGON_USER")
            dim xNameUserTmp : xNameUserTmp = Split(xNameUser,"\")
            nameUser = "Carlos_Tapia" 'xNameUserTmp(1) 
            if nameUser <> "" Then 
                
                Set objUser = GetUsuarioByNameUser(nameUser)
                If (Not objUser is nothing) Then
                    idUsuario = objUser.item("IdUsuario")

                    'Preguntar si tiene acceso del sistema de reserva
                    set objUserAutorizado =  AutorizarByIdUsuario(idUsuario)
                    if (Not objUserAutorizado is nothing) then
                        
                        'Tiempo de duracion de la sesion
                        Session.Timeout = 30
                        Session("sesion_usuario_id") = objUserAutorizado.item("IdUsuario")
                        Session("sesion_usuario_nombres") = objUserAutorizado.item("Nombres") + " " + objUserAutorizado.item("Apellidos")
                        Session("sesion_usuario_tipo") = "ADMIN"
                        Response.Redirect("adm_salas.asp")
                    else
                        
                        'Preguntamos si el tipo de usuario logueado es abogado
                        set objAbogado = GetAbogadoByIdUsuario(idUsuario)
                        If (Not objAbogado is Nothing) then

                            Session("sesion_usuario_id") = idUsuario
                            Session("sesion_usuario_nombres") = objAbogado.item("Nombres") + " " + objAbogado.item("Apellidos")
                            Session("sesion_usuario_tipo") = "ABOGADO"
                            Response.Redirect("adm_salas.asp")
                        else
                            
                            'Preguntar si es secretaria
                            set objSecretaria = GetSecretariaByIdUsuario(idUsuario)
                            if (Not objSecretaria is nothing) then 

                                Session("sesion_usuario_id") = idUsuario
                                Session("sesion_usuario_nombres") = objSecretaria.item("Nombres") + " " + objSecretaria.item("Apellidos")
                                Session("sesion_usuario_tipo") = "SECRETARIA"
                                Response.Redirect("adm_salas.asp")
                            else
                                'Response.Redirect("sin_acceso.htm")
                            end if
                        end if

                    end if 
                Else
                        Response.Redirect("sin_acceso.htm")
                End if
            End if
        %>
        <form action="default.asp" method="post">
        <div style="border: 0px solid; width: 100%; height: 100%">
            <table border="0" cellpadding="0" cellspacing="0" style="width: 800px; height: 600px;
                border-color: #517DBF; margin: 0 auto 0 auto">
                <tr style="height: 10%">
                    <td colspan="2" align="center">
                        <h1>
                            Sistema de Reserva de Salas</h1>
                    </td>
                </tr>
                <tr style="height: 90%">
                    <td>
                        <p align="center">
                            <img border="0" src="./img/logo.jpg" width="120px" height="168px" alt="" /></p>
                    </td>
                    <td>
                        <div align="center" style="border: 1px solid; height: 168px; border-radius: 20px">
                            <table border="0" cellpadding="0" cellspacing="0" style="width: 70%; border-collapse: collapse"
                                bordercolor="#111111">
                                <tr>
                                    <td colspan="2" align="center">
                                        <label for="LblLogin">
                                            Login de Usuario:</label>
                                    </td>
                                </tr>
                                <tr>
                                    <td width="30%">
                                        <p align="right">
                                            <b><font face="Verdana" size="2">Nombre de Usuario:</font></b></p>
                                    </td>
                                    <td width="70%" align="left">
                                        <input type="text" name="USR" size="25" style="font-family: Verdana; font-size: 10pt;
                                            height: 20px" />
                                    </td>
                                </tr>
                                <tr>
                                    <td width="30%">
                                        <p align="right">
                                            <b><font face="Verdana" size="2">Contraseña:</font></b></p>
                                    </td>
                                    <td width="70%" align="left">
                                        <input type="password" name="PWD" size="25" style="font-family: Verdana; font-size: 10pt;
                                            height: 20px" />
                                    </td>
                                </tr>
                                <tr>
                                    <td width="40%">
                                        &nbsp;
                                    </td>
                                    <td width="60%" align="left">
                                        <input type="submit" value="Ingresar" name="B1" style="background: #4285F4; color: White;
                                            height: 27px; border: 1px solid rgba(0, 0, 0, 0.1);" />
                                    </td>
                                </tr>
                            </table>
                        </div>
                    </td>
                </tr>
            </table>
        </div>
        </form>
    </body>
    </html>
