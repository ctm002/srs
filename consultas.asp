    <!--#include file="parametros_globales.asp"-->
    <%
        LoggedIn = Session("sesion_usuario_id") 
        if(LoggedIn <> "") then
            action = Request.Form("Action")
            select case action 
                case "GetSalasByIdPiso"
                    id = Request.Form("Id")
                    if (id <> "") then
                        set collSalas = getSalasByIdPiso(id)
                    
                        for each key in collSalas
                            set objSala = collSalas.item(key)

                            jsonSala = ""
                            for each f in objSala
                                dato = """" & f & """" & ":" & """" & objSala.item(f) & """"
                                jsonSala = jsonSala & dato & ","
                            next
                            jsonSala = left(jsonSala, len(jsonSala) -1)
                            jsonSalas = jsonSalas & "{" & jsonSala & "},"
                         next

                        jsonSalas = left(jsonSalas, len(jsonSalas) - 1)
                        jsonSalas = "{""d""" & ":[" & jsonSalas & "]}"
                    
                        Response.ContentType = "application/json"
                        Response.Write jsonSalas
                    end if
            end select
        else
            Response.Redirect("default.asp")
        end if

        function getSalasByIdPiso(id)
            Dim salas : Set salas = Server.createObject("Scripting.Dictionary")
            if (id <> "") then
                set xfp_rs = CreateObject("ADODB.Recordset")
                query = "Select * from " + GLOBAL_DB_ADM + GLOBAL_TABLA_SALA + " where fk_id_piso=" & id
                xfp_rs.Open query, GLOBAL_DSN

                If Not xfp_rs.eof then

                    While Not xfp_rs.eof

                        Dim fila : Set fila = Server.createObject("Scripting.Dictionary")
                
                        For each f in xfp_rs.Fields
                            fila.Add f.Name , f.Value
                        Next

                        salas.add fila.item("id") , fila
                        xfp_rs.movenext
                    Wend

                End if
                xfp_rs.close
                set xfp_rs = nothing
            end if
           Set getSalasByIdPiso = salas
        end function

        function GetUsuarios()
               Dim usuarios : Set usuarios = Server.createObject("Scripting.Dictionary")
               sql = "select u.* from " + GLOBAL_DB_ADM + "ADM_Usuario u, " + GLOBAL_DB_ADM + " ADM_Usuario_Perfil up, " + GLOBAL_DB_ADM + "ADM_Perfil p where " & _
                     "u.IdUsuario = up.IdUsuario and up.IdPerfil = p.IdPerfil  and " & _ 
                     "p.IdSistema= 36" '36 es el Sistema de reserva de salas
               'response.write sql
               set rs_usuarios = Server.CreateObject("ADODB.Recordset")
               rs_usuarios.open sql, GLOBAL_DSN
               
               while(not rs_usuarios.eof and not rs_usuarios.bof) 
                    set usuario = Server.CreateObject("Scripting.Dictionary")
                    for each f in rs_usuarios.Fields
                        usuario.add f.Name, f.Value
                    next
                    usuarios.add usuario.item("IdUsuario"), usuario
                    rs_usuarios.MoveNext
               wend
               
               rs_usuarios.close
               Set rs_usuarios = Nothing
               Set GetUsuarios = usuarios
           end function
            
            function GetPisos()
                Dim pisos : Set pisos = Server.createObject("Scripting.Dictionary")
                set rs_pisos = CreateObject("ADODB.Recordset")
                'query = "Select * from " + GLOBAL_DB_ADM + GLOBAL_TABLA_PISO
                query = "select * from " + GLOBAL_DB_ADM + "PER_Piso p where  exists (select * from " + GLOBAL_DB_ADM + "SRS_Sala s where s.fk_id_piso = p.pk_id_per_piso)"
                rs_pisos.open query, GLOBAL_DSN
                while not rs_pisos.BOF and not rs_pisos.EOF
                    Dim piso : set piso = Server.createObject("Scripting.Dictionary")
                    for each f in rs_pisos.Fields
                        piso.add f.Name , f.Value
                    next
                    pisos.add piso.item("pk_id_per_piso") , piso
                    rs_pisos.MoveNext
                wend
                rs_pisos.close
                set rs_pisos = nothing
                set GetPisos = pisos
            end function

            Function AutorizarByIdUsuario(Id)
                fp_sQry = "SELECT Administracion..ADM_Usuario.*, Administracion..ADM_Sistema.IdSistema IdSistema FROM  Administracion..ADM_Usuario INNER JOIN Administracion..ADM_Usuario_Perfil" & _
                    " ON Administracion..ADM_Usuario.IdUsuario = Administracion..ADM_Usuario_Perfil.IdUsuario INNER JOIN" & _
                    " Administracion..ADM_Perfil ON Administracion..ADM_Usuario_Perfil.IdPerfil = Administracion..ADM_Perfil.IdPerfil INNER JOIN " & _
                    " Administracion..ADM_Sistema ON Administracion..ADM_Perfil.IdSistema = Administracion..ADM_Sistema.IdSistema "
                fp_sQry = fp_sQry & "WHERE (Administracion..ADM_Sistema.IdSistema = 36) AND Administracion..ADM_Usuario.IdUsuario='" & Id &"' ORDER BY Nombres ASC" 

                Dim fila : Set fila = Server.createObject("Scripting.Dictionary")
                set fp_rs = CreateObject("ADODB.Recordset")
                fp_rs.Open fp_sQry, GLOBAL_DSN

                if Not fp_rs.EOF and Not fp_rs.BOF Then
                    For each f in fp_rs.Fields
                        fila.Add f.Name , f.Value
                    Next
                    fp_rs.close
                    Set AutorizarByIdUsuario = fila
                Else
                    fp_rs.close
                    Set AutorizarByIdUsuario = Nothing
                End If
            End Function

            Function GetUsuarioByNameUser(NameUser)
                sql = "Select u.* from Administracion..ADM_Usuario u where u.LoginName='" & NameUser & "'"
                set fp_rs = CreateObject("ADODB.Recordset")
                fp_rs.Open sql, GLOBAL_DSN
                
                if Not fp_rs.BOF Then
                    Dim fila : Set fila = Server.createObject("Scripting.Dictionary")
                    For each f in fp_rs.Fields
                        fila.Add f.Name , f.Value
                    Next
                    fp_rs.close
                    Set GetUsuarioByNameUser = fila
                Else
                    fp_rs.close
                    Set GetUsuarioByNameUser = Nothing
                End If
            End Function

            function busca_reunion(xday,xmes,xyear, idusuarioAutor, byref descReserva)
                'Response.Write xday & "-" & xmes & "-" & xyear
                salida = ""
                if(idusuarioAutor <> "")then
                     set fp_rs = CreateObject("ADODB.Recordset")
                     fp_sQry = "SELECT dbo.SRS_ReservaSalas.correlativo as corr, dbo.PER_Piso.pk_id_per_piso as IdPiso, " & _
                        "dbo.SRS_ReservaSalas.fecha, dbo.SRS_ReservaSalas.sala IdSala, dbo.SRS_ReservaSalas.serial_inicio, " & _
                        "dbo.SRS_ReservaSalas.serial_termino, " & _
                        "dbo.PER_Piso.descripcion as NroPiso, " & _
                        "dbo.SRS_Sala.nro NroSala" & _
                    " FROM dbo.SRS_ReservaSalas INNER JOIN" & _
                        "  dbo.SRS_Sala ON dbo.SRS_ReservaSalas.sala = dbo.SRS_Sala.id INNER JOIN"  & _
                        "  dbo.PER_Piso ON dbo.SRS_Sala.fk_id_piso = dbo.PER_Piso.pk_id_per_piso" & _
                    " WHERE (DATEPART(dd, dbo.SRS_ReservaSalas.fecha) =" & cstr(xday) &")" & _
                        " AND (DATEPART(mm, dbo.SRS_ReservaSalas.fecha) =" & cstr(xmes) & ")" & _
                        " AND (DATEPART(YY, dbo.SRS_ReservaSalas.fecha) =" & cstr(xyear) & ") and dbo.SRS_ReservaSalas.IdUsuarioAutor=" & idusuarioAutor
                    fp_sQry = replace(fp_sQry,"dbo.",GLOBAL_DB_ADM)
                    'Response.Write "query->" & fp_sQry
                    fp_rs.Open fp_sQry, GLOBAL_DSN
            
                    if fp_rs.eof and fp_rs.bof then
                        salida = false
                        descReserva = ""
    	            else
                        descReserva=""
                
                        while not fp_rs.eof
                            global_desde = FormatoHora(fp_rs("serial_inicio"))
                            global_hasta = FormatoHora(fp_rs("serial_termino"))
                            referencia="editar.asp?id=" + cstr(fp_rs("corr")) + "&fecha=" + cstr(day(fp_rs("fecha"))) + "/" + cstr(month(fp_rs("fecha"))) & _
                                "/" + cstr(year(fp_rs("fecha"))) + "&hora=" + ingresa_cero(hour(fp_rs("fecha"))) + "&piso=" + cstr(fp_rs("IdPiso")) & _
                                "&sala=" + cstr(fp_rs("IdSala"))
                    
                            descReserva= descReserva + "<br><a style=""text-decoration: none"" href=""" & _
                                referencia & """>" & global_desde & "-" & global_hasta & " S." + cstr(fp_rs("NroSala")) + " P." + fp_rs("NroPiso") +"</a>"
                            descReserva = descReserva & "<hr />"
                            fp_rs.movenext
                        wend 

                        salida=true
                     end if
                     fp_rs.close
                     set fp_rs = nothing
                else
                    salida = false
                end if
                busca_reunion = salida
            end function

            function GetSecretariaByIdUsuario(Id)
               sql = "SELECT Administracion.dbo.ADM_Usuario.Nombres, Administracion.dbo.ADM_Usuario.Apellidos, Administracion.dbo.ADM_Usuario.IdUsuario" & _
                    " FROM  Administracion.dbo.ADM_Usuario INNER JOIN " & _
                    " cariola.dbo.TS_SECRETARIA ON Administracion.dbo.ADM_Usuario.IdUsuario = cariola.dbo.TS_SECRETARIA.IdUsuario " & _
                    " WHERE Administracion.dbo.ADM_Usuario.IdUsuario =" & Id
                set fp_rs = Server.CreateObject("ADODB.Recordset")
                fp_rs.Open sql, GLOBAL_DSN
                'Response.Write sql
                if Not fp_rs.EOF Then

                    Dim fila : Set fila = Server.createObject("Scripting.Dictionary")
                    For each f in fp_rs.Fields
                        fila.Add f.Name , f.Value
                    Next
                    fp_rs.close
                    Set GetSecretariaByIdUsuario = fila
                Else
                    fp_rs.close
                    Set GetSecretariaByIdUsuario = Nothing
                End If
            end function

            function GetAbogadoByIdUsuario(Id)
                sql = "SELECT cariola.dbo.TS_ABOGADO.abo_id, Administracion.dbo.ADM_Usuario.Nombres, Administracion.dbo.ADM_Usuario.Apellidos, Administracion.dbo.ADM_Usuario.IdUsuario " & _
                    " FROM cariola.dbo.TS_ABOGADO LEFT OUTER JOIN " & _
                    " Administracion.dbo.ADM_Usuario ON cariola.dbo.TS_ABOGADO.IdUsuario = Administracion.dbo.ADM_Usuario.IdUsuario " & _
                    " WHERE Administracion.dbo.ADM_Usuario.IdUsuario =" & Id
                set fp_rs = CreateObject("ADODB.Recordset")
                fp_rs.Open sql, GLOBAL_DSN
                'Response.Write sql
                if Not fp_rs.BOF Then
                    Dim fila : Set fila = Server.createObject("Scripting.Dictionary")
                    For each f in fp_rs.Fields
                        fila.Add f.Name , f.Value
                    Next
                    fp_rs.close
                    Set GetAbogadoByIdUsuario = fila
                Else
                    fp_rs.close
                    Set GetAbogadoByIdUsuario = Nothing
                End If
            end function

            function GetAbogadosByIdSecretaria(id)
                sql = " SELECT cariola.dbo.TS_ABOGADO.abo_id, Administracion.dbo.ADM_Usuario.IdUsuario," & _
                      " Administracion.dbo.ADM_Usuario.Nombres, Administracion.dbo.ADM_Usuario.Apellidos " & _
                      " FROM  Administracion.dbo.ADM_Usuario RIGHT OUTER JOIN " & _
                       " cariola.dbo.TS_SECRETARIA LEFT OUTER JOIN " & _
                       " cariola.dbo.TS_ABOGADO ON cariola.dbo.TS_SECRETARIA.sec_id = cariola.dbo.TS_ABOGADO.sec_id ON  " & _
                       " Administracion.dbo.ADM_Usuario.IdUsuario = cariola.dbo.TS_ABOGADO.IdUsuario " & _
                       " WHERE cariola.dbo.TS_SECRETARIA.IdUsuario =" & id
                'Response.Write sql
                set fp_rs = Server.CreateObject("ADODB.Recordset")
                fp_rs.Open sql, GLOBAL_DSN
                
                if Not fp_rs.EOF Then
                    Dim lstAbogados : Set lstAbogados = Server.createObject("Scripting.Dictionary")
                    while Not fp_rs.EOF And Not fp_rs.BOF 
                        Dim abogado : Set abogado = Server.createObject("Scripting.Dictionary")
                        For each f in fp_rs.Fields
                            abogado.Add f.Name , f.Value
                        Next
                        'Response.Write abogado.item("IdUsuario") & "<br />"
                        if (abogado.item("IdUsuario") <> "" ) then
                            lstAbogados.add abogado.item("IdUsuario") , abogado
                        end if
                        fp_rs.movenext
                    Wend
                    fp_rs.close
                    Set fp_rs = Nothing
                    Set GetAbogadosByIdSecretaria = lstAbogados
                Else
                    fp_rs.close
                    Set fp_rs = Nothing
                    Set GetAbogadosByIdSecretaria = Nothing
                End If
            end function

            sub getReservaByIdSalaAndHora(xsala,xhora, byref p_link, byref p_celda, id_sesion_usuario)
                salida_buscar = ""
                serial_hora = left(xhora,2) + "00"
                serial_hora_fin = left(xhora,2) + "59"
        
                fp_sQry = "Select * from " + GLOBAL_DB_ADM + GLOBAL_TABLA_AGENDA +" where datepart(dd,fecha)=" + xday +" and datepart(mm,fecha)=" + xmonth +" and datepart(YY,fecha)="+ xyear 
                fp_sQry = fp_sQry+ " and ((serial_inicio>=" + serial_hora +"  and serial_termino<=" + serial_hora_fin + ") or (" + serial_hora +"  between serial_inicio and serial_termino)) and sala=" + cstr(xsala)
                'Response.Write(fp_sQry)

                'No tengo idea para que sirve este if
                'Carlos RTM
                if Request.ServerVariables("REMOTE_ADDR")="10.1.11.201" then
                end if

                color_fondo1 = "#FFCC66"
                color_fondo2 = "#669900"
                color_fondo = color_fondo2
                
                set fp_rs = Server.CreateObject("ADODB.Recordset")
                fp_rs.Open fp_sQry, GLOBAL_DSN
                url = ""
                salida_celda = ""

                if fp_rs.eof and fp_rs.bof then
                    url = "fecha=" + cstr(xday) + "/"  + cstr(xmonth) + "/" + cstr(xyear)+ "&sala=" + cstr(xsala) + "&hora="+ ingresa_cero(xhora) + ":00&piso=" + global_piso
                    link = "crear.asp?" & url
                    global_desde = ""
                    global_hasta = ""
                    global_titulo = ""
                    global_id = 0
                    global_hora = ""
                    global_usuarios = ""
                else

                    tiene_reserva = true
                    anterior = ""

                    global_id = fp_rs("correlativo")
                    global_transac = fp_rs("transaccion")

                    if global_transac <> anterior then
                        global_usuarios = ""
                        if (salida_celda <> "") then
                            salida_celda = salida_celda+"<br><hr width=""150"">"
                        end if
                    end if

                    'fecha de inicio
                    caracter = ":"
                    serial_inicio = fp_rs("serial_inicio")
                    if (len(serial_inicio) = 4) then
                        global_desde = ingresa_cero(mid(serial_inicio,1,2) ) + caracter + ingresa_cero(mid(serial_inicio,3,2))
                    else
                        global_desde = ingresa_cero(mid(serial_inicio,1,1)) + caracter + ingresa_cero(mid(serial_inicio,2,2))
                    end if

                    'fecha de termino
                    serial_termino = fp_rs("serial_termino")
                    if (len(serial_termino) = 4) then
                        global_hasta = ingresa_cero(mid(serial_termino,1,2) ) + caracter + ingresa_cero(mid(serial_termino,3,2))
                    else
                        global_hasta = ingresa_cero(mid(serial_termino,1,1)) + caracter + ingresa_cero(mid(serial_termino,2,2))
                    end if

                    global_titulo = fp_rs("titulo")
                    texto_salida = global_titulo
                    global_hora = ingresa_cero(hour(fp_rs("fecha")))
                    global_usuarios = global_usuarios + chr(10)
                    global_usuarios = buscar_usuario(fp_rs("correlativo"),id_sesion_usuario)

                    if (anterior <> global_transac) then
                        salida_celda="<font face=""verdana"" color=""#808080"" size=""1"">" + salida_celda + global_desde + " - " + global_hasta +  "</font><br>" 
                    end if 

                    salida_celda = salida_celda + global_usuarios
                    salida_celda = salida_celda +"<div style=""border: 1px solid black; width: 70%;background-color:FFCC66; padding: 2px""><font face=""verdana"" color=""#000000"" size=""1"">"+ texto_salida + "</font></div>"
            
                    'url = "id=" + cstr(global_id) +"&xday=" + cstr(xday) + "&xmes="  + cstr(xmonth) + "&xyear=" + cstr(xyear) + "&piso=" + cstr(global_piso) + "&sala=" + cstr(xsala) + "&hora=" + cstr(xhora)
                    url = "id=" + cstr(global_id) +"&fecha=" + cstr(xday) +"/"+ cstr(xmonth) +"/"+ cstr(xyear) + "&piso=" + cstr(global_piso) + "&sala=" + cstr(xsala) + "&hora=" + cstr(xhora)
                    link = "editar.asp?"+ url
                end if

                fp_rs.close
                p_celda = salida_celda 
                p_link = link
            end sub

            function buscar_fecha(xsala,xhora, session_usuario_id)
                serial_hora=left(xhora,2) + "00"
                serial_hora_fin=left(xhora,2) + "59"
                fp_sQry="Select * from " + GLOBAL_DB + GLOBAL_TABLA_AGENDA +" where datepart(dd,fecha)=" + xday +" and datepart(mm,fecha)=" + xmonth +" and datepart(YY,fecha)="+ xyear 
                fp_sQry=fp_sQry+" and ((serial_inicio>=" + serial_hora +"  and serial_termino<" + serial_hora_fin + ") or (" + serial_hora +"  between serial_inicio and serial_termino)) and sala=" + cstr(xsala) 
                Direccion_Base ="http://webserver/agenda/ver_date.asp?xday=" + xday + " &xmes="+ xmonth+"&xyear=" + xyear +" &sala=" + cstr(xsala) + "&hora=" 

                if Request.ServerVariables("REMOTE_ADDR")="10.1.11.201" then
                    response.write fp_sQry
                end if

                fp_rs = Server.CreateObject("ADODB.Recordset")
                fp_rs.Open fp_sQry,GLOBAL_DSN
                if fp_rs.eof and fp_rs.bof then
                    buscar_fecha="nada.gif"
                    global_desde=""
                    global_hasta=""
                    global_titulo=""
                    global_id=0
                    global_hora=""
                    global_usuarios=""
                else

                    global_id = fp_rs("correlativo")
                    global_desde = ingresa_cero(cstr(hour(fp_rs("fecha"))))+":" + ingresa_cero(cstr(minute(fp_rs("fecha"))))
                    global_hasta = ingresa_cero(cstr(hour(fp_rs("fecha_termino"))))+":" + ingresa_cero(cstr(minute(fp_rs("fecha_termino"))))
                    global_titulo = fp_rs("TITULO")
                    buscar_fecha = "ocupado.gif"
                    global_hora = ingresa_cero(hour(fp_rs("fecha")))
                    global_usuarios = chr(10) + "Participantes : "

                    while not fp_rs.eof
                        global_usuarios = global_usuarios + chr(10) + "*" + buscar_usuario(cdbl(fp_rs("usuario")))
                        if cdbl(fp_rs("IdUsuarioAutor"))=cdbl(session_usuario_id) then
                            buscar_fecha="cita.gif"
                        end if
                        fp_rs.movenext
                    wend 

                end if

                fp_rs.close
            end function

            function buscar_usuario(codigo_correlativo, session_usuario_id)
                xfp_sQry="Select u.* from " + GLOBAL_DB_ADM + GLOBAL_TABLA_USUARIOS_RESERVA + " rs," + GLOBAL_DB_ADM + GLOBAL_TABLA_USUARIOS + " u where rs.IdUsuario=u.IdUsuario and rs.correlativo = " + cstr(codigo_correlativo)
                set xfp_rs = Server.CreateObject("ADODB.Recordset")
                xfp_rs.Open xfp_sQry, GLOBAL_DSN

                if xfp_rs.eof and xfp_rs.bof then
                    buscar_usuario = ""
                else

                    nombres_completo = ""
                    listado_nombres = ""

                    while not xfp_rs.eof
                        if cdbl(xfp_rs("IdUsuario"))=cdbl(session_usuario_id) then
                            nombres_completo = "<img border=""0"" src=""./img/ocupado.gif"" />" & xfp_rs("Nombres") + " " + xfp_rs("Apellidos") + "<br />"
                        else
                             nombres_completo = "<img border=""0"" src=""./img/cita.gif"" />" & xfp_rs("Nombres") + " " + xfp_rs("Apellidos") + "<br />"
                        end if
                        listado_nombres = listado_nombres + nombres_completo
                        xfp_rs.movenext
                    wend
                    buscar_usuario = listado_nombres
                end if
                xfp_rs.close
                set xfp_rs = nothing
            end function
    %>
