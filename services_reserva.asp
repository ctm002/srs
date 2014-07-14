<!--#include file="parametros_globales.asp"-->
<!--#include file="utils.asp"-->
<%
	LoggedIn = Session("sesion_usuario_id")	
	if(LoggedIn <> "") then
		tipoSesionUsuario = Session("sesion_usuario_tipo")
		op = Request.Form("op")
		select case op
			case "buscar"
                sala = Request.Form("sala")
                fecha = FormatDateString(Request.Form("fecha"),1)
                hora_inicio = replace(Request.Form("desde"),":","")
                hora_termino = replace(Request.Form("hasta"),":","")
                id_reserva = Request.Form("id_reserva")
                set objReserva = BuscarReservaByFechaAndHoraInicioTermino(id_reserva, sala, fecha, hora_inicio, hora_termino)
                respuesta = "0"
                if (not objReserva is nothing) then
                    respuesta = cstr(objReserva.item("correlativo"))
                end if
                Response.ContentType = "application/json"
                Response.Write "{""message"":""" & respuesta & """}"
			case "crear"
				crear
			case "editar"
				editar
			case "eliminar"
                respuesta = eliminar()
                Response.ContentType = "application/json"
                Response.Write "{""message"":""" + respuesta + """}"
            case "puedeEditar"
                Response.ContentType = "application/json"
                respuesta = PuedeEditarReserva(Request.Form("id_reserva"), LoggedIn)
                if(respuesta = true) then 
                    respuesta =  "true"
                else 
                    respuesta = "false"
                end if                
                Response.Write "{""message"":""" & respuesta & """}"
        end select
	else
		Response.Redirect("default.asp")
	end if

	sub ver()

	end sub

	sub crear()
        on Error Resume Next       
       	fecha = FormatDateString(Request.Form("txtFecha"),1)
 		usuario_creador = iff(tipoSesionUsuario = "SECRETARIA", Request.Form("cboAbogados"), LoggedIn)	
 		titulo = Request.Form("txtSub") 
        actividad = Request.Form("txtAct") 
        tipo_actividad = Request.Form("tipo_act") 
        avisar_por_email = Request.Form("aviso_email") 
        sala = Request.Form("cboSalas") 
        hora_inicio = replace(Request.Form("txtHoraDesde"),":","")
        hora_termino = replace(Request.Form("txtHoraHasta"),":","")
        participantes = split(Request.Form("cboParticipantes"),",")
        
        Set oConn = Server.CreateObject("ADODB.Connection")
        oConn.Open GLOBAL_DSN
        oConn.BeginTrans
        insert1 = "SET NOCOUNT ON; Insert " + GLOBAL_DB_ADM + GLOBAL_TABLA_AGENDA + "(fecha,IdUsuarioAutor,titulo,actividad,tipo_actividad,aviso_email,fecha_termino,sala,serial_inicio,serial_termino,transaccion)" & _
            + "values('" + cstr(fecha) + "'," + cstr(usuario_creador) + ",'" + titulo + "','" & _
            + actividad + "','" + tipo_actividad + "','" + avisar_por_email + "','" & _ 
            + cstr(fecha) + "',"  + sala + "," + hora_inicio + "," + hora_termino + ",'" + ToTransacion(fecha, sala, hora_inicio, hora_termino) + "');" & _
            " SET NOCOUNT OFF; SELECT SCOPE_IDENTITY() ID;"
        dim rs : set rs = oConn.Execute(insert1)
        id_reserva = rs(0)
        insert2 = "SET NOCOUNT ON;"
        for i = 0 to ubound(participantes)
            insert2 = insert2 & "Insert " + GLOBAL_DB_ADM + GLOBAL_TABLA_USUARIOS_RESERVA & _
                + "(correlativo,idUsuario) values (" + cstr(id_reserva) + "," + cstr(participantes(i)) + ");" 
        next
        insert2 = insert2 '& " SET NOCOUNT OF;"
        oConn.Execute(insert2)        
        rs.close
        rs = nothing
        oConn.CommitTrans
        If Err.Number <> 0 Then
            oConn.RollBackTrans
        End If
        oConn.close
        set oConn = nothing
    	url = "ver_dia.asp?piso=" + Request.Form("cboPisos") + "&fecha=" + FormatDateString(Request.Form("txtFecha"),2)
  		Response.Redirect(url)
	end sub

	sub editar()
		id_reserva = Request.Form("IdReserva") 
		fecha = FormatDateString(Request.Form("txtFecha"),1)
 		usuario_creador = iff(tipoSesionUsuario = "SECRETARIA", Request.Form("cboAbogados"), LoggedIn)
        'Response.write usuario_creador
 		titulo = Request.Form("txtSub") 
        actividad = Request.Form("txtAct") 
        tipo_actividad = Request.Form("tipo_act") 
        avisar_por_mail = Request.Form("aviso_email") 
        sala = Request.Form("cboSalas") 
        hora_inicio = replace(Request.Form("txtHoraDesde"),":","")
        hora_termino = replace(Request.Form("txtHoraHasta"),":","")
        participantes = split(Request.Form("cboParticipantes"),",")
        
        call editarReserva(id_reserva, fecha,hora_inicio, hora_termino, titulo,actividad,tipo_actividad,avisar_por_mail, sala, participantes, usuario_creador)
		url = "ver_dia.asp?piso=" + Request.Form("cboPisos") + "&fecha=" + FormatDateString(Request.Form("txtFecha"),2)
		Response.Redirect(url)
	end sub

    function PuedeEditarReserva(id_reserva,id_secretaria)
        PuedeEditarReserva = false
        set reserva = GetReservaById(id_reserva)
        if (not reserva is nothing) then

            if(cstr(reserva.item("IdUsuarioAutor")) = cstr(id_secretaria)) then
                PuedeEditarReserva = true
            else
                set abogados = GetAbogadosByIdSecretaria(id_secretaria)
                if (not abogados is nothing) then
                    for each key in abogados
                        set abogado = abogados.item(key)
                        'Response.write(reserva.item("IdUsuarioAutor") & "="  & abogado.item("IdUsuario")) & "<br/>"
                        if (cstr(reserva.item("IdUsuarioAutor")) = cstr(abogado.item("IdUsuario"))) then 
                            PuedeEditarReserva = true
                        end if
                    next
                end if 
            end if
        end if
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

	sub editarReserva(idReserva, fecha, hora_inicio, hora_termino, titulo, actividad, tipo_actividad, avisar_por_mail, sala, participantes_actuales, usuario_creador)
        on Error Resume Next
        Dim participantes_viejos()

        Set oConn = Server.CreateObject("ADODB.Connection")
        oConn.Open GLOBAL_DSN
        oConn.BeginTrans
        sql1 = "Select * from " & GLOBAL_DB_ADM & GLOBAL_TABLA_USUARIOS_RESERVA  & " where correlativo=" & cstr(idReserva)
        set rs = Server.CreateObject("ADODB.RecordSet")
        rs.CursorType = adOpenStatic
        set rs = oConn.Execute(sql1)
        if not rs.EOF then
            redim participantes_viejos(0)
            participantes_viejos(0) = cstr(rs("idUsuario"))
            rs.movenext
            while not rs.EOF
                redim preserve participantes_viejos(ubound(participantes_viejos) + 1) 
                participantes_viejos(ubound(participantes_viejos)) = cstr(rs("idUsuario"))
                rs.movenext
            wend
        else
            redim participantes_viejos(-1)
        end if
        
        rs.close
        set rs = nothing

        sql2 ="Update " + GLOBAL_DB_ADM + GLOBAL_TABLA_AGENDA + " SET " & _ 
        "idUsuarioAutor =" & usuario_creador & "," & _ 
        "fecha= '" & cstr(fecha) & "'," & _ 
        "titulo='" & titulo & "'," & _ 
        "actividad='" & actividad & "'," & _ 
        "tipo_actividad='" & tipo_actividad & "'," & _
        "aviso_email='" & avisar_por_mail & "'," & _
        "fecha_termino='" & fecha & "'," & _
        "sala=" & sala & "," & _
        "serial_inicio=" & hora_inicio & "," & _
        "serial_termino=" & hora_termino & "," & _
        "transaccion='" & ToTransacion(fecha, sala, hora_inicio, hora_termino) & "' where correlativo=" & cstr(idReserva)
        'response.write sql2
        oConn.Execute sql2
        
        participantes_nuevos = Filtrar(participantes_actuales, participantes_viejos)
        
        sql3 = ""
        for j = 0 to Ubound(participantes_nuevos)
            sql3 = sql3 & " Insert " + GLOBAL_DB_ADM + GLOBAL_TABLA_USUARIOS_RESERVA & _
            + "(correlativo,idUsuario) values (" + cstr(idReserva) + "," + cstr(participantes_nuevos(j)) + ");" 
        next
        if(sql3 <> "" ) then
            'Response.Write sql3
            oConn.Execute sql3 
        end if

        sql4 = ""
        participantes_eliminados = Filtrar(participantes_viejos, participantes_actuales)
        for k = 0 to uBound(participantes_eliminados)
            sql4 = sql4 & "delete " + GLOBAL_DB_ADM + GLOBAL_TABLA_USUARIOS_RESERVA & _
                + " where correlativo= " + cstr(idReserva) + " and idUsuario=" + cstr(participantes_eliminados(k)) + ";"
        next
        if (sql4 <> "") then
        	'Response.Write sql4
            oConn.Execute sql4
        end if

        oConn.CommitTrans
        If Err.Number <> 0 Then
            oConn.RollBackTrans
        End If

        oConn.close
        set oConn = nothing
	end sub

    function BuscarReservaByFechaAndHoraInicioTermino(id_reserva, sala, fecha, hora_inicio, hora_termino)
        query = "Select top 1 * from " + GLOBAL_DB_ADM + GLOBAL_TABLA_AGENDA + " Where sala=" + sala + " and fecha='" + fecha + "' and (" & _
        "serial_inicio  between " & hora_inicio & " and " & hora_termino & " or " & _
        "serial_termino  between "& hora_inicio & " and " & hora_termino & ") and correlativo <> " + id_reserva
        'Response.write query
        set rs_reserva = Server.CreateObject("ADODB.Recordset")
        rs_reserva.open query , GLOBAL_DSN
        if not rs_reserva.eof then
            Dim reserva : set reserva = Server.createObject("Scripting.Dictionary")
            for each f in rs_reserva.Fields
                reserva.add f.Name , f.Value
            next
            set BuscarReservaByFechaAndHoraInicioTermino = reserva
        else
            set BuscarReservaByFechaAndHoraInicioTermino = nothing
        end if
    end function

    Function eliminar()
        On Error resume Next
        idReserva = Request.Form("IdReserva")
        salida = ""
        Set objReserva = GetReservaById(idReserva)
        if (Not objReserva is Nothing) then
            
            if (PuedeEditarReserva(idReserva, LoggedIn)) then
                Set oConn = Server.CreateObject("ADODB.Connection")
                oConn.Open GLOBAL_DSN
                strSQL = "Delete from " + GLOBAL_DB_ADM + GLOBAL_TABLA_AGENDA + " Where correlativo=" & cstr(idReserva)
                set rs_salida = oConn.execute (strSQL)
                oConn.close
                set oConn = Nothing
                If Err.Number <> 0 then
                    salida = Err.Description
                else
                    salida= "1"
                End If
            else
                salida = "-2"    
            end if
        else
            salida = "-1"
        end if
        eliminar = salida
    End function

    function GetReservaById(id)
        set GetReservaById = Nothing
        if (id <> "") then
            set xfp_rs = CreateObject("ADODB.Recordset")
            xfp_sQry="Select u.* from " + GLOBAL_DB_ADM + GLOBAL_TABLA_AGENDA + " u where u.correlativo = " + cstr(id)
            xfp_rs.Open xfp_sQry, GLOBAL_DSN

            if not xfp_rs.eof and not xfp_rs.bof then
                Dim usuario : Set usuario = Server.createObject("Scripting.Dictionary")
                For each f in xfp_rs.Fields
                    usuario.Add f.Name , f.Value
                Next
                set GetReservaById = usuario
            else
                Set GetReservaById = Nothing
            end if

            xfp_rs.close
            set xfp_rs = Nothing
        end if
    end Function

    function GetParticipantesByIdReserva(id)
        Dim participantes : Set participantes = Server.createObject("Scripting.Dictionary")
        if (id <> "") Then
            set rs_participantes = CreateObject("ADODB.Recordset")
            query ="Select u.* from " + GLOBAL_DB_ADM + GLOBAL_TABLA_USUARIOS_RESERVA + " ru, " + GLOBAL_DB_ADM + GLOBAL_TABLA_USUARIOS + " u where ru.IdUsuario=u.IdUsuario and ru.correlativo = " & cstr(id)
            rs_participantes.Open query, GLOBAL_DSN
            
            if not rs_participantes.eof and not rs_participantes.bof then

                while not rs_participantes.eof
                    Dim participante : Set participante = Server.createObject("Scripting.Dictionary")
                    for each f in rs_participantes.Fields
                        participante.add f.Name , f.Value
                    next
                    participantes.add participante.item("IdUsuario") , participante
                    rs_participantes.movenext
                wend

            end if

            rs_participantes.close
            set rs_participantes = nothing
        end if 
        set GetParticipantesByIdReserva = participantes
    end function

    Function deleteReservaById(idReserva, idSesionUsuario)
        On Error resume Next
        salida = ""
        Set objReserva = GetReservaById(idReserva)
        if (Not objReserva is Nothing) then
            if (cstr(objReserva.item("IdUsuarioAutor")) = cstr(idSesionUsuario)) then 
                Set oConn = Server.CreateObject("ADODB.Connection")
                oConn.Open GLOBAL_DSN
                strSQL = "Delete from " + GLOBAL_DB_ADM + GLOBAL_TABLA_AGENDA + " Where correlativo=" & cstr(idReserva)
                set rs_salida = oConn.execute (strSQL)
                oConn.close
                set oConn = Nothing
                If Err.Number <> 0 then
                    salida = Err.Description
                else
                    salida= "1"
                End If
            else
                deleted = false

                'Preguntamos si la id de usuario tiene abogados asociados
                set objAbogados = GetAbogadosByIdSecretaria(idSesionUsuario)
                if (Not objAbogados is Nothing) then 
                    for each key in objAbogados
                        set objAbogado = objAbogados.item(key)
                        if (not objAbogado is nothing) then 
                            if (cstr(objAbogado.item("IdUsuario")) = cstr(objReserva.item("IdUsuarioAutor"))) then
                                Set oConn = Server.CreateObject("ADODB.Connection")
                                oConn.Open GLOBAL_DSN
                                strSQL = "Delete from " + GLOBAL_DB_ADM + GLOBAL_TABLA_AGENDA + " Where correlativo=" & cstr(idReserva)
                                set rs_salida = oConn.execute (strSQL)
                                oConn.close
                                set oConn = Nothing
                                If Err.Number <> 0 then
                                    salida = Err.Description
                                else
                                    salida= "1"
                                End If
                                deleted = true
                            End if
                        End if
                    Next
                    if (not deleted) then 
                        salida= "-2"
                    end if
                end if
            end if
        else
            salida = "-1"
        end if
        deleteReservaById = salida
   End function
%>