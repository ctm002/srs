<%@ Language=VBScript%>
<!--#include file="services_reserva.asp"-->
<!--#include file="utils.asp"-->
<!--#include file="consultas.asp"-->
<% 
	url = split(Request.ServerVariables("HTTP_REFERER"), "?")
	paginaAnterior = cstr(url(0)) & "?" & cstr(Request.ServerVariables("QUERY_STRING"))
	idUsuarioCreador = Request.QueryString("id_usuario")
	nombresSesionUsuario = Session("sesion_usuario_nombres")
	idSesionUsuario = Session("sesion_usuario_id")
	tipoSesionUsuario = Session("sesion_usuario_tipo")
	idSala = Request.QueryString("sala")
	idPiso = Request.QueryString("piso")
%>
<html>
	<head>
		<title>Editar Reserva</title>
		<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
    	<link rel="stylesheet" href="./jquery/css/jquery.ui.all.css" /> 
		<script src="./jquery/jquery-1.10.2.js" type="text/javascript"></script>
    	<script src="./jquery/ui/jquery.ui.core.js" type="text/javascript"></script>
		<script src="./jquery/ui/jquery.ui.widget.js" type="text/javascript"></script>
		<script src="./jquery/ui/jquery.ui.datepicker.js" type="text/javascript"></script>
    	<script src="./jquery/jquery.inputmask.js" type="text/javascript"></script>
    	<script src="./jquery/jquery.timepicker.js" type="text/javascript"></script>
    	<link rel="stylesheet" type="text/css" href="./css/jquery.timepicker.css" />
    	<link rel="stylesheet" href="./css/demos.css" />
    	<script src="./jquery/jquery.validate.js" type="text/javascript"></script>

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

	        $(function () {
	            $("#txtFecha").datepicker();
	            $("#txtFecha").inputmask("99-99-9999")
	            $("#txtHoraDesde").timepicker({ 'timeFormat': 'H:i', 'minTime': '7:00am', 'maxTime': '11:00pm', 'step': 60 });
	            $("#txtHoraHasta").timepicker({ 'timeFormat': 'H:i', 'minTime': '7:00am', 'maxTime': '11:00pm', 'step': 60 });
	            $('#txtHoraHasta').on('changeTime', function () {
	                var hora_hasta = $(this).val();
	                var hora_desde = $('#txtHoraDesde').val();
	                if (hora_desde >= hora_hasta) {
	                    $(this).val(hora_desde);
	                }
	            });

	            $("#frmEditar").validate({
	                rules: {
	                    txtFecha: {
	                        required: true
	                    },
	                    txtHoraDesde: {
	                        required: true
	                    },
	                    txtHoraHasta: {
	                        required: true
	                    },
	                    cboSalas:{
	                        required: true
	                    },
	                    txtSub: {
	                        required: true,
	                        maxlength: 100
	                    },
	                    txtAct: {
	                        required: true,
	                        maxlength: 250
	                    }
	                },
	                messages: {
	                    txtFecha: { required: "Ingrese fecha de la reserva" },
	                    txtHoraDesde:{required: "Ingrese hora de inicio de la reserva" },
	                    txtHoraHasta: { required: "Ingrese hora de termino de la reserva" },
	                    cboSalas:{required:"Seleccione sala para la reserva"},
	                    txtSub: {
	                        required: "* Por favor ingrese motivo de la reunion",
	                        maxlength: "* El motivo de la reunion no debe superar los 50 caracteres"
	                    },
	                    txtAct: {
	                        required: "* Por favor ingrese la descripcion de la reunion",
	                        maxlength: "* La descripcion de la reunion no puede superar los 250 caracteres"
	                    }
	                }
	            });

				$("#frmEditar").submit(function (event){
                    var idReserva = document.forms[0].elements["IdReserva"].value; 
                    var idSala = document.forms[0].elements["cboSalas"].value;
                    var fecha = document.forms[0].elements["txtFecha"].value;
                    var horaInicio = document.forms[0].elements["txtHoraDesde"].value;
                    var horaTermino = document.forms[0].elements["txtHoraHasta"].value;

                    var params = "&id_reserva=" + idReserva + "&sala=" + idSala + "&fecha=" + fecha + "&desde=" + horaInicio + "&hasta=" + horaTermino;                
                    $.ajax({
                        type: "POST",
                        url: "services_reserva.asp",
                        contentType: "application/x-www-form-urlencoded; charset=UTF-8",
                        data: "op=buscar" + params ,
                        dataType: "json",
                        async: false,
                        error: function (xhr, status, error) {
                            alert(error);
                            event.preventDefault();
                        }
                    }).done(function (data) {
                        respuesta = data.message;
                        if (respuesta == "0") 
                        {                           	                          	
						    $.ajax({
			                    type: "POST",
			                    url: "services_reserva.asp",
			                    contentType: "application/x-www-form-urlencoded; charset=UTF-8",
			                    data: "op=puedeEditar&id_reserva=" + idReserva ,
			                    dataType: "json",
			                    async: false,
			                    error: function (xhr, status, error) {
			                        alert(error);
			                        event.preventDefault();
			                    }
			                }).done(function (data) {
			                    respuesta = data.message;
			                    if (respuesta == "false")
			                    {
			                    	alert("Solo el usuario creador de la reserva puede editarla");
                        			event.preventDefault();
			                    }
			                });							
                        }
                        else 
                        {
                        	alert("Ya existe una reserva con ese horario");
                            event.preventDefault();
                        }
                    }); 
                });
	        });
			
			function appendParticipante() {
	            var e = document.getElementById("cboUsuarios");
	            var strIdUser = e.options[e.selectedIndex].value;
	            var strNombreUser = e.options[e.selectedIndex].text;

	            var opt = document.createElement('option');
	            opt.value = strIdUser;
	            opt.innerHTML = strNombreUser;
	            opt.selected = true;
	            var cboParticpantes = document.getElementById("cboParticipantes")
	            cboParticpantes.appendChild(opt);
	        }

	        function removeParticipante() {
	            var cboParticipantes = document.getElementById("cboParticipantes");
	            var opt = cboParticipantes.options[cboParticipantes.selectedIndex]
	            cboParticipantes.removeChild(opt);
	        }

	        function onChange() {
	            var e = document.getElementById("cboPisos");
	            var id = e.options[e.selectedIndex].value;
	            var text = e.options[e.selectedIndex].text;
	            $.ajax({
	                type: "POST",
	                url: "consultas.asp",
	                contentType: "application/x-www-form-urlencoded; charset=UTF-8",
	                data: "Action=GetSalasByIdPiso&Id=" + id,
	                dataType: "json",
	                async: false,
	                success: function (data) {
	                    var $select = $('#cboSalas');
	                    $select.find('option').remove();
	                    $.each(data.d, function (key, value) {
	                        $select.append('<option value=' + data.d[key].id + '>' + 'Sala ' + value.nro + '</option>');
	                    });
	                },
	                error: function (xhr, status, error) {
	                    console.log(arguments);
	                    var err = eval("(" + xhr.responseText + ")");
	                    alert(err.Message);
	                }
	            });
	        }

	        function onBtnVolver_Click(urlAnterior) {
            	document.location.href = urlAnterior;
        	}

        	function onBtnEliminar_Click(idSesionUsuario) {
	            var url = ""; 
	            var response = confirm("¿Esta seguro que desea eliminar la reserva de sala?")
	            if (response == true) {
	                var idReserva = document.forms[0].elements["IdReserva"].value;
	                if (idReserva != "") {
	                    $.ajax({
	                        type: "POST",
	                        url: "services_reserva.asp",
	                        contentType: "application/x-www-form-urlencoded; charset=UTF-8",
	                        data: "op=eliminar&IdReserva=" + idReserva + "&IdSesionUsuario=" + idSesionUsuario,
	                        dataType: "json",
	                        async: false,
	                        success: function (data) { },
	                        error: function (xhr, status, error) {
	                            alert(error)
	                        }
	                    }).done(function (data) {
	                        if (data.message == 1) {
	                            var idPiso = document.forms[0].elements["IdPiso"].value;
	                            var fecha = document.forms[0].elements["txtFecha"].value;
	                            var params = "?piso=" + idPiso + "&fecha=" + fecha;
	                            url = "ver_dia.asp" + params;
	                            document.location.href = url;
	                        } else if (data.message == -1) {
	                            alert("La reserva no existe en la base de datos");
	                        } else if (data.message == -2) {
	                            alert("Solo el usuario autor de la reserva puede eliminarla");
	                        } else {
	                            alert(data.message);
	                        }
	                    });
	                }
	            }
            }

            function checkReserva()
            {
                $("#frmEditar" ).submit();  
            }            
        </script>
	</head>
	<body>
	<%
		idReserva = Request.QueryString("id")
       	set reserva = GetReservaById(idReserva)
        subject = reserva.item("titulo")
        actividad = reserva.item("actividad")
        fecha = FormatDateTime(reserva.item("fecha"),2)
        horaInicio = reserva.item("serial_inicio")
        horaInicio = Iff(len(horaInicio) = 3,"0" & horaInicio, horaInicio)
        horaTermino = reserva.item("serial_termino")
        horaTermino = Iff(len(horaTermino) = 3,"0" & horaTermino, horaTermino)
        tipo_actividad = reserva.item("tipo_actividad")
        aviso_email = reserva.item("aviso_email")
	%>

	<form id="frmEditar" method="post" action="services_reserva.asp" accept-charset="UTF-8">
    <table border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td>
                <label for="lblUsuarioCreador">RESERVA DE SALA:<b><%=nombresSesionUsuario%></b></label>
            </td>
        </tr>
        <tr>
            <td>
            	<input type="hidden" name="op" value="editar" />
           		<input type="hidden" name="IdReserva" value="<%=idReserva%>" />
				<input type="hidden" name="IdPiso" value="<%=idPiso%>" />
                <input type="hidden" name="IdSala" value="<%=idSala%>" />
            </td>
        </tr>
        <% if(tipoSesionUsuario = "SECRETARIA")then %>
   
            <% if (not reserva is nothing) then %>       
                
                <tr>
                <td>
                    <label for="lblAbogado">Abogado:</label>
                </td>
                </tr>
                <tr>
                <td>
                    <select id="cboAbogado2" name="cboAbogados" style="width:193px">
                    <%
                        set abogados = GetAbogadosByIdSecretaria(Cint(idSesionUsuario))
                        for each key in abogados
                            set abogado = abogados.item(key)
                            if(Not reserva is Nothing) then
                                if (cstr(abogado.item("IdUsuario"))=cstr(reserva.item("IdUsuarioAutor"))) then %>
                                    <option selected="selected" value="<%=abogado.item("IdUsuario")%>"><%=abogado.item("Nombres") + " " + abogado.item("Apellidos")%></option>
                            
                                <%else%>
                                     <option value="<%=abogado.item("IdUsuario")%>"><%=abogado.item("Nombres") + " " + abogado.item("Apellidos") %></option>
                                <%end if
                            else 
                                if ( cstr(Request.Form("cboAbogados")) = cstr(abogado.item("IdUsuario")) ) then %>
                                    <option selected="selected" value="<%=abogado.item("IdUsuario")%>"><%=abogado.item("Nombres") + " " + abogado.item("Apellidos") %></option>
                                <% else %>
                                     <option value="<%=abogado.item("IdUsuario")%>"><%=abogado.item("Nombres") + " " + abogado.item("Apellidos") %></option>
                                <%end if%>
                            <%end if%>
                    <% next %>
                    </select>
                </td>
                </tr>                

            <%end if%> 
       	<%end if%>
        <tr>
            <td>
                <label for="lblHorario">Horario de Reserva:</label>
            </td>
        </tr>
        <tr>
            <td>
                <p style="font-size:15px">
                <input name="txtFecha" id="txtFecha" type="text" value="<%=fecha%>" />
                <input name="txtHoraDesde" id="txtHoraDesde" type="text" value="<%=horaInicio%>" class="time" />
                <label for="lblDesdeHasta">a</label>
                <input name="txtHoraHasta" id="txtHoraHasta" type="text" value="<%=horaTermino%>" class="time"/>
                </p>
            </td>
        </tr>
        <tr>
            <td>
                <label for="lblPisos">Seleccione Piso:</label>
            </td>
        </tr>
        <tr>
            <td>
                <select name="cboPisos" id="cboPisos" style="width:193px" onchange="onChange();">
                    <%  
                        set pisos = GetPisos()
                        for each key in pisos
                            set piso = pisos.item(key)
                            pk_id_piso = piso.item("pk_id_per_piso")
                            nombrePiso = "Piso " & piso.item("descripcion")
                            if (cstr(pk_id_piso) = cstr(idPiso)) then%>
                                <option value="<%=pk_id_piso%>" selected="selected"><%=nombrePiso%></option>
                            <%else%>
                                <option value="<%=pk_id_piso%>"><%=nombrePiso%></option>
                            <%  end if 
                        next 
                    %>
                </select>
            </td>
        </tr>
        <tr>
            <td>
                <label for="lblSalas">Seleccione Sala:</label>
            </td>
        </tr>
        <tr>
            <td>
                <select name="cboSalas" id="cboSalas" style="width:193px">
                    <%
                        'if (idSala <> "")Then 
                            set lstSalas = getSalasByIdPiso(idPiso)
                            for each key in lstSalas
                                set sala = lstSalas.item(key)
                                nro_sala = "Sala " + cstr(sala.item("nro"))
                                id_sala = sala.item("id")
                                if(cstr(id_sala) = cstr(idSala)) then%>
                                     <option value="<%=id_sala%>" selected="selected"><%=nro_sala%></option>
                                <%else%>
                                     <option value="<%=id_sala%>"><%=nro_sala%></option>
                                <% end if 
                            next
                       'end if
                     %>
                </select>
            </td>
        </tr>
        <tr>
            <td>
                <label for="lblUsuarios">Participantes:</label>
            </td>
        </tr>
        <tr>
            <td>
                <select id="cboUsuarios">
                    <% 
                set usuarios = GetUsuarios()
                for each key in usuarios
                    set objUsuario = usuarios.item(key)
                    id_usuario = objUsuario.item("IdUsuario")
                    nombre_usuario = objUsuario.item("Nombres") & " " & objUsuario.item("Apellidos")
                    %>
                    <option value="<%=id_usuario%>"><%=nombre_usuario%></option>
                    <% next %>
                </select>
                <input id="BtnAdd" type="button" value="+" onclick="appendParticipante();" style="width:25px" />
                 <input id="BtnDelete" type="button" value="-" onclick="removeParticipante();"  style="width:25px" />
            </td>
        </tr>
        <tr>
            <td>
                <select name="cboParticipantes" id="cboParticipantes" style="width: 100%" multiple="multiple">
                    <% 
                    set participantes = GetParticipantesByIdReserva(IdReserva)
                    for each key in participantes
                        set participante = participantes.item(key)
                        nombres_participante = participante.item("Nombres") + " " + participante.item("Apellidos") 
                        id_participante = participante.item("IdUsuario")
                    %>
                    <option selected="selected" value="<%=id_participante%>"><%=nombres_participante%></option>
                    <%  next  %>
                </select>
            </td>
        </tr>
        <tr>
            <td>
               
                <table border="0" cellpadding="0" cellspacing="0" width="100%">
                    <tr>
                        <td>
                            <label for="lblActividad">Actividad</label>
                        </td>
                        <td>
                            <label for="lblAvisoEmail">Aviso Email</label>
                        </td>
                    </tr>
                    <tr>
                        <td>
                             <!--Tipo de actividad-->
                             <%
                               if (tipo_actividad = "PERSONAL" or tipo_actividad = "") then
                             %>
                                 <input type="radio" name="tipo_act" value="PERSONAL" checked="checked"/>Personal
                                 <input type="radio" name="tipo_act" value="OFICINA" />Oficina
                             <%else%>
                                <input type="radio" name="tipo_act" value="PERSONAL"/>Personal
                                <input type="radio" name="tipo_act" value="OFICINA" checked="checked" />Oficina
                             <%end if%>
                        </td>
                        <td>
                            <!--Aviso Email-->
                            <%
                              if (aviso_email = "SI" or aviso_email ="") then
                             %>
                                <input type="radio" name="aviso_email" value="SI" checked="checked"/>Si
                                <input type="radio" name="aviso_email" value="NO" />No
                            <%else%>
                                <input type="radio" name="aviso_email" value="SI" />Si
                                <input type="radio" name="aviso_email" value="NO" checked="checked"/>No
                            <%end if%>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td>
                <label for="lblSubject">Subject:</label>
            </td>
        </tr>
        <tr>
            <td>
                <input type="text" id="txtSub" name="txtSub" size="10" style="width: 99%" value="<%=cstr(subject)%>" />
            </td>
        </tr>
        <tr>
            <td>
                <label for="lblActividad">Actividad:</label>
            </td>
        </tr>
        <tr>
            <td>
                <textarea id="txtAct" name="txtAct" cols="40" rows="5" style="font-family: Verdana; font-size: 8pt; width:99%"><%=cstr(actividad)%></textarea>
            </td>
        </tr>
        <tr>
            <td>
                <input type="button" id="BtnVolver" value="Volver" onclick="onBtnVolver_Click('<%=paginaAnterior%>');" style="height: 27px;border:1px solid rgba(0, 0, 0, 0.1);" />
                <input type="button" id="BtnAceptar" value="Editar" style="background:#4285F4; color:White; height: 27px;border:1px solid rgba(0, 0, 0, 0.1);" onclick="checkReserva();"/>
				<input type="button" id="BtnEliminar" value="Eliminar" onclick="onBtnEliminar_Click('<%=idSesionUsuario%>');" style="height: 27px;border:1px solid rgba(0, 0, 0, 0.1);"/>
            </td>
        </tr>
    </table>
    </form>	
	</body>
</html>