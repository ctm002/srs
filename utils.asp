<%
    function Iff(condicion, verdadero, falso)
        if (condicion) then 
            Iff = verdadero
        else
            Iff = falso
        end if
    end function

    function FormatoHora(hora)
        hora_temp = cstr(hora)
        if (len(hora_temp) = 3) then 
            hora_temp = "0" & hora_temp
        end if 
        dim hh
        hh = Mid(hora_temp,1,2)
        dim mm
        mm = Mid(hora_temp,3,2)
        FormatoHora = hh & ":" & mm
    end function

    function FormatDateString(fecha,tipo)
        sDia = day(fecha)
        sMes = month(fecha)
        sAnio = year(fecha)
        if Len(sDia) = 1 then sDia = "0" & sDia
        if Len(sMes) = 1 then sMes = "0" & sMes
        if (tipo = 1) then
            FormatDateString = sAnio & sMes & sDia
        else 
            FormatDateString = sDia & "-" & sMes & "-" & sAnio
        end if 
    end function	
	
	function ToTransacion(fecha, sala, hora_inicio, hora_termino)
		ToTransacion = fecha + "." + sala + "." + hora_inicio + "." + hora_termino
	end function

    function Filtrar(participantes_nuevos, participantes_viejos)
        dim datos()
        redim datos(-1)
        encontrado = false
        For i = 0 to Ubound(participantes_nuevos)
            For j = 0 to Ubound(participantes_viejos)
                if (cint(participantes_nuevos(i)) = cint(participantes_viejos(j))) then
                    encontrado = true
                end if
            next

            if (encontrado = false) then 
                if(ubound(datos) < 0) then
                    redim datos(0)
                    datos(0) = participantes_nuevos(i)
                else
                    tamanio = ubound(datos) + 1 
                    redim preserve datos(tamanio)
                    datos(tamanio) = participantes_nuevos(i)
                end if
            end if
            encontrado = false
        next
        Filtrar = datos
    end function
%>
