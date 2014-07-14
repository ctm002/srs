<%@ LANGUAGE="VBSCRIPT"%>
<%
    NroPiso = Request.QueryString("piso")
    if NroPiso  = "" then
        Nav = "trabajo.htm?piso=9"
    end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Frameset//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-frameset.dtd">
<html>
<head>
    <title>Reserva de Salas</title>
</head>
<frameset rows="*" framespacing="0" border="1" frameborder="0">
    <frameset cols="225,*">
        <frameset rows="210,*">
            <frame name="principal" src="calendario.asp" target="_self" frameborder="1">
            <frame name="menu" src="menu.asp" frameborder="1">
        </frameset>
        <frame name="trabajo" src="<%=Nav%>" scrolling="auto" target="_self"></frame>
    </frameset>
    <noframes>
        <p>Esta página usa marcos, pero su explorador no los admite.</p>
    </noframes>
</frameset>
</html>
