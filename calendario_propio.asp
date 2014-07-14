<!--#include file="consultas.asp"-->
<!--#include file="utils.asp"-->
<%
set fp_rs = CreateObject("ADODB.Recordset")
idSesionUsuario = Session("sesion_usuario_id")
tipoSesionUsuario = Session("sesion_usuario_tipo")

Main
function ingresa_cero(xnumero)
	numero=trim(cstr(xnumero))
	largo=len(numero)
	salida=string(2-largo,"0") + numero
	ingresa_cero=salida
end function

Sub Main
matriz=array("","Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre")
	Global_Mes=Request.Form("XMES")
	Global_Year=Request.Form("XYEAR")

	if Global_Mes="" then Global_Mes=Request.Querystring("XMES")
	if Global_Year="" then Global_Year=Request.Querystring("XYEAR")

	if Global_Mes="" then Global_Mes=cstr(month(date))
	if Global_Year="" then Global_Year=cstr(year(date))

	xfecha=Global_Mes+"/"+"01/"+Global_Year
	xfirst=weekday(Cdate(xfecha))-1
	if xfirst=0 then
		xfirst=7
	end if
%>
<html>
<head>
<style type="text/css">
    .cell_over { BACKGROUND-COLOR: #FFFF99 }
    .cell_out { BACKGROUND-COLOR: #FFFFFF }
    .cell_over1 { BACKGROUND-COLOR: #C9E4FC }
    .cell_out1 { BACKGROUND-COLOR: #FFFFCC }
</style> 
<script type="text/javascript">
    function cambiar_fecha(xyear,xmes,xdia)
    {
        Direccion="ver_dia.asp?xyear="+ xyear+"&xmes=" + xmes +"&xdia=" + xdia ;
        parent.trabajo.location.href=Direccion;
    }
</script>
<meta http-equiv="Content-Type" content="text/html; charset=Utf-8"/>
<title>Mis Reuniones</title>
</head>

<body bgcolor="#CCCCCC" topmargin="0" leftmargin="8" rightmargin="0" bottommargin="0">
<form method="post" action="calendario_propio.asp">
<table border="1" align="center" width="500" id="table4">
	<tr>
		<td width="100%" bgcolor="#EEF1E8">
		
			<!--webbot bot="SaveResults" U-File="fpweb:///_private/form_results.csv" S-Format="TEXT/CSV" S-Label-Fields="TRUE" -->
			<table border="0" width="100%" id="table5">
				<tr>
					<td bgcolor="#45496F" width="100%" valign="middle">
					<p align="center">
					<font face="Verdana" size="2" color="#FFFFFF"><b>Mes & Año 
					</b></font></td>
				</tr>
				<tr>
					<td width="100%" valign="top">
					<p align="center">
						<select OnChange="document.forms[0].submit();" size="1" name="xmes" style="font-family: Verdana; font-size: 8pt">
						<% for i = 1 to 12%>
						<option value="<% response.write i %>" <%if cdbl(i)=cdbl(Global_Mes) then response.write " selected "%>><% response.write matriz(i) %></option>
						<%next %>
						</select>
						<select OnChange="document.forms[0].submit();" size="1" name="xyear" style="font-family: Verdana; font-size: 8pt">
						<% for i = 2000 to 2014%>
						<option value="<%response.write i%>" <%if cdbl(i)=cdbl(Global_Year) then response.write " selected "%>><%response.write i%></option>
						<% next%>
						</select>
					</td>
				</tr>
				<% if(tipoSesionUsuario = "SECRETARIA")then %>
			        <tr>
			            <td width="100%" valign="top">
			            	<p align="center">
			                <select id="cboAbogados" name="cboAbogados" style="width:193px" OnChange="document.forms[0].submit();" >
			                    <%
			                        set abogados = GetAbogadosByIdSecretaria(Cint(idSesionUsuario))
			                        for each key in abogados
			                            set abogado = abogados.item(key)
			                            if (cstr(Request.form("cboAbogados")) = cstr(abogado.item("IdUsuario"))) then %>
			                              	<option selected="selected" value="<%=abogado.item("IdUsuario")%>"><%=abogado.item("Nombres") + " " + abogado.item("Apellidos") %></option>
			                			<% else %>
			                				<option value="<%=abogado.item("IdUsuario")%>"><%=abogado.item("Nombres") + " " + abogado.item("Apellidos") %></option>
			                			<% end if %>
			                	<%  next %>
			                </select>
			            	</p>
			            </td>
			        </tr>
			        <tr>
			    <% end if %>
			</table>
		</td>
	</tr>
</table>
</form>
<p align="center"><font face="Verdana" size="5" color="#000000">Mis Reuniones <% response.write matriz(Global_Mes)%></font></p>
<table border="1" width="550" align="center" id="table1" cellspacing="0" bordercolor="#45496F">
	<tr>
		<td>
		<table border="1" width="800px" id="table2" cellspacing="0">
			<tr>
				<td height="100%" align="center" bgcolor="#45496F" ><b>
				<font face="Verdana" size="1" color="#FFFFFF">Lu</font></b></td>
				<td height="100%" align="center" bgcolor="#45496F" ><b>
				<font face="Verdana" size="1"  color="#FFFFFF">Ma</font></b></td>
				<td height="100%" align="center" bgcolor="#45496F" ><b>
				<font face="Verdana"  size="1" color="#FFFFFF">Mi</font></b></td>
				<td height="100%" align="center" bgcolor="#45496F" ><b>
				<font face="Verdana"  size="1" color="#FFFFFF">Ju</font></b></td>
				<td height="100%" align="center" bgcolor="#45496F" ><b>
				<font face="Verdana"  size="1" color="#FFFFFF">Vi</font></b></td>
				<td height="100%" align="center" bgcolor="#45496F" ><b>
				<font face="Verdana"  size="1" color="#FFFFFF">Sa</font></b></td>
				<td height="100%" align="center" bgcolor="#45496F" ><b>
				<font face="Verdana"  size="1" color="#FFFFFF">Do</font></b></td>
			</tr>
		<%cuenta_global=1
		  cuenta_dia=1
			for i = 1 to 6%>
			<tr>

				<% for x = 1 to 7
					xfecha=cstr(cuenta_dia) + "/"+Global_Mes+"/"+Global_Year
					select case x
						case 6
							back="#FFFFCC"
						case 7
							back="#FFFFCC"
						case else
							back="#ffffff"
					end select
				'	back=Buscar_Stand(xfecha)
                %>

				<td onmouseover="this.className='cell_over';" onmouseout="this.className='cell_out';" bgcolor="<% response.write back%>">
				<div align="center">
				<table border="0"  cellSpacing="0" width="90%" id="table3" height="100%">
					<tr>

						<% if cuenta_global>=xfirst and isdate(xfecha) then %>
						    <td align="left" width="79" background="./img/numeros/<% response.write cuenta_dia%>.jpg" height="79" valign="top" onmouseover="this.className='cell_over';"  onmouseout="this.className='cell_out';">
							    
                                <%referencia="href=""javascript:cambiar_fecha("+ global_year+","+ global_mes+","+ cstr(cuenta_dia)+")"">"%>							
                                <% 
                                	usuario_autor = ""
                                	if (tipoSesionUsuario = "SECRETARIA") then
                                		usuario_autor = cstr(Request.Form("cboAbogados"))
                                	else
                                		usuario_autor = idSesionUsuario
                                	end if 
                                	reserva_text = ""
                                	if busca_reunion(cuenta_dia, global_mes, global_year, usuario_autor, reserva_text) then %>
                                    	<b><font color="0000ff"><%response.write reserva_text %></font></b>
                                <%end if%>
						    </td>
                            
                            <% cuenta_dia=cuenta_dia+1%>
						
                        <% else %>						    
                            <td align="left" width="79" height="79" valign="top" onmouseover="this.className='cell_over';"  onmouseout="this.className='cell_out';"></td>
						
                        <% end if
						   cuenta_global= cuenta_global + 1
						%>
					</tr>
				</table>
				</div>
				</td>
				<%next %>
			</tr>
		    <%next%>
        </table>
		</td>
	</tr>
</table>
</body>
</html>
<%end sub%>