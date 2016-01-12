<table  width="760" border="0" cellpadding="0" cellspacing="0">			
	<tr>
		<td rowspan = "2" valign="middle" align="right" width="100">
			<%		
			if session("language") = "French" then 
				'response.write "<img src=""images/x e-edi_french.gif"" alt=""e-EDI.ca"" name=""e-edi_french.gif"" />"
				response.write "<img src=""images/e-edi.jpg"" alt=""e-EDI.ca"" name=""e-edi.jpg"" height=""40px"" />"
			else
				response.write "<img src=""images/e-edi.jpg"" alt=""e-EDI.ca"" name=""e-edi.jpg"" height=""40px"" />"
        end if
        %>
        &nbsp;
		</td>
		<td align="left" width="605" nowrap="nowrap">
		<%		
			if session("language") = "French" then 
				response.write "<table width=""605"" border=""0"">"
				response.write "<tr>"
				response.write "<td align=""center"">"
				response.write "<font class=""frenchHeaderBlue"">INSTRUMENT DE MESURE DU DÉVELOPPEMENT DE LA PETITE ENFANCE</font>"				
			else
				response.write "<table width=""385"" border=""0"">"
				response.write "<tr>"
				response.write "<td align=""center"">"
				response.write "<font class=""headerBlue"">EARLY DEVELOPMENT INSTRUMENT</font>"
			end if
			response.write "</td>"
			response.write "</tr>"
		    response.write "</table>"
		%>			
		</td>
	</tr>
	<tr>
		<td align="left" nowrap="nowrap">
		<%
			if session("language") = "French" then 
				response.write "<table width=""605"" border=""0"">"
				response.write "<tr>"
				response.write "<td align=""center"">"
				response.write "<font class=""frenchSubHeaderBlue"">Une mesure axée sur la population à l’intention des collectivités</font>"				
			else
				response.write "<table width=""385"" border=""0"">"
				response.write "<tr>"
				response.write "<td align=""center"">"
				response.write "<font class=""subHeaderBlue"">A Population-Based Measure for Communities</font>"
			end if
			response.write "</td>"
			response.write "</tr>"
			response.write "</table>"
		%>
		</td>
		<td>			
			<%		
			'language selector
			if Request.ServerVariables("URL") = "/edi_teacher.asp" or Request.ServerVariables("URL") = "/edi_teacher_class.asp"then 
				response.write "<input name=""language"" type=""checkbox"" "
																				
				' check to see language
				if session("language") = "French" then 				
					response.write " onchange=""javascript:document.location=location.pathname +'?Language=English';"""
				elseif session("language") = "English" then 				
					response.write "checked onchange=""javascript:document.location=location.pathname +'?Language=French'"""
				end if 
				
				response.write "data-toggle=""toggle"" data-off=""French"" data-on=""English"" data-onstyle=""info"" data-offstyle=""warning"" data-size=""mini"">"
			end if 
			%>			
		</td>
	</tr>
</table>
<br />
<%= strTopBar %>