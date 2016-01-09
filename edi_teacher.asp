<!-- #include virtual="/shared/security.asp" -->
<%
' if the user has not logged in they will not be able to see the page
if blnSecurity then
%>

<html>
<!-- #include virtual="/shared/head.asp" -->	
<body>
	<!-- #include virtual="/shared/page_header.inc" -->
	<br>
	<table width="760" border="0">
		<tr>
			<td>
				<%
				'if session("language") = "English" then 
					Response.Write "<font class=""boldtextblack"">"& strHome &"</font>"
					
				'else
				'	Response.Write "<font class=""boldtextblack"">accueil</font>"
				    
				'end if 
				%>
			</td>
			<td align="right">
				<!--<a class="reglinkMaroon" href="default.asp?status=logout">Logout</a>-->
				<%
				'if session("language") = "English" then 
					Response.Write "<input type=""button"" onclick=""javascript:window.location='default.asp?status=logout';"" name=""" & lblLogout & """ value=""" & lblLogout & """>"
				'else
				'	Response.Write "<input type=""button"" onclick=""javascript:window.location='default.asp?status=logout';"" name=""" & lblLogout & """ value=""" & lblLogout & """>"
				'end if 
				
				%>	
			</td>
		</tr>
	</table>
	<%
	'call open_adodb(conn,"TABLES")
    call open_adodb(conn, "MACEDI")
	set rstTables = server.CreateObject ("adodb.recordset")
	' get all the page headings
	strQuery = "SELECT [" & session("Language") & "] as menu_text, [link], issectionheader, isheader, [section] FROM page_teacher ORDER BY [section],intOrder, isHeader, english"
	rstTables.Open strQuery, conn
				
	if not (rstTables.bof and rstTables.eof) then 
		do while not rsttables.EOF
		    ' used to target the links - new window vs self
		    strTarget = "_self"
			' section header
			if rstTables("issectionheader") then
				' outer table/border
				Response.Write "<table width=""760"" cellpadding=""2"" cellspacing=""0"" border=""1"" bordercolor=""006600"">"
				Response.Write "<tr><td>"
				' inner table
				Response.Write "<table width=""750"" cellpadding=""0"" cellspacing=""0"" border=""0"">"
				Response.Write "<tr><td colspan=""2"" align=""center""><font class=""headerBlue"">" & rstTables("menu_text") & "</font></td></tr>"
				' spacer line
				Response.Write "<tr><td colspan=""2""><br /></td></tr>"
			' odd header
			elseif rstTables("isheader") then
				' odd section header
				if rstTables("section") mod 2 = 1 then 
					' if not the first section then close the last section
					if rstTables("section") > 1 then 
						Response.Write "</table></td></tr>"
					end if 
					' section table on left side
					Response.Write "<tr><td valign=""top"" align=""right""><table width=""350"" cellpadding=""0"" cellspacing=""0"" border=""0"">"	
				' even section header
				elseif rstTables("section") mod 2 = 0 then 
					' spacer line
					Response.Write "<tr><td colspan=""2"" align=""left""><br /></td></tr>"
					' section table on right side
					Response.Write "</table></td><td valign=""top""><table width=""350"" cellpadding=""0"" cellspacing=""0"" border=""0"">"	
				end if 
				
				' write the header
				Response.Write "<tr><td colspan=""2"" align=""left""><font class=""subheaderBlue"">" & rstTables("menu_text") & "</font></td></tr>"
			else
				strLink = rstTables("link")
				
				if strLink = "documents\EDI%20Guide.pdf" Then 
					select case session("province")
					    ' Alberta
					    case 3 
					        strLink = "documents\2015EDIGuideAlberta" & session("Language") & ".pdf"
					    ' Manitoba
					    case 2 
					        strLink = "documents\2015EDIGuideManitoba" & session("Language") & ".pdf"
					    ' Saskatchewan
					    Case 5 
					        strLink = "documents\2015EDIGuideSaskatchewan" & session("Language") & ".pdf"
					    Case 6 					       
					        strLink = "documents\2015EDIGuideNorthwestTerritories" & session("Language") & ".pdf"
					    Case 7 					      
					        strLink = "documents\2015EDIGuideNewfoundlandLabrador" & session("Language") & ".pdf"
					    Case 8 					        
					        strLink = "documents\2015EDIGuideNovaScotia" & session("Language") & ".pdf"
					    ' Ontario
					    Case 1 					    
					        strLink = "documents\2015EDIGuideOntario" & session("Language") & ".pdf"
					    ' all others use ontario
					    case else
					        strLink = "documents\2015EDIGuideOntario" & session("Language") & ".pdf"
					end select 
					
					'if session("language") = "French" then 
					'    strLink = replace(strLink,".pdf","French.pdf")
					'end if 
					session("strLink") = strLink
					strTarget = "_blank"
				end if
					
				
				' write the links 
				Response.Write "<tr><td width=""25"">&nbsp;</td><td width=""325"" align=""left""><a target=""" & strTarget & """ href=""" & strLink & """ class=""regLinkMaroon"">" & rstTables("menu_text") & "</a></td></tr>"
			end if 
			'Response.Write rstTables("menu_text") & " " & rstTables("issectionheader")
			rstTables.MoveNext 
		loop
		
		' close the last section table
		Response.Write "</table></td></tr>"
		
		' email webmaster  
		Response.Write "<tr><td align=""center"" colspan = ""2""><br />"
		Response.Write "<a href=""mailto:webmaster@e-edi.ca"" class=""reglinkblue"">" & strQuestions & ": webmaster@e-edi.ca</a>"
		Response.Write "</td></tr>"
		
		' close the inner table
		Response.Write "</table>"
		
		' close the bordered table
		Response.Write "</td></tr></table>"
	end if 
	
	' kill and close all connections and recordsets
	call close_adodb(rstTables)
	call close_adodb(conn)
	%>
	<!-- #include virtual="/shared/page_footer.inc" -->
</body>
</html>
<%
end if
%>