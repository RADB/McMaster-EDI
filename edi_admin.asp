<!-- #include virtual="/shared/admin_security.asp" -->
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
				<font class="boldtextblack">Home</font>
			</td>
			<td align="right">
				<!--<a class="reglinkMaroon" href="default.asp?status=logout">Logout</a>-->
				<input type="button" onclick="javascript:window.location='default.asp?status=logout';" name="Logout" value="Logout">
			</td>
		</tr>
	</table>
	<%
	call open_adodb(conn,"MACEDI")

	set rstTables = server.CreateObject ("adodb.recordset")
	' get all the page headings
	strQuery = "SELECT [english], [link], [issectionheader], [isheader], [section], [intSecurity] FROM page_admin ORDER BY [section],[intOrder], [isHeader], [english]"
	rstTables.Open strQuery, conn
				
	if not (rstTables.bof and rstTables.eof) then 
		do while not rsttables.EOF
			' section header
			if rstTables("issectionheader") then
				' outer table/border
				Response.Write "<table width=""760"" cellpadding=""2"" cellspacing=""0"" border=""1"" bordercolor=""006600"">"
				Response.Write "<tr><td>"
				' inner table
				Response.Write "<table width=""750"" cellpadding=""0"" cellspacing=""0"" border=""0"">"
				Response.Write "<tr><td colspan=""2"" align=""center""><font class=""headerBlue"">" & rstTables("english") & "</font></td></tr>"
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
				Response.Write "<tr><td colspan=""2"" align=""left""><font class=""subheaderBlue"">" & rstTables("english") & "</font></td></tr>"				
			else
				'check security clearance
				if session("access") >= rstTables("intSecurity") then 
					' write the links 
					Response.Write "<tr><td width=""25"">&nbsp;</td><td width=""325"" align=""left""><a href=""" & rstTables("link") & """ class=""regLinkMaroon"">" & rstTables("english") & "</a></td></tr>"
				end if 
			end if 
			'Response.Write rstTables("english") & " " & rstTables("issectionheader")
			rstTables.MoveNext 
		loop
		
		' close the last section table
		Response.Write "</table></td></tr>"
		
		' email webmaster  
		Response.Write "<tr><td align=""center"" colspan = ""2""><br />"
		Response.Write "<a href=""mailto:webmaster@e-edi.ca"" class=""reglinkblue"">Questions or Comments: webmaster@e-edi.ca</a>"
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