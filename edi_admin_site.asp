<!-- #include virtual="/shared/admin_security.asp" -->
<%
' public variables
dim intSite, strName, strCoordinator, strAddress, strCity, intProvince, strPostal,	strPhone, strFax, strEmail, strQ6, strQ7, strQ8, strQ9,	strQ10, strQ11, strComments
dim aData 
on error resume next

' if the user has not logged in they will not be able to see the page
if blnSecurity then
	'call open_adodb(conn, "DATA")
    call open_adodb(conn, "MACEDI")
	set rstData = server.CreateObject("adodb.recordset")

	' delete record
	if Request.Form("Action") = "Delete" then 
		strSql = "DELETE FROM sites WHERE intSiteID = " & Request.Form("code")
		conn.execute strSql 	
		
		if conn.errors.count > 0 then 
			strError = "<font class=""regtextred"">Error Number : " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/></font>"
		end if 
		
		' get the first record in the recordset
		intCode = 0 
	' Add a site - loads an empty form 
	'		     - set all values = ""
	elseif Request.Form("Action") = "Add" then 
		call add_mode()
		intSites = Request.Form("sites")
		intCode = 0
	elseif Request.Form("Action") = "Update" then
		intCode = Request.Form("code") 
		strSql = "UPDATE sites " & _
  				 "SET strName = " & checkNull(Request.Form("name")) & ", strCoordinator = " & checkNull(Request.Form("coord")) & ", strAddress = " & checkNull(Request.Form("address")) & ", strCity = " & checkNull(Request.Form("city")) & ", intProvince = " & checkNull(Request.Form("province")) & ", strPostal = " & checkNull(Request.Form("postal")) & ", strPhone = " & checkNull(Request.Form("phone")) & ", strFax = " & checkNull(Request.Form("fax")) & ", strEmail = " & checkNull(Request.Form("email")) & ", strQ6 = " & checkNull(Request.Form("q6")) & ", strQ7 = " & checkNull(Request.Form("q7")) & ", strQ8 = " & checkNull(Request.Form("q8")) & ", strQ9 = " & checkNull(Request.Form("q9")) & ", strQ10 = " & checkNull(Request.Form("q10")) & ", strQ11 = " & checkNull(Request.Form("q11")) & ", strComments = " & checkNull(Request.Form("comments")) & " " & _
				 "WHERE intSiteID = " & intCode
		
		' update the record
		conn.execute strSql
		
		if conn.errors.count > 0 then 
			strError = "<font class=""regtextred"">Error Number : " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/></font>"
		end if 
	elseif Request.Form("Action") = "Save" then
		intCode = Request.Form("code") 
		strSQL = "INSERT INTO sites (intSiteID, strName, strCoordinator, strAddress, strCity, intProvince, strPostal, strPhone, strFax, strEmail, strQ6, strQ7, strQ8, strQ9, strQ10, strQ11, strComments) VALUES" & _
				 "(" & intCode & "," & checkNull(Request.Form("name")) & "," & checkNull(Request.Form("coord")) & "," & checkNull(Request.Form("address")) & "," & checkNull(Request.Form("city")) & "," & checkNull(Request.Form("province")) & "," & checkNull(Request.Form("postal")) & "," & checkNull(Request.Form("phone")) & "," & checkNull(Request.Form("fax")) & "," & checkNull(Request.Form("email")) & "," & checkNull(Request.Form("q6")) & "," & checkNull(Request.Form("q7")) & "," & checkNull(Request.Form("q8")) & "," & checkNull(Request.Form("q9")) & "," & checkNull(Request.Form("q10")) & "," & checkNull(Request.Form("q10")) & "," & checkNull(Request.Form("comments")) & ")"
				 
		' insert the record
		conn.execute strSql
		
		if conn.errors.count > 0 then 
			strError = "<font class=""regtextred"">Error Number : " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/></font>"
		end if 
	'elseif Request.Form("Action") = "Email" then 
		
	else 
		if Request.QueryString("site").Count > 0 then 
			intCode = Request.QueryString("site")
		else
			intCode = 0
		end if 
	end if 
	
	if Request.Form("Action") <> "Add" Then 
		' select the site just inserted
		rstData.Open "SELECT * FROM [sites] ORDER BY intSiteID", conn
	
		if not rstData.EOF then 
			' store info in array
			aData = rstData.GetRows 
								
			' get the total number of sites
			intSites = ubound(aData,2) + 1							
			
		'	Response.Write intCode
			' set values		
			call load_values(intCode)
		else
			intSites = 0 
			call add_mode
		end if		 
		
		' close the recordset
		rstData.Close 
	end if 
%>
<html>
<head>
	<!--<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />-->
	<!-- added UTF8 Encoding to get rid of funny characters -->
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" /> 
	<!-- Start CSS files-->
		<link rel="stylesheet" type="text/css" href="Styles/edi.css">
	<!-- End CSS files -->
	<script language="javascript" type="text/javascript" src="js/form.js"></script>
</head>
<body>
	<!-- #include virtual="/shared/page_header.inc" -->
	<%	
	' open edi connection
	call open_adodb(conn, "MACEDI")
	set rstProvinces = server.CreateObject("Adodb.recordset")
	%>
	<form name="Screens" method="POST" action="edi_admin_site.asp"> 
		<a class="reglinkMaroon" href="edi_admin.asp">Home</a>&nbsp;<font class="regtextblack">></font>&nbsp;<font class="boldtextblack">Site Information</font>
		<table border="1" bordercolor="006600" cellpadding="0" cellspacing="0" width="760">
		<tr><td>
			<table border="0" cellpadding="0" cellspacing="0" width="750s	" align="center">
			<tr>
				<td align="right" width="430"><font class="headerBlue">Site Information (<%=intSites%>)</font></td>
				<td align="right">
					<input type="hidden" name="sites" value="<%=intSites%>">
					<input type="hidden" name="Action" value="">
				<%
				if Request.form("Action") <> "Add" AND intSites > 0 then  
				%>
					<input type="button" value="Add" name="SubmitAction" title="ADD SITE" onClick="javascript:confirm_Add(this.value);">
					<input type="button" value="Delete" name="SubmitAction" title="DELETE SITE" onClick="javascript:confirm_Delete(this.value);">
					<%
					'if intSites > 1 then 
					'	Response.Write "<input type=""button"" value=""Find"" name=""Find"" title=""FIND SITE"">"
					'end if 
					%>
					<input type="button" value="Update" name="SubmitAction" title="UPDATE SITE" onClick="javascript:update_Check(this.value);">
				<%
				else
				%>
					<input type="button" value="Save" name="SubmitAction" title="SAVE SITE" onClick="javascript:update_Check(this.value);">
				<%
				end if 
				%>
					<input type="button" value="Exit" name="Exit" title="EXIT Screen" onClick="javascript:window.location='edi_admin.asp';">
					&nbsp;
				</td>
			</tr>
			<tr><td colspan="2"><%=strError%></td></tr>
			<!-- sections here-->
			</table>
			<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">
				<tr valign="top">
					<td align="right">
						<font class="boldtextblack">Site Code :&nbsp;&nbsp;</font>
					</td>
					<td>
						<%		
						if Request.Form("Action") = "Add" or intSites = 0 then 
							Response.Write "<input type=""text"" size=""30"" name=""code"" maxlength=""3"">"
						else
							' .selectedIndex - index
							Response.Write "<select name=""code"" onChange=""javascript:window.location='edi_admin_site.asp?site=' + this.value;"">"
							
							for intRow = 0 to ubound(aData,2)
								Response.Write "<option value = """ & right("000" & aData(0,intRow),3) & """"
								
								' if code is selected show it
								if cint(intSite) = aData(0,intRow) then 
									Response.write " selected"
								end if 
						
								Response.Write ">" & right("000" & aData(0,intRow),3) & "</option>"
							next
							Response.Write "</select>"
							
						end if 
						%>
					</td>
				</tr>
				<tr valign="top">
					<td align="right">
						<font class="boldtextblack">Site Name :&nbsp;&nbsp;</font>
					</td>
					<td>
						<input type="text" size="80" name="name" value="<%=strName%>">
					</td>
				</tr>
				<tr valign="top">
					<td align="right" nowrap>
						<font class="boldtextblack">Site Coordinator :&nbsp;&nbsp;</font>
					</td>
					<td>
						<input type="text" size="80" name="coord" value="<%=strCoordinator%>">
					</td>
				</tr>
				<tr valign="top">
					<td align="right">
						<font class="boldtextblack">Address :&nbsp;&nbsp;</font>
					</td>
					<td>
						<input type="text" size="80" name="address" value="<%=strAddress%>">
					</td>
				</tr>
				<tr valign="top">
					<td align="right">
						<font class="boldtextblack">City :&nbsp;&nbsp;</font>
					</td>
					<td>
						<input type="text" size="25" name="city" value="<%=strCity%>"> 
						<font class="boldtextblack">Province :&nbsp;&nbsp;</font>
						<select name="province">
							<option value=""></option>
						<%
						' build the options from the lookup table
						rstProvinces.Open "SELECT pid, english FROM [LU Provinces] ORDER BY english", conn
						
						do while not rstProvinces.eof						
							Response.Write "<option value = """ & rstProvinces("pid") & """"
							
							' if that province is selected than show it
							if intProvince = rstProvinces("pid") then 
								Response.write " selected"
							end if 
							
							' write the province name
							Response.Write ">" & rstProvinces("english") & "</option>"
							rstProvinces.MoveNext 
						loop
						
						' reset recordset 
						rstProvinces.MoveFirst 
						
						' put provinces in array
						aProvinces = rstProvinces.GetRows 
						
	'					' close and kill the provinces object
						call close_adodb(rstProvinces)
						%>
						</select>
						
						<font class="boldtextblack">Postal Code :&nbsp;&nbsp;</font>
						<input type="text" size="10" name="postal" value="<%=strPostal%>" maxlength="7" title="Enter postal code without dashes or spaces"> 
					</td>
				</tr>
				<tr valign="top">
					<td align="right">
						<font class="boldtextblack">Phone :&nbsp;&nbsp;</font>
					</td>
					<td>
						<input type="text" size="15" name="phone" value="<%=strPhone%>" maxlength="14" title="Enter phone number without brackets or spaces"> 
						<font class="boldtextblack">Fax :&nbsp;&nbsp;</font>
						<input type="text" size="15" name="fax" value="<%=strFax%>" maxlength="14" title="Enter fax number without brackets or spaces"> 
						<font class="boldtextblack">Email :&nbsp;&nbsp;</font>
						<input type="text" size="25" name="email" value="<%=strEmail%>">  
					</td>
				</tr>
			</table>
			<br />
			<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">
				<tr>
					<td width="15">&nbsp;</td>
					<td align="left">
						<font class="subheaderBlue">Additional site specific questions for Section E.</font> 
						<br />	
						<br />
					</td>
				</tr>
			</table>
			<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">
				<tr valign="top">
					<td align="right">
						<font class="boldtextblack">Question 6 :&nbsp;&nbsp;</font>
					</td>
					<td>
						<input type="text" size="90" name="q6" value="<%=strQ6%>" />
					</td>
				</tr>
				<tr valign="top">
					<td align="right">
						<font class="boldtextblack">Question 7 :&nbsp;&nbsp;</font>
					</td>
					<td>
						<input type="text" size="90" name="q7" value="<%=strQ7%>" />
					</td>
				</tr>
				<tr valign="top">
					<td align="right">
						<font class="boldtextblack">Question 8 :&nbsp;&nbsp;</font>
					</td>
					<td>
						<input type="text" size="90" name="q8" value="<%=strQ8%>" />
					</td>
				</tr>
				<tr valign="top">
					<td align="right">
						<font class="boldtextblack">Question 9 :&nbsp;&nbsp;</font>
					</td>
					<td>
						<input type="text" size="90" name="q9" value="<%=strQ9%>" />
					</td>
				</tr>
				<tr valign="top">
					<td align="right">
						<font class="boldtextblack">Question 10 :&nbsp;&nbsp;</font>
					</td>
					<td>
						<input type="text" size="90" name="q10" value="<%=strQ10%>" />
					</td>
				</tr>
				<tr valign="top">
					<td align="right">
						<font class="boldtextblack">Question 11 :&nbsp;&nbsp;</font>
					</td>
					<td>
						<input type="text" size="90" name="q11" value="<%=strQ11%>" />
					</td>
				</tr>
				<tr valign="top">
					<td align="right">
						<font class="boldtextblack">Comments :&nbsp;&nbsp;</font>
					</td>
					<td>
						<textarea rows="3" cols="78" name="comments"><%=strComments%></textarea>
					</td>
				</tr>
			</table>
			<%
			' only show schools if on a site not in add mode
			if Request.Form("Action") <> "Add" AND intSites > 0 then 
			%>
			<hr>
			<br />
			<table border="1" cellpadding="0" cellspacing="0" width="750" align="center">
				<tr>
					<td align="center"><font class="boldtextblack">School ID</font></td>
					<td align="center"><font class="boldtextblack">Name</font></td>
					<td align="center"><font class="boldtextblack">City</font></td>
					<td align="center"><font class="boldtextblack">Province</font></td>
				</tr>
				<%
				' select all schools at this site 
				rstData.Open "SELECT intSchoolid, strName, strCity, intProvince FROM schools WHERE intSiteID = " & intSite & " ORDER BY intSchoolID", conn
				
				if rstData.EOF then 
					Response.Write "<tr><td colspan=""4"">&nbsp;<font class=""regtextmaroon"">There are no schools at this site.</font></td></tr>"
				else
					do while not rstData.EOF 
						Response.Write "<tr><td><a href=""edi_admin_school.asp?site=" & right("000"&intSite,3) & "&school=" & right("000" & rstData("intSchoolID"),3) & """ class=""reglinkBlue"">" & right("000000" & rstData("intSchoolID"),6) & "</a></td>"
						Response.Write "<td><font class=""regtextblack"">" & rstData("strName") & "</font></td>"
						Response.Write "<td><font class=""regtextblack"">" & rstData("strCity") & "</font></td>"
						Response.Write "<td><font class=""regtextblack"">" 
						' get the province & " " & ubound(aProvinces,2)
						for introw = 0 to ubound(aProvinces,2)
							if rstData("intProvince") = aProvinces(0,introw) then 
								Response.Write aProvinces(1,intRow)
								exit for
							end if
						next
						
						Response.Write "</font></td></tr>"
						rstData.MoveNext 
					loop
				end if 
				%>
			</table>
			<%
			end if 
			%>
			<br />
			</td>
		</tr>
		</table>
	</form>
	
	<!-- #include virtual="/shared/page_footer.inc" -->
</body>
</html>
<%
	' close and kill recordset and connection
	call close_adodb(rstData)
	'call close_adodb(conn)
	call close_adodb(conn)
end if

' set form defaults
sub add_mode()
	strName = ""
	strCoordinator = "" 
	strAddress = ""
	strCity = ""
	intProvince = 1
	strPostal = ""
	strPhone = "" 
	strFax = ""
	strEmail = "" 
	strQ6 = "" 
	strQ7 = "" 
	strQ8 = ""
	strQ9 = "" 
	strQ10 = ""
	strQ11 = ""
	strComments = ""
end sub

sub load_values(intCode)	
	if intCode = 0 then 
		introw = 0 
	else
		for introw = 0 to ubound(aData,2)
			if cint(intCode) = aData(0,introw) then 
				exit for
			end if
		next

		if intRow > ubound(aData,2) then 
			intRow = 0 
		end if 	
	end if 
	
		' set values		
		intSite = aData(0,introw)
		strName = aData(1,introw)
		strCoordinator = aData(2,introw)
		strAddress = aData(3,introw)
		strCity = aData(4,introw)
		intProvince = aData(5,introw)
		strPostal = aData(6,introw)
		strPhone = aData(7,introw) 
		strFax = aData(8,introw)
		strEmail = aData(9,introw) 
		strQ6 = aData(10,introw) 
		strQ7 = aData(11,introw) 
		strQ8 = aData(12,introw)
		strQ9 = aData(13,introw) 
		strQ10 = aData(14,introw)
		strQ11 = aData(15,introw)
		strComments = aData(16,introw)
end sub
%>
