<!-- #include virtual="/shared/admin_security.asp" -->
<%
' public variables
dim intSchool, intSite
dim strName, strCoordinator, strAddress, strCity, intProvince, strPostal,	strPhone, strFax, strEmail, strQ6, strQ7, strQ8, strQ9,	strQ10, strComments,intELP
dim aData, aSites
on error resume next

' if the user has not logged in they will not be able to see the page
if blnSecurity then
	'call open_adodb(conn, "DATA")
	call open_adodb(conn, "MACEDI")
	
	set rstData = server.CreateObject("adodb.recordset")
	set rstLanguages = server.CreateObject("adodb.recordset")
	
	' open all languages
	rstLanguages.Open "SELECT LID, english FROM [LU Languages] ORDER BY english", conn
	
	' store all languages in array
	aLanguages = rstLanguages.GetRows 
	
	' close and kill the langauges recordset
	call close_adodb(rstLanguages)
	
	' delete record - December 16
	if Request.Form("Action") = "Delete" then 
		intSite = Request.form("site")
		' delete the unique school 
		strSql = "DELETE FROM schools WHERE intSchoolID = " & intSite & Request.Form("code")
		
		' execute the sql
		conn.execute strSql 	
		
		' build the error string
		if conn.errors.count > 0 then 
			strError = "<font class=""regtextred"">Error Number : " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/></font>"
		end if

	' Add a site - loads an empty form 
	'		     - set all values = ""
	elseif Request.Form("Action") = "Add" then 
		intSite = Request.Form("site")
		call add_mode()
		
	' update records - December 16
	elseif Request.Form("Action") = "Update" then
		intSite = Request.Form("site")
		intSchool = Request.Form("code") 

        strSql = "UPDATE schools " & _
  				 "SET strName = " & checkNull(Request.Form("name")) & ", strCity = " & checkNull(Request.Form("city")) & ", intProvince = " & checkNull(Request.Form("province")) & ", strPostal = " & checkNull(Request.Form("postal")) & ", strPhone = " & checkNull(Request.Form("phone")) & ", strFax = " & checkNull(Request.Form("fax")) & ", strEmail = " & checkNull(Request.Form("email")) & ", strComments = " & checkNull(Request.Form("comments")) 

		if request.Form("ELP") = "on" then
			strSql = strSql & ",intELP=1"
		else
			strSql = strSql & ",intELP=0"
		end if 
		' build the SQL statement
		
	    strSql = strSql & " WHERE intSchoolID = " & intSite & intSchool
		
		' update the record
		conn.execute strSql
		
		' build the error string
		if conn.errors.count > 0 then 
			strError = "<font class=""regtextred"">Error Number : " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/></font>"
		end if
	' insert record - December 16	
	elseif Request.Form("Action") = "Save" then
		intSite = Request.Form("site") 
		intSchool = Request.Form("code")
		if request.Form("ELP") = "on" then
			strELPSQL = "1"
		else
			strELPSQL = "0"
		end if 

		strSQL = "INSERT INTO schools (intSchoolID,intSiteID, strName, strCity, intProvince, strPostal, strPhone, strFax, strEmail, strComments,intELP) VALUES" & _
				 "(" & intSite & intSchool & "," & intSite & "," & checkNull(Request.Form("name")) & "," & checkNull(Request.Form("city")) & "," & checkNull(Request.Form("province")) & "," & checkNull(Request.Form("postal")) & "," & checkNull(Request.Form("phone")) & "," & checkNull(Request.Form("fax")) & "," & checkNull(Request.Form("email")) & "," & checkNull(Request.Form("comments")) & "," & strELPSQL & ")"
	    
		' insert the record
		conn.execute strSql
		
		if conn.errors.count > 0 then 
			strError = "<font class=""regtextred"">Error Number : " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/></font>"
		end if 
	' if there is an email present
	elseif Request.Form("hiddenAction") <> "" AND Request.Form("hiddenAction").Count > 0 then 
		intSite = Request.Form("site")
		intSchool = Request.Form("code") 
		
		' put the to email address in a variable
		strEmail = Request.form("hiddenAction")
		
		
		set rstEmail = server.CreateObject("adodb.recordset")
	
		' find the record with the email address in it	
		strSql = "SELECT name, email, password FROM users WHERE email ='" & strEmail & "'"	
	
		' open the recordset
		rstEmail.Open strSql, conn
			
		' build the email	
		htmltext ="<html><head><title>Password</title></head><body><center><img src=""http://www.e-edi.ca/images/e-edi.gif"" alt=""e-EDI"" name=""e-edi.gif""><br><br><font color=""black"">Your username at e-EDI is: <b>" & strEmail & "</b><br /><br /><font color=""black"">Your password at e-EDI is: <b>" & rstEmail("password") & "</b></font></center></body></html>"
	
		' close the recordset
		call close_adodb(rstEmail)
			
		' send the email
		set objmail = server.CreateObject("CDONTS.NewMail")
			objmail.From = "webmaster@e-edi.ca"
			objmail.To = strEmail
			objmail.Subject = "e-EDI Password"
			objmail.BodyFormat = 0
			objmail.MailFormat = 0
			objmail.Body = htmlText
			objmail.Send 
		set objmail = nothing		
			
		' build the reminder string
		strReminder = "<font class=""boldTextBlack"">Your password reminder has been sent to " & strEmail & ".</font>"
	end if 
		
	' get all sites
	rstData.Open "SELECT intSiteID FROM [sites] ORDER BY intSiteID", conn
	
	if not rstData.EOF then 
		' store info in array
		aSites = rstData.GetRows 
						
		' close the recordset 
		rstData.Close 
					
		' get the total number of sites
		intSites = ubound(aSites,2) + 1							
	else
		intSites = 0 
	end if 
		
	' first time to page - load defaults
	if Request.Form("Action") <> "Add" AND intSites > 0 Then 
		' get the total number of schools
		rstData.Open "SELECT COUNT(intSchoolID) FROM schools", conn
	
		' if more than 0 schools
		if not rstData.EOF then 
			' get the total number of schools
			intTotalSchools = rstData(0)
				
			' close the recordset
			rstData.Close
				
			if Request.QueryString("site").Count = 0 then 
				' this will have the current site if updated
				if intSite = "" then 
					intSite = aSites(0,0) 
				end if 
			else
				intSite = Request.QueryString("site")
			end if 
				
			strSql = "SELECT * FROM [schools] WHERE intSiteID = " & intSite & " ORDER BY intSchoolID"
			' get the site specific schools
			rstData.Open strSql, conn	
				
			if not rstData.EOF then
				' store info in array
				aData = rstData.GetRows 
										
				' get the total number of schools
				intSchools = ubound(aData,2) + 1							
					
				' get the school
				if Request.QueryString("school").Count = 0 then		
					' this will have the current school if updated
					if intSchool = ""  then 
						intSchool = right(aData(0,0),3)
					end if 
				else
					intSchool = Request.QueryString("school")
				end if
				
				' load the values
			    call load_values(intSite & intSchool)
			else
				intSchools = 0
				call add_mode
			end if 
		' if 0 schools
		else
			' 0 schools in the database
			intTotalSchools = 0 
			call add_mode
		end if		 	
		
		' close the recordset
		rstData.Close 
	' add mode
	else
		' get total number of schools
		intTotalSchools = Request.Form("schools")
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
	' provinces recordset
	set rstProvinces = server.CreateObject("Adodb.recordset")
	%>
	<form name="Screens" method="POST" action="edi_admin_school.asp"> 
		<a class="reglinkMaroon" href="edi_admin.asp">Home</a>&nbsp;<font class="regtextblack">></font>&nbsp;<font class="boldtextblack">School Information</font>
		<table border="1" bordercolor="006600" cellpadding="0" cellspacing="0" width="760">
		<tr><td>
			<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">
			<tr>
				<td align="right" width="430"><font class="headerBlue">School Information (<%=intTotalSchools%>)</font></td>
				<td align="right">
					<input type="hidden" name="hiddenAction" value="">
					<input type="hidden" name="Action" value="">
					<input type="hidden" name="schools" value="<%=intTotalSchools%>">
				<%
				' checks to be sure that there are active sites
				' doesn't allow administration if no sites
				if intSites = 0 then 
					Response.Write "</td></tr><tr><td colspan=""2""><br /></td></tr><tr><td colspan=""2"" align=""left"">"
					Response.Write "<font class=""regtextred"">Please add a site before attempting to administer schools.</font>" 		
					Response.Write "</td></tr><tr><td colspan=""2""><br /></td></tr></table>"
					Response.Write "</td></tr></table>"
					Response.Write "</form>"
				else
					' if there are no schools then automatically in add mode
					' also add mode if user chooses to add
					if Request.form("Action") <> "Add" AND intSchools > 0 then  
					%>
						<input type="button" value="Add" name="SubmitAction" title="ADD SITE" onClick="javascript:confirm_Add(this.value);">
						<input type="button" value="Delete" name="SubmitAction" title="DELETE SCHOOL" onClick="javascript:confirm_Delete(this.value);">
						<%
						'if intSchools > 1 then 
						'	Response.Write "<input type=""button"" value=""Find"" name=""Find"" title=""FIND SITE"">"
						'end if 
						%>
						<input type="button" value="Update" name="SubmitAction" title="UPDATE SCHOOL" onClick="javascript:update_Check(this.value);">
					<%
					else
					%>
						<input type="button" value="Cancel" name="Cancel" title="Go to this sites schools" onClick="javascript:window.location='edi_admin_school.asp?site=' + document.forms.Screens.site.value;">
						<input type="button" value="Save" name="SubmitAction" title="SAVE SCHOOL" onClick="javascript:update_Check(this.value);">
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
						<font class="boldtextblack">School Code :&nbsp;&nbsp;</font>
					</td>
					<td>
						<%		
						' if there are no schools then automatically in add mode
						' also add mode if user chooses to add
						' they choose the active site and then enter the school code
						if Request.Form("Action") = "Add" or intSchools = 0 then 
							' .selectedIndex - index - 
							Response.Write "<select name=""site"">"
							
							for intRow = 0 to ubound(aSites,2)
								Response.Write "<option value = """ & aSites(0,intRow) & """"
								
								' if code is selected show it
								if cint(intSite) = aSites(0,intRow) then 
									Response.write " selected"
								end if 
						
								Response.Write ">" & right("000" & aSites(0,intRow),3) & "</option>"
							next
							Response.Write "</select>"
							Response.Write "<input type=""text"" size=""15"" name=""code"" maxlength=""3"" title=""Enter the 3 digit school code"">"
						' not add mode
						else
							' build site selection box - they select a site - load all schools in that site
							Response.Write "<select name=""site"" onChange=""javascript:window.location='edi_admin_school.asp?site=' + this.value;"">"
							
							for intRow = 0 to ubound(aSites,2)
								Response.Write "<option value = """ & right("000" & aSites(0,intRow),3) & """"
								
								' if code is selected show it
								if cint(intSite) = aSites(0,intRow) then 
									Response.write " selected"
								end if 
						
								Response.Write ">" & right("000" & aSites(0,intRow),3) & "</option>"
							next
							Response.Write "</select>"
							
							' build the list of schools at this site -  + '&code=' + this.value
							Response.Write "<select name=""code"" onChange=""javascript:window.location='edi_admin_school.asp?site=' + document.forms.Screens.site.value + '&school=' + this.options[this.selectedIndex].text;"">"
							
							for intRow = 0 to ubound(aData,2)
								Response.Write "<option value = """ & right(aData(0,intRow),3) & """"
								
								' if school is selected show it
								if intSchool = right(aData(0,intRow),3) then 
									Response.write " selected"
								end if 
						
								Response.Write ">" & right(aData(0,intRow),3) & "</option>"
							next
							Response.Write "</select>"
							Response.Write "&nbsp;<font class=""regtextgreen"">" & intSchools & " School"
							
							' if more than one school - plural
							if intSchools > 1 then 
								Response.Write "s"
							end if 
							Response.Write " at this site</font>"
						end if 
						%>
					</td>
				</tr>
				<tr valign="top">
					<td align="right">
						<font class="boldtextblack">School Name :&nbsp;&nbsp;</font>
					</td>
					<td>
						<input type="text" size="80" name="name" value="<%=strName%>">
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
				<tr valign="top">
					<td align="right">
						<font class="boldtextblack">Comments :&nbsp;&nbsp;</font>
					</td>
					<td>
						<textarea rows="3" cols="70" name="comments"><%=strComments%></textarea>
					</td>
				</tr>
                <tr>
                    <td align="right"><font class="boldtextblack">ELP</font></td>
                    <td>                                                                                                    
                        <%
                        response.write "<input type=""checkbox"" id=""ELP"" name=""ELP"""
                        if intELP = "True" then 
                            response.write " checked=""CHECKED"""
                        end if 
                        response.write "/>"
                        %>						
                    </td>
                </tr>
			</table>
			<%
			' only show schools if on a site not in add mode
			if Request.Form("Action") <> "Add" AND intSchools > 0 then 
			%>
			<hr>
			<br />
			<%
			if Request.Form("hiddenAction") <> "" then  
				Response.Write "<p align=""center"">" & strReminder & "</p>"
			end if 
			
			%>
			<table border="1" cellpadding="0" cellspacing="0" width="750" align="center">
				<tr>
					<td align="center"><font class="boldtextblack">Teacher ID</font></td>
					<td align="center"><font class="boldtextblack">Teacher Name</font></td>
					<td align="center"><font class="boldtextblack">Email</font></td>
					<!--<td align="center"><font class="boldtextblack">Size</font></td>
					<td align="center"><font class="boldtextblack">Completed</font></td>-->
					<td align="center"><font class="boldtextblack">Password</font></td>
				</tr>
				<%
				' select all classes at the school 				
				' strSql = "SELECT c.intClassID, t.strName, t.strEmail, c.intLanguage, c.intStudents, Count(ch.chkCompleted) AS Completed " & _
				'		 "FROM (teachers t RIGHT JOIN classes c ON t.intTeacherID = c.intTeacherID) LEFT JOIN children ch ON c.intClassID = ch.intClassID " & _
				'		 "WHERE t.intSchoolID = " & intSite & intSchool & _
				'		 " GROUP BY c.intClassID, t.strName, t.strEmail, c.intLanguage, c.intStudents"
				
				strSql = "SELECT t.intTeacherID, t.strName, t.strEmail " & _
						 "FROM teachers t " & _
						 "WHERE t.intSchoolID = " & intSite & intSchool & _
						 " GROUP BY t.intTeacherID, t.strName, t.strEmail"
				
				
				'Response.Write strSQL
				
				' open list of classes and teachers at this school
				rstData.Open strSql, conn				
				if rstData.EOF then 
					Response.Write "<tr><td colspan=""7"">&nbsp;<font class=""regtextmaroon"">There are no teachers at this school.</font></td></tr>"
				else
					do while not rstData.EOF 
						'Response.Write "<tr><td><a href=""edi_admin_classandchild.asp?site=" & left(right("000" & rstData("intClassID"),9),3) & "&school=" & mid(right("000" & rstData("intClassID"),9),4,3) & "&teacher=" & mid(right("000" & rstData("intClassID"),9),7,2) & "&class=" & right(rstData("intClassID"),1) & """ class=""reglinkBlue"">" & right("000" & rstData("intClassID"),9) & "</a></td>"
						Response.Write "<tr><td><a href=""edi_admin_teacher.asp?site=" & left(right("000" & rstData("intTeacherID"),8),3) & "&school=" & mid(right("000" & rstData("intTeacherID"),8),4,3) & "&teacher=" & mid(right("000" & rstData("intTeacherID"),8),7,2) & """ class=""reglinkBlue"">" & right("000" & rstData("intTeacherID"),8) & "</a></td>"
						Response.Write "<td><font class=""regtextblack"">" & rstData("strName") & "</font></td>"
						Response.Write "<td><font class=""regtextblack"">" & rstData("strEmail") & "</font></td>"
						'Response.Write "<td><font class=""regtextblack"">" 
						
						' get the language
						'for introw = 0 to ubound(aLanguages,2)
						'	if rstData("intLanguage") = aLanguages(0,introw) then 
						'		Response.Write aLanguages(1,intRow)
						'		exit for
						'	end if
						'next
						'Response.Write  "</font></td>"
						'Response.Write "<td><font class=""regtextblack"">" & rstData("intStudents") & "</font></td>"
						'Response.Write "<td><font class=""regtextblack"">" & rstData("Completed") & "</font></td>"
						Response.Write "<td><a href=""javascript:email_Password('" & rstData("strEmail") & "');"" class=""reglinkBlue"">Send Password Reminder</a></td></tr>"
						rstData.MoveNext 
					loop
				end if 
				%>
			</table>
			<%
			' not add
			end if 
			%>
			<br />
			</td>
		</tr>
		</table>
	</form>
	<%
	' schools	
	end if
	%>	
	<!-- #include virtual="/shared/page_footer.inc" -->
</body>
</html>
<%
	' close and kill recordset and connection
	call close_adodb(rstData)
	'call close_adodb(conn)
	call close_adodb(conn)
' security
end if

' set form defaults
sub add_mode()
	' load the first site
	intSchool = ""
	strName = ""
'	strAddress = ""
	strCity = ""
	intProvince = 1
	strPostal = ""
	strPhone = "" 
	strFax = ""
	strEmail = "" 
	strComments = ""
    intELP = 0
end sub

sub load_values(intSchool)	
	if intSite = 0 then 
		introw = 0 
	else
		for introw = 0 to ubound(aData,2)
			if clng(intSchool) = aData(0,introw) then 
				exit for
			end if
		next

		if intRow > ubound(aData,2) then 
			intRow = 0 
		end if 	
	end if 
	
	' set values		
	intSite = aData(1,introw)
	intSchool = right(aData(0,introw),3)
	strName = aData(2,introw)
'	strPrincipal = aData(3,introw)
'	strAddress = aData(4,introw)
	strCity = aData(5,introw)
	intProvince = aData(6,introw)
	strPostal = aData(7,introw)
	strPhone = aData(8,introw) 
	strFax = aData(9,introw)
	strEmail = aData(10,introw) 
	strComments = aData(11,introw)
    intELP = aData(12,introw)    
end sub
%>
