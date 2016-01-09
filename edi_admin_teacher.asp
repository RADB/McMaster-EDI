<!-- #include virtual="/shared/admin_security.asp" -->
<%
' public variables
dim intTotalTeachers 
dim aData
dim strName, strEmail,intSex, intAge, strPhone, strFax, intQ5a, intQ5b,intQ5c, intQ6a, intQ6b, intQ6c, intQ6d, intQ6e, intQ6f, intQ6g, intQ6h, intQ6i, intQ6j, intQ6k
dim intMth1, intMth2, intMth3,intMth4, intMth5, intMth6,intMth7, intMth8, intMth9, intYr1, intYr2, intYr3,intYr4, intYr5, intYr6,intYr7, intYr8, intYr9
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
	
	' delete record
	if Request.Form("Action") = "Delete" then 
		intSite = Request.form("site")
		intSchool = Request.form("school")
		
		' delete the unique teacher
		strSql = "DELETE FROM teachers WHERE intTeacherID = " & intSite & intSchool & Request.Form("code")
		
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
		intSchool = Request.Form("school")
		call add_mode()
		
	' update records - December 16
	elseif Request.Form("Action") = "Update" then
		intSite = Request.Form("site")
		intSchool = Request.Form("school")
		intTeacher = Request.Form("code") 
			
		' build the SQL statement
		strSql = "UPDATE teachers " & _
  				 "SET strName = " & checkNull(Request.Form("name")) & ", strEmail = " & checkNull(Request.Form("email")) & ", strPhone = " & checkNull(Request.Form("phone")) & ", strFax = " & checkNull(Request.Form("fax")) & ", intSex = " & checkNull(Request.Form("sex")) & ", intAge = " & checkNull(Request.Form("age")) & ", intQ5a = " & checkNull(Request.Form("intQ5a")) & ",intQ5b = " & checkNull(Request.Form("intQ5b")) & ",intQ5c = " & checkNull(Request.Form("intQ5c")) & ",intQ6a = " & checkNull(Request.Form("intQ6a")) & ", intQ6b = " & checkNull(Request.Form("intQ6b")) & ",intQ6c = " & checkNull(Request.Form("intQ6c")) & ",intQ6d = " & checkNull(Request.Form("intQ6d")) & ",intQ6e = " & checkNull(Request.Form("intQ6e")) & ",intQ6f = " & checkNull(Request.Form("intQ6f")) & ",intQ6g = " & checkNull(Request.Form("intQ6g")) & ",intQ6h = " & checkNull(Request.Form("intQ6h")) & ",intQ6i = " & checkNull(Request.Form("intQ6i")) & ",intQ6j = " & checkNull(Request.Form("intQ6j")) & ",intQ6k = " & checkNull(Request.Form("intQ6k")) & _
				 " WHERE intTeacherID = " & intSite & intSchool & intTeacher
	
		'Response.Write strSql
		' update the record
		conn.execute strSql
		
		' build the error string
		if conn.errors.count > 0 then 
			strError = "<font class=""regtextred"">Error Number : " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/></font>"
		else
			for introw = 1 to Request.form("intClasses")
				' extract the months from the form 
				intInsertMonths = right(Request.Form("intQ5d" & intRow + 3),len(Request.Form("intQ5d" & intRow + 3))-9)
				
				' update the classes
				strSQL = "UPDATE classes SET intMonths =" & intInsertMonths & " WHERE intClassID = " &  left(Request.Form("intQ5d" & intRow + 3),9)
				conn.execute strSql
			next 
		end if
	' insert record - December 16	
	elseif Request.Form("Action") = "Save" then
		intSite = Request.Form("site") 
		intSchool = Request.Form("school")
		intTeacher = Request.Form("code")
		
		strSQL = "INSERT INTO teachers (intTeacherID,intSchoolID,strName, strEmail, strPassword, strPhone, strFax, intSex, intAge, intQ5a, intQ5b, intQ5c, intQ6a, intQ6b, intQ6c, intQ6d, intQ6e, intQ6f, intQ6g, intQ6h, intQ6i, intQ6j, intQ6k ) VALUES" & _
				 "(" & intSite & intSchool & intTeacher & "," & intSite & intSchool & "," & checkNull(Request.Form("name")) & "," & checkNull(Request.Form("email")) &"," & checkNull(right("000" &intSite & intSchool & intTeacher,8)) & "," & checkNull(Request.Form("phone")) & "," & checkNull(Request.Form("fax")) & "," & checknull(Request.Form("sex")) & "," & checkNull(Request.Form("age")) & "," & checkNull(request.form("intQ5a")) & "," & checkNull(request.form("intQ5b")) & "," & checkNull(request.form("intQ5c")) & "," & checkNull(request.form("intQ6a")) & "," & checkNull(request.form("intQ6b")) & "," & checkNull(request.form("intQ6c")) & "," & checkNull(request.form("intQ6d")) & "," & checkNull(request.form("intQ6e")) & "," & checkNull(request.form("intQ6f")) & "," & checkNull(request.form("intQ6g")) & "," & checkNull(request.form("intQ6h")) & "," & checkNull(request.form("intQ6i")) & "," & checkNull(request.form("intQ6j")) & "," & checkNull(request.form("intQ6k")) & ")"
					 
		' insert the record
		conn.execute strSql
		
		if conn.errors.count > 0 then 
			strError = "<font class=""regtextred"">Error Number : " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/></font>" 
		end if 
		
		strSQL = "INSERT INTO teacherParticipation (intTeacherID) VALUES" & _
				 "(" & intSite & intSchool & intTeacher & ")"
					 
		' insert the record
		conn.execute strSql
		
		if conn.errors.count > 0 then 
			strError = "<font class=""regtextred"">Error Number : " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/></font>" 
		end if 

        strSQL = "INSERT INTO teacherTrainingFeedback (intTeacherID) VALUES" & _
				 "(" & intSite & intSchool & intTeacher & ")"
					 
		' insert the record
		conn.execute strSql
		
		if conn.errors.count > 0 then 
			strError = "<font class=""regtextred"">Error Number : " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/></font>" 
		end if 
	end if 

	'////////////////////////////////////////////
	' 
	' Get all sites and number of sites - aSites & intSites
	' Get all schools and number of schools - aSchools & intSchools
	'  	
	'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
	
	' get all Schools
	rstData.Open "SELECT DISTINCT intSchoolID, intSiteID FROM [schools] ORDER BY intSchoolID", conn
	
	if not rstData.EOF then 
		' store info in array
		aSchools = rstData.GetRows 
						
		' close the recordset 
		rstData.Close 
					
		' get the total number of schools
		intSchools = ubound(aSchools,2) + 1							
	else
		intSchools = 0 
	end if 
	
	' get all sites - used for site list 
	rstData.Open "SELECT DISTINCT intSiteID FROM [schools] ORDER BY intSiteID", conn
	
	if not rstData.EOF then 
		' store info in array
		aSites = rstData.GetRows 
						
		' close the recordset 
		rstData.Close 
					
		' get the total number of sites with schools
		intSites = ubound(aSites,2) + 1							
	else
		intSites = 0 
	end if 
	
	'///////////////////////////////////////	
	' Author - Andrew Renner
	' Description - FIRST TIME TO PAGE
	' Details:	A) Get total number of teachers  - intTotalTeachers if their are teachers then continue
	'					1) Check querystring for a site 
	'						i) if site is there then save the site
	'						ii) if no site then get the first site with a school
	'						iii) get all schools at the specific site
	'							a) count the number of schools at this site
	'							b) if none then exit telling user that there must be a school to add a teacher
	'					2) Check querystring for a school
	'						i) if school is there then save the school
	'						ii) if no school then get the first school 
	'						iii) get all teachers at the specified school
	'							a) count the number of teachers at this school
	'							b) if none then go into add mode		
	'					3) Check querystring for a teacher
	'						i) if teacher is there then save the teacher
	'						ii) if no teacher then get all teachers at the school
	'						iii) load the first teachers values
	'				B) if no teachers but there are schools then enter add mode
	'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
	if Request.Form("Action") <> "Add" AND intSchools > 0 Then 
		' get the total number of teachers
		rstData.Open "SELECT COUNT(intTeacherID) FROM teachers", conn
	
		' if more than 0 teachers
		if not rstData.EOF then 
			' get the total number of teachers
			intTotalTeachers = rstData(0)
				
			' close the recordset
			rstData.Close
				
			' check querystring for site
			if Request.QueryString("site").Count = 0 then 
				' this will have the current site if updated
				if intSite = "" then 
					' get the site from the schools array
					intSite = right("000" & aSchools(1,0),3)
				end if 
			else
				intSite = Request.QueryString("site")
			end if 
			
			strSql = "SELECT DISTINCT intSchoolID FROM [schools] WHERE intSiteID = " & intSite & " ORDER BY intSchoolID"
			' get the site specific schools
			rstData.Open strSql, conn	
				
			if not rstData.EOF then
				' store info in array
				aSchools = rstData.GetRows 
				
				' get the total number of schools at this site
				intSchools = ubound(aSchools,2) + 1							
					
				' get the school
				if Request.QueryString("school").Count = 0 then		
					' this will have the current school if updated
					if intSchool = ""  then 
						intSchool = right(aSchools(0,0),3)
					end if 
				else
					intSchool = Request.QueryString("school")
				end if
				
				' close the recordset
				rstData.Close			
				
	
				strSql = "SELECT DISTINCT * FROM [teachers] WHERE intSchoolID = " & intSite & intSchool & " ORDER BY intTeacherID"
	
				' get the school specific teachers
				rstData.Open strSql, conn	
								
				if not rstData.EOF then
					' store info in array
					aData = rstData.GetRows 
											
					' get the number of teachers at this school
					intTeachers = ubound(aData,2) + 1							
							
					' get the teacher
					if Request.QueryString("teacher").Count = 0 then		
						' this will have the current teacher if updated
						if intTeacher = ""  then 
							intTeacher = right(aData(0,0),2)
						end if 
					else
						intTeacher = Request.QueryString("teacher")
					end if
					
					' load the values
				   call load_values(intSite & intSchool & intTeacher)
				else
					intTeachers = 0
					call add_mode
				end if 
			else
				intSchools = 0
				intTeachers = 0
			end if 
		' if 0 teachers
		else
			' 0 teachers in the database
			intTotalTeachers = 0 
			call add_mode
		end if		 	
		
		' close the recordset
		rstData.Close 
		
		' select all classes that this teacher has  (0) None Selected (1) English (2) French (3) Other				
		'strSql = "SELECT c.intClassID, iif(c.intLanguage=1, 'English',iif(c.intLanguage=2, 'French',iif(c.intLanguage=3, 'Other','Unknown'))) as strLanguage, count(ch.strEDIID) as intStudents, sum(iif(ch.chkCompleted=true,1,0)) AS Completed, int(c.intMonths/12) as years, (c.intMonths mod 12) as months " & _
		'			"FROM classes c LEFT JOIN children ch ON c.intClassID = ch.intClassID " & _
		'			"WHERE c.intTeacherID = " & intSite & intSchool & intTeacher & _
		'			" GROUP BY c.intClassID, iif(c.intLanguage=1, 'English',iif(c.intLanguage=2, 'French',iif(c.intLanguage=3, 'Other','Unknown'))), int(c.intMonths/12), (c.intMonths mod 12)"
        strSql = "SELECT intClassID, strLanguage, intStudents, completed,years, months FROM GetTeacherClasses WHERE intTeacherID = " & intSite & intSchool & intTeacher

		'Response.Write strSQL
				
		' open list of classes and teachers at this school
		rstData.Open strSql, conn,1
		
		if rstData.EOF then 
			intClasses = 0
		else
			intClasses = rstData.RecordCount 
		end if 
	' add mode
	else
		' get total number of teachers
		intTotalTeachers = Request.Form("teachers")
		
		' get all Schools at the selected site
		rstData.Open "SELECT DISTINCT intSchoolID FROM [schools] WHERE intSiteID = " & Request.Form("site") & " ORDER BY intSchoolID", conn
		
		if not rstData.EOF then 
			' store info in array
			aSchools = rstData.GetRows 
							
			' close the recordset 
			rstData.Close  
		end if 
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
	<form name="Screens" method="POST" action="edi_admin_teacher.asp"> 
		<a class="reglinkMaroon" href="edi_admin.asp">Home</a>&nbsp;<font class="regtextblack">></font>&nbsp;<font class="boldtextblack">Teacher Information</font>
		<table border="1" bordercolor="006600" cellpadding="0" cellspacing="0" width="760">
		<tr><td>
			<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">
			<tr>
				<td align="right" width="430"><font class="headerBlue">Teacher Information (<%=intTotalTeachers%>)</font></td>
				<td align="right">
					<input type="hidden" name="Action" value="">
					<input type="hidden" name="teachers" value="<%=intTotalTeachers%>">
				<%
				' checks to be sure that there are active schools
				' doesn't allow administration if no schools
				if intSchools = 0 then 
					Response.Write "</td></tr><tr><td colspan=""2""><br /></td></tr><tr><td colspan=""2"" align=""left"">"
					Response.Write "<font class=""regtextred"">Please add a school before attempting to administer teachers.</font>" 		
					Response.Write "</td></tr><tr><td colspan=""2""><br /></td></tr></table>"
					Response.Write "</td></tr></table>"
					Response.Write "</form>"
				else
					' if there are no schools then automatically in add mode
					' also add mode if user chooses to add
					if Request.form("Action") <> "Add" AND intTeachers > 0 then  
					%>
						<input type="button" value="Add" name="SubmitAction" title="ADD TEACHER" onClick="javascript:confirm_Add(this.value);">
						<input type="button" value="Delete" name="SubmitAction" title="DELETE TEACHER" onClick="javascript:confirm_Delete(this.value);">
						<%
						'if intSchools > 1 then 
						'	Response.Write "<input type=""button"" value=""Find"" name=""Find"" title=""FIND SITE"">"
						'end if 
						%>
						<input type="button" value="Update" name="SubmitAction" title="UPDATE TEACHER" onClick="javascript:update_TeacherCheck(this.value);">
					<%
					else
					%>
						<input type="button" value="Cancel" name="Cancel" title="Go to this schools teachers" onClick="javascript:window.location='edi_admin_teacher.asp?site=' + document.forms.Screens.site.value + '&school=' + document.forms.Screens.school.options(document.forms.Screens.school.selectedIndex).value;">
						<input type="button" value="Save" name="SubmitAction" title="SAVE TEACHER" onClick="javascript:update_TeacherCheck(this.value);">
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
						<font class="boldtextblack">Teacher Code :&nbsp;&nbsp;</font>
					</td>
					<td>
						<%								
						' if there are no teachers then automatically in add mode
						' also add mode if user chooses to add
						' they choose the active school and then enter the teacher code
						
						if Request.Form("Action") = "Add" or intTeachers = 0 then 
							' site
							Response.Write "<select name=""site"" onChange=""javascript:window.location='edi_admin_teacher.asp?site=' + this.value;"">"
								
							' show all sites with schools
							for intRow = 0 to ubound(aSites,2)
								strSite = right("000" & aSites(0,intRow),3)
								if intRow = 0 OR  strSite <> strLast then 									
									Response.Write "<option value = """ & strSite & """"
									
									' if code is selected show it
									if intSite = strSite then 
										Response.write " selected"
									end if 
						
									Response.Write ">" & strSite & "</option>"
								end if 
							next
							Response.Write "</select>"

							' school
							Response.Write "<select name=""school"">"
								
							' show all schools at this site
							for intRow = 0 to ubound(aSchools,2)
								strSchool = right("000" & aSchools(0,intRow),3)
												
								Response.Write "<option value = """ & strSchool & """"
									
								' if code is selected show it
								if intSchool = strSchool then 
									Response.write " selected"
								end if 
						
								Response.Write ">" & strSchool & "</option>"
							next
							Response.Write "</select>"
	
							' teacher code input
							Response.Write "<input type=""text"" size=""15"" name=""code"" maxlength=""2"" title=""Enter the 2 digit teacher code"">"
						' not add mode
						else
							' site
							Response.Write "<select name=""site"" onChange=""javascript:window.location='edi_admin_teacher.asp?site=' + this.value;"">"
								
							' show all sites with schools
							for intRow = 0 to ubound(aSites,2)
								strSite = right("000" & aSites(0,intRow),3)
								if intRow = 0 OR  strSite <> strLast then 									
									Response.Write "<option value = """ & strSite & """"
									
									' if code is selected show it
									if intSite = strSite then 
										Response.write " selected"
									end if 
						
									Response.Write ">" & strSite & "</option>"
								end if 
							next
							Response.Write "</select>"

							' school
							Response.Write "<select name=""school"" onChange=""javascript:window.location='edi_admin_teacher.asp?site=' + document.forms.Screens.site.options(document.forms.Screens.site.selectedIndex).value + '&school=' + this.value;"">"
								
							' show all schools at this site
							for intRow = 0 to ubound(aSchools,2)
								strSchool = right("000" & aSchools(0,intRow),3)
												
								Response.Write "<option value = """ & strSchool & """"
									
								' if code is selected show it
								if intSchool = strSchool then 
									Response.write " selected"
								end if 
						
								Response.Write ">" & strSchool & "</option>"
							next
							Response.Write "</select>"
								
							' teacher						
							Response.Write "<select name=""code"" onChange=""javascript:window.location='edi_admin_teacher.asp?site=' + document.forms.Screens.site.options(document.forms.Screens.site.selectedIndex).value + '&school=' + document.forms.Screens.school.options(document.forms.Screens.school.selectedIndex).value + '&teacher=' + this.value;"">"
							for intRow = 0 to ubound(aData,2)						
								Response.Write "<option value = """ & right(aData(0,intRow),2) & """"
								' write the teacher
								if intTeacher = right(aData(0,intRow),2) then 
									Response.Write " selected"
								end if
								Response.Write ">" & right(aData(0,intRow),2) & "</option>"
							next
							Response.Write "</select>"
							Response.Write "&nbsp;<font class=""regtextgreen"">" & intTeachers & " Teacher"
							
							' if more than one teacher - plural
							if intTeachers > 1 then 
								Response.Write "s"
							end if 
							Response.Write " at this school.</font>"
						end if 
						%>
					</td>
				</tr>
				<tr valign="top">
					<td align="right">
						<font class="boldtextblack">Teacher Name :&nbsp;&nbsp;</font>
					</td>
					<td>
						<input type="text" size="80" name="name" value="<%=strName%>">
					</td>
				</tr>
				<tr valign="top">
					<td align="right">
						<font class="boldtextblack">Sex :&nbsp;&nbsp;</font>
					</td>
					<td>
						<select name="sex">
							<option value=""></option>
							<%
								Response.Write "<option value=""1"""
								if intSex = 1 then Response.Write " selected"
								Response.Write ">Male</option>"
								Response.Write "<option value=""2"""
								if intSex = 2 then Response.Write " selected"
								Response.Write ">Female</option>"
							%>
						</select>
				
						<font class="boldtextblack">Age :&nbsp;&nbsp;</font>
						<select name="age">
							<option value=""></option>
							<%
							Response.Write "<option value=""2"""
							if intAge = 2 then Response.Write " selected"
							Response.Write ">20-29</option>"
							Response.Write "<option value=""3"""
							if intAge = 3 then Response.Write " selected"
							Response.Write ">30-39</option>"
							Response.Write "<option value=""4"""
							if intAge = 4 then Response.Write " selected"
							Response.Write ">40-49</option>"
							Response.Write "<option value=""5"""
							if intAge = 5 then Response.Write " selected"
							Response.Write ">50-59</option>"
							Response.Write "<option value=""6"""
							if intAge = 6 then Response.Write " selected"
							Response.Write ">60 +</option>"
							
							%>
						</select>
						
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
			<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">
				<tr valign="top">
					<td align="left">
						<br />
						<font class="subheaderBlue">How long have you been :&nbsp;&nbsp;</font>
					</td>
				</tr>
			</table>
			
			<br />
			
			<!-- table of other educational pursuits-->

			<table border="1" cellpadding="0" cellspacing="0" width="550" align="center">
			<tr>
				<td>
					<table border="0" cellpadding="0" cellspacing="0" width="540" align="center">
						<tr>
							<td>
								<font class="boldtextblack">
									&nbsp;&nbsp;a) a teacher &nbsp;&nbsp;
								</font>
							</td>
							<td align="center">
								<select name="yr1" onChange="javascript:document.forms.Screens.intQ5a.value = Number(this.value * 12) + Number(document.forms.Screens.mth1.value);">
									<%
									for introw = 0 to 40 
										Response.Write "<option value=""" & introw & """"
										if introw = intyr1 then 
											Response.Write " selected"
										end if 
										Response.Write ">" & introw & "</option>"
									next 
									%>
								</select>
								<font class="boldtextblack">Yrs</font>
								<select name="mth1" onChange="javascript:document.forms.Screens.intQ5a.value = Number(document.forms.Screens.yr1.value * 12) + Number(this.value);">
									<%
									for introw = 0 to 11 
										Response.Write "<option value=""" & introw & """"
										if introw = intmth1 then 
											Response.Write " selected"
										end if 
										Response.Write ">" & introw & "</option>"
									next 
									%>
								</select>
								<font class="boldtextblack">Mths</font>
								<input type="hidden" name="intQ5a" size="5" value="<%=intQ5a%>">
							</td>
						</tr>
						<tr>
							<td>
								<font class="boldtextblack">
									&nbsp;&nbsp;b) a teacher at this school&nbsp;&nbsp;
								</font>
							</td>
							<td align="center">
								<select name="yr2" onChange="javascript:document.forms.Screens.intQ5b.value = Number(this.value * 12) + Number(document.forms.Screens.mth2.value);">
									<%
									for introw = 0 to 40 
										Response.Write "<option value=""" & introw & """"
										if introw = intyr2 then 
											Response.Write " selected"
										end if 
										Response.Write ">" & introw & "</option>"
									next 
									%>
								</select>
								<font class="boldtextblack">Yrs</font>
								<select name="mth2" onChange="javascript:document.forms.Screens.intQ5b.value = Number(document.forms.Screens.yr2.value * 12) + Number(this.value);">
									<%
									for introw = 0 to 11 
										Response.Write "<option value=""" & introw & """"
										if introw = intmth2 then 
											Response.Write " selected"
										end if 
										Response.Write ">" & introw & "</option>"
									next 
									%>
								</select>
								<font class="boldtextblack">Mths</font>
								<input type="hidden" name="intQ5b" size="5" value="<%=intQ5b%>">
							</td>
						</tr>
						<tr>
							<td>
								<font class="boldtextblack">
									&nbsp;&nbsp;c) a teacher of this grade&nbsp;&nbsp;
								</font>
							</td>
							<td align="center">
								<select name="yr3" onChange="javascript:document.forms.Screens.intQ5c.value = Number(this.value * 12) + Number(document.forms.Screens.mth3.value);">
									<%
									for introw = 0 to 40 
										Response.Write "<option value=""" & introw & """"
										if introw = intyr3 then 
											Response.Write " selected"
										end if 
										Response.Write ">" & introw & "</option>"
									next 
									%>
								</select>
								<font class="boldtextblack">Yrs</font>
								<select name="mth3" onChange="javascript:document.forms.Screens.intQ5c.value = Number(document.forms.Screens.yr3.value * 12) + Number(this.value);">
									<%
									for introw = 0 to 11 
										Response.Write "<option value=""" & introw & """"
										if introw = intmth3 then 
											Response.Write " selected"
										end if 
										Response.Write ">" & introw & "</option>"
									next 
									%>
								</select>
								<font class="boldtextblack">Mths</font>
								<input type="hidden" name="intQ5c" size="5" value="<%=intQ5c%>">
							</td>
						</tr>
						<%
						' display the length of time at each class
						if intClasses > 0 then 
							intCount = 4
							Response.Write "<tr><td colspan=""2""><font class=""boldtextblack"">&nbsp;&nbsp;d) a teacher of this class</td></tr>"
							do while not rstData.eof
								Response.Write "<tr><td align=""right""><font class=""boldtextblack"">" & right("000" & rstData("intClassID"),9) & "</font></td>"
								Response.Write "<td align=""center""><select name=""yr" & intcount & """ onChange=""javascript:document.forms.Screens.intQ5d" & intCount  & ".value = '" & right("000" & rstData("intClassID"),9) & "' + (Number(this.value * 12) + Number(document.forms.Screens.mth" & intCount & ".value));"">"
								for introw = 0 to 40 
									Response.Write "<option value=""" & introw & """"
									if introw = rstData("years") then 
										Response.Write " selected"
									end if 
									Response.Write ">" & introw & "</option>"
								next 
								Response.Write "</select>"
									
								Response.Write "<font class=""boldtextblack"">&nbsp;Yrs&nbsp;</font>"
								Response.Write "<select name=""mth" & intCount & """ onChange=""javascript:document.forms.Screens.intQ5d" & intCount & ".value = '" & right("000" & rstData("intClassID"),9) & "' + (Number(document.forms.Screens.yr" & intCount & ".value * 12) + Number(this.value));"">"
								for introw = 0 to 11 
									Response.Write "<option value=""" & introw & """"
									if introw = rstData("months") then 
										Response.Write " selected"
									end if 
									Response.Write ">" & introw & "</option>"
								next 
								Response.Write "</select>"
								Response.Write "<font class=""boldtextblack"">&nbsp;Mths&nbsp;</font>"
								Response.Write "<input type=""hidden"" name=""intQ5d" & intCount & """ size=""5"" value=""" & right("000" & rstData("intClassID"),9) & (rstData("years") * 12 + rstData("months")) & """>"
								
								Response.Write "</td></tr>"
								intCount = intCount + 1 
								rstData.MoveNext 
							loop 
							' passes the number of classes for updating
							Response.Write "<input type=""hidden"" name=""intClasses"" size=""5"" value=""" & intClasses & """>"
							rstData.movefirst	
						end if 
						%>
					</table>
				</td>
			</tr>
			</table>
			
			<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">
				<tr valign="top">
					<td align="left">
						<br />
						<font class="subheaderBlue">Completed levels of education(Check one or more if applicable) :&nbsp;&nbsp;</font>
					</td>
				</tr>
			</table>
			
			<br />
			
			<!-- table of other educational pursuits-->

			<table border="1" cellpadding="0" cellspacing="0" width="550" align="center">
			<tr>
				<td>
					<table border="0" cellpadding="0" cellspacing="0" width="540" align="center">
						<tr>
							<td>
								<font class="boldtextblack">
									&nbsp;&nbsp;a) some coursework towards a Bachelor's degree &nbsp;&nbsp;
								</font>
							</td>
							<td align="center">
								<font class="boldtextblack">
									<input type="radio" name="intQ6a" value="1" <%if intq6a = 1 then Response.Write "checked"%>>Yes &nbsp;&nbsp;<input type="radio" name="intQ6a" value="2" <%if intq6a = 2 then Response.Write "checked"%>>No
								</font>
							</td>
						</tr>
						<tr>
							<td>
								<font class="boldtextblack">
									&nbsp;&nbsp;b) a teaching certificate, diploma, or license &nbsp;&nbsp;
								</font>
							</td>
							<td align="center">
								<font class="boldtextblack">
									<input type="radio" name="intQ6b" value="1" <%if intq6b = 1 then Response.Write "checked"%>>Yes &nbsp;&nbsp;<input type="radio" name="intQ6b" value="2" <%if intq6b = 2 then Response.Write "checked"%>>No
								</font>
							</td>
						</tr>
						<tr>
							<td>
								<font class="boldtextblack">
									&nbsp;&nbsp;c) a Bachelor's degree &nbsp;&nbsp;
								</font>
							</td>
							<td align="center">
								<font class="boldtextblack">
									<input type="radio" name="intQ6c" value="1" <%if intq6c = 1 then Response.Write "checked"%>>Yes &nbsp;&nbsp;<input type="radio" name="intQ6c" value="2" <%if intq6c = 2 then Response.Write "checked"%>>No
								</font>
							</td>
						</tr>
						<tr>
							<td>
								<font class="boldtextblack">
									&nbsp;&nbsp;d) a Bachelor of Education degree &nbsp;&nbsp;
								</font>
							</td>
							<td align="center">
								<font class="boldtextblack">
									<input type="radio" name="intQ6d" value="1" <%if intq6d = 1 then Response.Write "checked"%>>Yes &nbsp;&nbsp;<input type="radio" name="intQ6d" value="2" <%if intq6d = 2 then Response.Write "checked"%>>No
								</font>
							</td>
						</tr>
						<tr>
							<td>
								<font class="boldtextblack">
									&nbsp;&nbsp;e) some post-baccalaureate coursework &nbsp;&nbsp;
								</font>
							</td>
							<td align="center">
								<font class="boldtextblack">
									<input type="radio" name="intQ6e" value="1" <%if intq6e = 1 then Response.Write "checked"%>>Yes &nbsp;&nbsp;<input type="radio"  name="intQ6e" value="2" <%if intq6e = 2 then Response.Write "checked"%>>No
								</font>
							</td>
						</tr>
						<tr>
							<td>
								<font class="boldtextblack">
									&nbsp;&nbsp;f) a post-baccalaureate diploma or certificate &nbsp;&nbsp;
								</font>
							</td>
							<td align="center">
								<font class="boldtextblack">
									<input type="radio" name="intQ6f" value="1" <%if intq6f = 1 then Response.Write "checked"%>>Yes &nbsp;&nbsp;<input type="radio"  name="intQ6f" value="2" <%if intq6f = 2 then Response.Write "checked"%>>No
								</font>
							</td>
						</tr>
						<tr>
							<td>
								<font class="boldtextblack">
									&nbsp;&nbsp;g) some coursework towards a Master's degree &nbsp;&nbsp;
								</font>
							</td>
							<td align="center">
								<font class="boldtextblack">
									<input type="radio" name="intQ6g" value="1" <%if intq6g = 1 then Response.Write "checked"%>>Yes &nbsp;&nbsp;<input type="radio"  name="intQ6g" value="2" <%if intq6g = 2 then Response.Write "checked"%>>No
								</font>
							</td>
						</tr>
						<tr>
							<td>
								<font class="boldtextblack">
									&nbsp;&nbsp;h) a Master's degree &nbsp;&nbsp;
								</font>
							</td>
							<td align="center">
								<font class="boldtextblack">
									<input type="radio" name="intQ6h" value="1" <%if intq6h = 1 then Response.Write "checked"%>>Yes &nbsp;&nbsp;<input type="radio"  name="intQ6h" value="2" <%if intq6h = 2 then Response.Write "checked"%>>No
								</font>
							</td>
						</tr>
						<tr>
							<td>
								<font class="boldtextblack">
									&nbsp;&nbsp;i) some coursework towards a Doctorate &nbsp;&nbsp;
								</font>
							</td>
							<td align="center">
								<font class="boldtextblack">
									<input type="radio" name="intQ6i" value="1" <%if intq6i = 1 then Response.Write "checked"%>>Yes &nbsp;&nbsp;<input type="radio"  name="intQ6i" value="2" <%if intq6i = 2 then Response.Write "checked"%>>No
								</font>
							</td>
						</tr>
						<tr>
							<td>
								<font class="boldtextblack">
									&nbsp;&nbsp;j) a Doctorate &nbsp;&nbsp;
								</font>
							</td>
							<td align="center">
								<font class="boldtextblack">
									<input type="radio" name="intQ6j" value="1" <%if intq6j = 1 then Response.Write "checked"%>>Yes &nbsp;&nbsp;<input type="radio"  name="intQ6j" value="2" <%if intq6j = 2 then Response.Write "checked"%>>No
								</font>
							</td>
						</tr>
						<tr>
							<td>
								<font class="boldtextblack">
									&nbsp;&nbsp;k) Other &nbsp;&nbsp;
								</font>
							</td>
							<td align="center">
								<font class="boldtextblack">
									<input type="radio" name="intQ6k" value="1" <%if intq6k = 1 then Response.Write "checked"%>>Yes &nbsp;&nbsp;<input type="radio"  name="intQ6k" value="2" <%if intq6k = 2 then Response.Write "checked"%>>No
								</font>
							</td>
						</tr>
					</table>
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
					<td align="center"><font class="boldtextblack">Class ID</font></td>
					<td align="center"><font class="boldtextblack">Language</font></td>
					<td align="center"><font class="boldtextblack">Size</font></td>
					<td align="center"><font class="boldtextblack">Completed</font></td>
				</tr>
				<%
								
				if intClasses = 0 then 
					Response.Write "<tr><td colspan=""7"">&nbsp;<font class=""regtextmaroon"">There are no classes at this school.</font></td></tr>"
				else
					do while not rstData.EOF 
						Response.Write "<tr><td><a href=""edi_admin_class.asp?site=" & left(right("000" & rstData("intClassID"),9),3) & "&school=" & mid(right("000" & rstData("intClassID"),9),4,3) & "&teacher=" & mid(right("000" & rstData("intClassID"),9),7,2) & "&class=" & right(rstData("intClassID"),1) & """ class=""reglinkBlue"">" & right("000" & rstData("intClassID"),9) & "</a></td>"
						Response.Write "<td><font class=""regtextblack"">" & rstData("strLanguage") & "</font></td>"
						Response.Write "<td align=""center""><font class=""regtextblack"">" & rstData("intStudents") & "</font></td>"
						Response.Write "<td align=""center""><font class=""regtextblack"">" & rstData("Completed") & "</font></td>"
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
	<%end if%>	
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
	' load the first school
	strName = ""
'	strAddress = ""
	strCity = ""
	intProvince = 1
	strPostal = ""
	strPhone = "" 
	strFax = ""
	strEmail = "" 
	strComments = ""
	intQ5a = 0
	intQ5b = 0
	intQ5c = 0
end sub

sub load_values(intTeacher)	
	if intTeacher = 0 then 
		introw = 0 
	else
		for introw = 0 to ubound(aData,2)
			if clng(intTeacher) = aData(0,introw) then 
				exit for
			end if
		next

		if intRow > ubound(aData,2) then 
			intRow = 0 
		end if 	
	end if 
	
	strName = aData(2,intRow)
	strEmail = aData(3,intRow)
	strPassword = aData(4,introw)
	strPhone = aData(5,introw)
	strFax = aData(6,introw) 
	intSex = aData(7,introw) 
	intAge = aData(8,introw)
	intQ5a = aData(9,introw)
	intyr1 = int(intQ5a / 12)
	intMth1 = intQ5a mod 12
	
	intQ5b = aData(10,introw)
	intyr2 = int(intQ5b / 12)
	intMth2 = intQ5b mod 12
	
	intQ5c = aData(11,introw)
	intyr3 = int(intQ5c / 12)
	intMth3 = intQ5c mod 12
	
	intQ6a = aData(13,introw) 
	intQ6b = aData(14,introw) 
	intQ6c = aData(15,introw) 
	intQ6d = aData(16,introw) 
	intQ6e = aData(17,introw) 
	intQ6f = aData(18,introw) 
	intQ6g = aData(19,introw) 
	intQ6h = aData(20,introw) 
	intQ6i = aData(21,introw) 
	intQ6j = aData(22,introw) 
	intQ6k = aData(23,introw)
end sub
%>
