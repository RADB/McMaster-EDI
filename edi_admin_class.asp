<!-- #include virtual="/shared/admin_security.asp" -->
<%
' public variables

' totals
dim intSites, intSchools, intTeachers, intClasses, intLanguage, intEDIYear
' fields
dim strName, strEmail, strComments
' arrays
dim aData, aSites, aTeachers, aClasses, aLanguages, aClass, rstYear
on error resume next

' initialize the variable
intLanguage = ""

' if the user has not logged in they will not be able to see the page
if blnSecurity then
	'call open_adodb(conn, "DATA")
	'call open_adodb(conn, "EDI")
	call open_adodb(conn, "MACEDI")

	set rstData = server.CreateObject("adodb.recordset")
	set rstYear = server.CreateObject("adodb.recordset")

	'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
	' get the languages for the drop down box
	'//////////////////////////////////////////////////////////////////////	
	set rstLanguages = server.CreateObject("adodb.recordset")
	
	' open all languages
	rstLanguages.Open "SELECT LID, english FROM [LU Languages] ORDER BY english", conn
	
	' store all languages in array
	aLanguages = rstLanguages.GetRows 
	
	' close and kill the langauges recordset
	call close_adodb(rstLanguages)
	
	
	'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
	' get the class times for the drop down box
	'//////////////////////////////////////////////////////////////////////	
	set rstClass = server.CreateObject("adodb.recordset")
	
	' open all languages
	rstClass.Open "SELECT intClassID, English FROM [LU Classes]", conn
	
	' store all languages in array
	aClass = rstClass.GetRows 
	
	' close and kill the langauges recordset
	call close_adodb(rstClass)
	
	'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
	' Form Actions 
	' - December 17, 2002
	' - Andrew Renner
	'//////////////////////////////////////////////////////////////////////
	' delete record - December 18
	if Request.Form("Action") = "Delete" then 
		intSite = Request.Form("site")
		intSchool = Request.Form("school")
		intTeacher = Request.Form("teacher")
		intClass = Request.Form("code") 
		
		' delete the unique CLASS 
		strSql = "DELETE FROM classes WHERE intCLASSID = " & intSite & intSchool & intTeacher & intClass
		
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
		intTeacher = Request.Form("teacher")
		call add_mode()
	' update records - December 16
	elseif Request.Form("Action") = "Update" then
		intEDIYear = Request.Form("frmEDIYear")
		intSite = Request.Form("site")
		intSchool = Request.Form("school")
		intTeacher = Request.Form("teacher")
		intClass = Request.Form("classtime") 
		intLanguage = Request.Form("language")
				
		' build the SQL statement
		strSql = "UPDATE classes " & _
  				 "SET intClassID = " & intSite & intSchool & intTeacher & intClass & ", intLanguage = " & intLanguage & ", strComments = " & checkNull(Request.Form("comments")) & _
				 " WHERE intCLASSID = " & Request.Form("hiddenclass")
		
		' update the record
		conn.execute strSql
		
		' build the error string
		if conn.errors.count > 0 then 
			strError = "<font class=""regtextred"">Error Number : " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/></font>"
		else
			strSql = "UPDATE children SET strEDIID = '" & intEDIYear & intSite & intSchool & intTeacher & intClass & "' + right(strEDIID,2) WHERE intCLASSID = " & intSite & intSchool & intTeacher & intClass
			'response.write strSQL
			' update the ediID
			conn.execute strSql
		
			' build the error string
			if conn.errors.count > 0 then 
				strError = strError & "<font class=""regtextred"">Error Number : " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/></font>"
			end if
		end if
	' insert record - December 17	
	elseif Request.Form("Action") = "Save" then
		intSite = Request.Form("site") 
		intSchool = Request.Form("school")
		intTeacher = Request.Form("teacher")
		intClass = Request.Form("code")
		intLanguage = Request.Form("language")
		
		strSQL = "INSERT INTO classes (intClassID,intTeacherID, intLanguage,strComments) VALUES" & _
				 "(" & intSite & intSchool & intTeacher & intClass & "," & intSite & intSchool & intTeacher & "," & intLanguage & "," & checkNull(Request.Form("comments")) & ")"
		
		'Response.Write strSQL  & Request.Form("strlanguage")
		' insert the record
		conn.execute strSql
		
		if conn.errors.count > 0 then 
			strError = "<font class=""regtextred"">Error Number : " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/></font>"
		end if 
	end if
	'\\\\\\\\\\\\\\\\\\\
	' END FORM ACTIONS '
	'///////////////////
	
	
	
	'////////////////////////////////////////////
	' 
	' Get all teachers and number of teachers 
	' Get all classes and number of classes 
	'  	
	'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
	
	' get all Teachers
	rstData.Open "SELECT DISTINCT t.intTeacherID, t.intSchoolID, s.intSiteID FROM teachers t, schools s WHERE s.intSchoolID=t.intSchoolID ORDER BY intTeacherID", conn
	
	if not rstData.EOF then 
		' store info in array
		aTeachers = rstData.GetRows 
					
		' get the total number of teachers
		intTeachers = ubound(aTeachers,2) + 1							
	else
		intTeachers = 0 
	end if 
						
	' close the recordset 
	rstData.Close 
	
	' get all schools - used for school list 
	rstData.Open "SELECT DISTINCT intSchoolID FROM [teachers] ORDER BY intSchoolID", conn
	
	if not rstData.EOF then 
		' store info in array
		aSchools = rstData.GetRows 
											
		' get the total number of teachers with classes
		intSchools = ubound(aSchools,2) + 1							
	else
		intSchools = 0 
	end if 

	' close the recordset 
	rstData.Close
	
	' get all the sites
	rstData.open "SELECT DISTINCT s.intSiteID FROM schools s INNER JOIN teachers t ON s.intSchoolID = t.intSchoolID", conn
		
	if not rstData.EOF then 
		' store info in array
		aSites = rstData.GetRows 			
		
		'get total number of sites with teachers
		intSites = ubound(aSites,2)
	else
		intSites = 0
	end if

	' close the recordset 
	rstData.Close  
	
	'///////////////////////////////////////	
	' Author - Andrew Renner
	' Description - FIRST TIME TO PAGE
	' Details:	A) Get total number of classes  - intTotalClasses if their are teachers then continue
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
	'						iii) get all classes the teacher has
	'							a) count the number of classes the teacher has
	'							b) if none go into add mode
	'					4) Check querystring for a class
	'						i) if class is there then save the class
	'						ii) if no class then get all classes at the school
	'						iii) load the first class values
	'				B) if no classes but there are teachers then enter add mode
	'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
	if Request.Form("Action") <> "Add" AND intTeachers > 0 Then 
		' get the total number of classes
		rstData.Open "SELECT COUNT(intClassID) FROM classes", conn
	
		' if more than 0 teachers
		if not rstData.EOF then 
			' get the total number of classes
			intTotalClasses = rstData(0)
	
			' close the recordset
			rstData.Close
		
			' check querystring for site
			if Request.QueryString("site").Count = 0 then 
				' this will have the current site if updated
				if intSite = "" then 
					' get the site from the teachers array
					intSite = right("000" & aTeachers(2,0),3)
				end if 
			else
				intSite = Request.QueryString("site")
			end if 
			
			strSql = "SELECT DISTINCT t.intSchoolID FROM [teachers] t, schools s WHERE t.intSchoolID = s.intSchoolID AND s.intSiteID = " & intSite & " ORDER BY t.intSchoolID"

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
				
	
				strSql = "SELECT DISTINCT intTeacherID FROM [teachers] WHERE intSchoolID = " & intSite & intSchool & " ORDER BY intTeacherID"
	
				' get the school specific teachers
				rstData.Open strSql, conn	
								
				if not rstData.EOF then
					' store info in array
					aTeachers = rstData.GetRows 
											
					' get the number of teachers at this school
					intTeachers = ubound(aTeachers,2) + 1							
							
					' get the teacher
					if Request.QueryString("teacher").Count = 0 then		
						' this will have the current teacher if updated
						if intTeacher = ""  then 
							intTeacher = right(aTeachers(0,0),2)
						end if 
					else
						intTeacher = Request.QueryString("teacher")
					end if
					
					' close the recordset
					rstData.Close			
					
					strSql = "SELECT DISTINCT c.intClassID, c.intLanguage, t.strName, t.strEmail, c.strComments " & _
								"FROM classes c LEFT JOIN Teachers t ON c.intTeacherID = t.intTeacherID " & _
								"WHERE c.intTeacherID = " & intSite & intSchool & intTeacher & _
								" ORDER BY c.intClassID"
					
					' get the school specific teachers
					rstData.Open strSql, conn	
					
					if not rstData.EOF then
						' store info in array
						aData = rstData.GetRows 
												
						' get the number of classes the specified teacher has
						intClasses = ubound(aData,2) + 1							
								
						' get the class
						if Request.QueryString("class").Count = 0 then		
							' this will have the current class if updated
							if intClass = ""  then 
								intClass = right(aData(0,0),1)
							end if 
						else
							intClass= Request.QueryString("class")
						end if
						
						' load the values
					   call load_values(intSite & intSchool & intTeacher & intClass)
					else
						intClasses = 0
						call add_mode
					end if 
				else
					intTeachers = 0
					intClasses = 0
				end if 
			else
				intSchools = 0
				intTeachers = 0
				intClasses = 0
			end if 
		' if 0 teachers
		else
			' 0 classes in the database
			intTotalClasses = 0 
			call add_mode
		end if		 	
		
		' close the recordset
		rstData.Close 
		
		strSql = "SELECT top 1 left(strEDIID,4) FROM GetClassChildren WHERE intClassID = " & intSite & intSchool & intTeacher & intClass
		' get the EDI year
		rstyear.Open strSql, conn
		if NOT rstYear.EOF then 
			intEDIYear = rstyear(0)
		else
			if month(date) >8 then
				intEDIYear = year(date)
			else 
				intEDIYear = year(date) - 1
			end if 
		end if 
		call close_adodb(rstYear)
	' add mode
	else
		' get total number of classes
		intTotalClasses = Request.Form("classes")
		
		' get all schools at the selected site
		rstData.Open "SELECT DISTINCT t.intSchoolID FROM [teachers] t, schools s WHERE t.intSchoolID = s.intSchoolID AND left(format(t.intSchoolID,'000000'),3) = " & Request.Form("site") & " ORDER BY t.intSchoolID", conn

		if not rstData.EOF then 
			' store info in array
			aSchools = rstData.GetRows 
		end if
							
		' close the recordset 
		rstData.Close  
		
		' get all teachers at the selected school
		rstData.Open "SELECT DISTINCT intTeacherID FROM [teachers] WHERE intSchoolID = " & Request.Form("site") & Request.Form("school") & " ORDER BY intTeacherID", conn
		
		if not rstData.EOF then 
			' store info in array
			aTeachers = rstData.GetRows 
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
		
		<br />
		<a class="reglinkMaroon" href="edi_admin.asp">Home</a>&nbsp;<font class="regtextblack">></font>&nbsp;<font class="boldtextblack">Class Information</font>
		<table border="1" bordercolor="006600" cellpadding="0" cellspacing="0" width="760">
		<tr><td>
			<form name="Screens" method="POST" action="edi_admin_class.asp"> 
			<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">
			<tr>
				<td align="right" width="430"><font class="headerBlue">Class Information (<%=intTotalClasses%>)</font></td>
				<td align="right">
					<input type="hidden" name="Action" value="">
					<input type="hidden" name="classes" value="<%=intTotalClasses%>">
					<input type="hidden" name="hiddenclass" value="<%=intSite & intSchool & intTeacher & intClass%>">
					<input type="hidden" name="strLanguage" value="">
					<input type="hidden" name="frmEDIYear" value="">
				<%
				' checks to be sure that there are active sites
				' doesn't allow administration if no sites - works
				if intTeachers = 0 then 					
					Response.Write "</td></tr><tr><td colspan=""2""><br /></td></tr><tr><td colspan=""2"" align=""left"">"
					Response.Write "<font class=""regtextred"">Please add a teacher before attempting to administer classes.</font>" 		
					Response.Write "</td></tr><tr><td colspan=""2""><br /></td></tr></table>"
					Response.Write "</td></tr></table>"
					Response.Write "</form>"
				else
					' if there are no classes then automatically in add mode
					' also add mode if user chooses to add
					if Request.form("Action") <> "Add" AND intClasses > 0 then  
					%>
						<input type="button" value="Add" name="SubmitAction" title="ADD CLASS" onClick="javascript:confirm_Add(this.value);">
						<input type="button" value="Delete" name="SubmitAction" title="DELETE CLASS" onClick="javascript:confirm_Delete(this.value);">
						<%
						'if intclasses > 1 then 
						'	Response.Write "<input type=""button"" value=""Find"" name=""Find"" title=""FIND SITE"">"
						'end if 
						%>
						<input type="button" value="Update" name="SubmitAction" title="UPDATE CLASS" onClick="javascript:update_Class_Check(this.value,<%=intEDIyear%>);">
					<%
					else
					%>
						<input type="button" value="Cancel" name="Cancel" title="Go to this teachers classes" onClick="javascript:window.location='edi_admin_class.asp?site=' + document.forms.Screens.site.value + '&school=' + document.forms.Screens.school.options(document.forms.Screens.school.selectedIndex).value + '&teacher=' + document.forms.Screens.teacher.options(document.forms.Screens.teacher.selectedIndex).value;">
						<input type="button" value="Save" name="SubmitAction" title="SAVE CLASS" onClick="javascript:update_Class_Check(this.value,<%=intEDIyear%>);">
					<%
					end if 
					%>
					<input type="button" value="Exit" name="Exit" title="EXIT Screen" onClick="javascript:window.location='edi_admin.asp';">
					&nbsp;
				</td>
			</tr>
			<!-- show error if any -->
			<tr><td colspan="2"><%=strError%></td></tr>
			<!-- end error-->
			</table>
			<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">
				<tr valign="top">
					<td align="right">
						<font class="boldtextblack">Class Code :&nbsp;&nbsp;</font>
					</td>
					<td>
						<%		
						if Request.Form("Action") = "Add" or intClasses = 0 then 
							' site
							Response.Write "<select name=""site"" onChange=""javascript:window.location='edi_admin_class.asp?site=' + this.value;"">"
								
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
							Response.Write "<select name=""school"" onChange=""javascript:window.location='edi_admin_class.asp?site=' + document.forms.Screens.site.options(document.forms.Screens.site.selectedIndex).value + '&school=' + this.value;"">"
								
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
							Response.Write "<select name=""teacher"">"
							
							' show all teachers at this school
							for intRow = 0 to ubound(aTeachers,2)
								strTeacher = right("000" & aTeachers(0,intRow),2)
												
								Response.Write "<option value = """ & strTeacher & """"
									
								' if code is selected show it
								if intTeacher = strTeacher then 
									Response.write " selected"
								end if 
						
								Response.Write ">" & strTeacher & "</option>"
							next
							Response.Write "</select>"
	
							' show the class selection box
							Response.Write "<select name=""code"">"
							Response.Write "<option value = ""-1""></option>"
							for intRow = 0 to ubound(aClass,2)						
								Response.Write "<option value = """ & aClass(0,intRow) & """"
								' write the class dsescription name
								Response.Write ">" & aClass(1,intRow) & "</option>"
							next
							Response.Write "</select>"
						' not add mode
						else
							' site
							Response.Write "<select name=""site"" onChange=""javascript:window.location='edi_admin_class.asp?site=' + this.value;"">"
								
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
							Response.Write "<select name=""school"" onChange=""javascript:window.location='edi_admin_class.asp?site=' + document.forms.Screens.site.options(document.forms.Screens.site.selectedIndex).value + '&school=' + this.value;"">"
								
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
							Response.Write "<select name=""teacher"" onChange=""javascript:window.location='edi_admin_class.asp?site=' + document.forms.Screens.site.options(document.forms.Screens.site.selectedIndex).value + '&school=' + document.forms.Screens.school.options(document.forms.Screens.school.selectedIndex).value + '&teacher=' + this.value;"">"
							for intRow = 0 to ubound(aTeachers,2)						
								Response.Write "<option value = """ & right(aTeachers(0,intRow),2) & """"
								' write the teacher
								if intTeacher = right(aTeachers(0,intRow),2) then 
									Response.Write " selected"
								end if
								Response.Write ">" & right(aTeachers(0,intRow),2) & "</option>"
							next
							Response.Write "</select>"
							
							' classes						
							Response.Write "<select name=""code"" onChange=""javascript:window.location='edi_admin_class.asp?site=' + document.forms.Screens.site.options(document.forms.Screens.site.selectedIndex).value + '&school=' + document.forms.Screens.school.options(document.forms.Screens.school.selectedIndex).value + '&teacher=' + document.forms.Screens.teacher.options(document.forms.Screens.teacher.selectedIndex).value + '&class=' + this.value;"">"
							for intRow = 0 to ubound(aData,2)						
								Response.Write "<option value = """ & right(aData(0,intRow),1) & """"
								' write the class dsescription name
								if intClass = right(aData(0,intRow),1) then 
									Response.Write " selected"
								end if
								Response.Write ">" & right(aData(0,intRow),1) & "</option>"
							next
							Response.Write "</select>"
							
							' display number of classes at this school for the teacher
							Response.Write "&nbsp;<font class=""regtextgreen"">This teacher has " & intClasses & " Class"
							
							' if more than one class - plural
							if intClasses > 1 then 
								Response.Write "es"
							end if 
							Response.Write " at this school.</font>"
						' end add mode
						end if 
					%>
					</td>
				</tr>
					<% 
					if Request.Form("Action") <> "Add" AND intclasses > 0 then 
					%>
				<tr valign="top">
					<td align="right">
						<font class="boldtextblack">Teacher Name :&nbsp;&nbsp;</font>
					</td>
					<td>
						<input type="text" size="80" name="name" value="<%=strName%>" readonly>
					</td>
				</tr>
	
				<tr valign="top">
					<td align="right">
						<font class="boldtextblack">Class Time :&nbsp;&nbsp;</font>
					</td>
					<td>
						<%
						Response.Write "<select name=""classtime"">"
						for intRow = 0 to ubound(aClass,2)						
							Response.Write "<option value = """ & aClass(0,intRow) & """"
							' select the selected class time
							if intClass = right(aClass(0,intRow),1) then
								Response.Write " selected"
							end if 
							' write the class description name
							Response.Write ">" & aClass(1,intRow) & "</option>"
						next
						Response.Write "</select>" 
						%>
					</td>
				</tr>
					<% 
					end if 
					%>

				<tr valign="top">
					<td align="right">
						<font class="boldtextblack">Language :&nbsp;&nbsp;</font>
					</td>
					<td>
						<select name="language">
						<%
						'if intLanguage = "" then 
							' only show blank option when they are adding records
						'	Response.Write "<option value = """" selected></option>"			
						'end if 
						
						aClassLanguage = array("","English","French","Other")
						
						for intRow = 0 to ubound(aClassLanguage)						
							Response.Write "<option value = """ & intRow & """"
							
							' if that province is selected than show it
							if intLanguage = intRow then 
								Response.write " selected"
							end if 
							
							' write the Language
							Response.Write ">" & aClassLanguage(intRow) & "</option>"
						next
						
						%>
						</select>
					</td>
				</tr>
					<% 
					if Request.Form("Action") <> "Add" AND intclasses > 0 then 
					%>

				<tr valign="top">
					<td align="right">
						<font class="boldtextblack">Email :&nbsp;&nbsp;</font>
					</td>
					<td>
						<input type="text" size="25" name="email" value="<%=strEmail%>" readonly>  
					</td>
				</tr>
					<% 
					end if
					%>				
				<tr valign="top">
					<td align="right">
						<font class="boldtextblack">Comments :&nbsp;&nbsp;</font>
					</td>
					<td>
						<textarea rows="3" cols="70" name="comments"><%=strComments%></textarea>
					</td>
					</tr>
				</table>
			</form>
			</td>
			</tr>
			
			
			
				<%
				' only show classes if on a site not in add mode
				if Request.Form("Action") <> "Add" AND intclasses > 0 then 
				%>
		
			<tr>
				<td>
					<form name="Children" method="POST" action=""> 	
						<input type="hidden" name="frmSite" value="">
						<input type="hidden" name="frmSchool" value="">
						<input type="hidden" name="frmTeacher" value="">
						<input type="hidden" name="frmClass" value="">					
						<input type="hidden" name="frmChild" value="">
						<input type="hidden" name="frmEDIYear" value="">
						<br />
						<table border="1" cellpadding="0" cellspacing="0" width="750" align="center">
							<tr>
								<td align="center"><font class="boldtextblack">EDI ID</font></td>
								<td align="center"><font class="boldtextblack">Local ID</font></td>
								<!--<td align="center"><font class="boldtextblack">Language</font></td>-->
								<td align="center"><font class="boldtextblack">Gender</font></td>
								<td align="center"><font class="boldtextblack">Date of Birth</font></td>
								<td align="center"><font class="boldtextblack">Postal Code</font></td>
								<td align="center"><font class="boldtextblack">Completed</font></td>
								<td align="center"><font class="boldtextblack">EDI</font></td>
							</tr>
							<%
							' select all children in the CLASS 				
							strSql = "SELECT strEDIID, strLocalID, gender, dtmDOB, strPostal, chkCompleted, dtmDate FROM GetClassChildren WHERE intClassID = " & intSite & intSchool & intTeacher & intClass & " ORDER BY strEDIID"
							'intLanguage removed - feb 8
							'Response.Write strSQL
							
							' open list of classes and teachers at this CLASS
							rstData.Open strSql, conn		
							if rstData.EOF then 
								Response.Write "<tr><td colspan=""8"">&nbsp;<font class=""regtextmaroon"">There are no children in this class.</font></td></tr>"
							else
								intEDIYear = left(rstData("strEDIID"),4)
								do while not rstData.EOF 
									Response.Write "<tr><td><a href=""javascript:window.location='edi_admin_student.asp?EDIYear=" & left(rstData("strEDIID"),4) & "&site=" & mid(rstData("strEDIID"),5,3) & "&school=" & mid(rstData("strEDIID"),8,3) & "&teacher=" & mid(rstData("strEDIID"),11,2) & "&class=" & mid(rstData("strEDIID"),13,1) & "&child=" & right(rstData("strEDIID"),2) & "';"" class=""reglinkBlue"">" & right("000" & rstData("strEDIID"),15) & "</a></td>"
									Response.Write "<td><font class=""regtextblack"">" & rstData("strLocalID") & "</font></td>"
									'Response.Write "<td><font class=""regtextblack"">"
									'for intRow = 0 to ubound(aLanguages,2)															
										' show the language
									'	if rstData("intLanguage") = aLanguages(0,intRow) then 
									'		Response.write aLanguages(1,introw)
									'		exit for 
									'	end if 
									'next
									'Response.Write "</font></td>"
									Response.Write "<td><font class=""regtextblack"">" & rstData("gender") & "</font></td>"
									Response.Write "<td align=""center""><font class=""regtextblack"">" & day(rstData("dtmDOB")) & "-" & monthname(datepart("m",rstData("dtmDOB")),true) & "-" & year(rstData("dtmDOB")) & "</font></td>"
									
									' changed 2006-02-27
									' Andrew Renner
									if err.number = 94 then 'Invalid use of null - when dob not available
										Response.Write "<td align=""center""><font class=""regtextblack"">NA</font></td>"
										err.Clear 
									end if
									'Response.Write "<td><font class=""regtextblack"">" 
									
									' get the language
									'for introw = 0 to ubound(aLanguages,2)
									'	if rstData("intLanguage") = aLanguages(0,introw) then 
									'		Response.Write aLanguages(1,intRow)
									'		exit for
									'	end if
									'next
									'Response.Write  "</font></td>"
									Response.Write "<td align=""center""><font class=""regtextblack"">" & rstData("strPostal") & "</font></td>"
									Response.Write "<td><font class=""regtextblack""><input type=""checkbox"" name=""completed"" disabled "
									if rstData("chkCompleted") then 
										Response.Write "checked>&nbsp;" & day(rstData("dtmDate")) & "-" & monthname(datepart("m",rstData("dtmDate")),true) & "-" & year(rstData("dtmDate"))
									else
										Response.Write ">"
									end if 
									Response.Write "</font></td>"
									Response.Write "<td><a href=""javascript:goEDI('edi_admin_questionnaire.asp','" & left(rstData("strEDIID"),4) & "','" & mid(rstData("strEDIID"),5,3) & "','" & mid(rstData("strEDIID"),8,3) & "','" & mid(rstData("strEDIID"),11,2) & "','" & mid(rstData("strEDIID"),13,1) & "','" & right(rstData("strEDIID"),2) & "');"" class=""reglinkBlue"">EDI</a></td>"
									'Response.Write "<td><a href=""javascript:goFormStudentDelete(" & checkNull(right("000" & rstData("strEDIID"),11)) & ");"" class=""regLinkMaroon"">Delete</a></td>"
									Response.Write "</tr>"
									rstData.MoveNext 
								loop
							end if 
							%>
						</table>
					</form>
					<!--<br />-->
				</td>
			</tr>
			<%
			end if 
			%>					
		</table>
	<%end if%>	
	<!-- #include virtual="/shared/page_footer.inc" -->
</body>
</html>
<%
	' close and kill recordset and connection
	call close_adodb(rstData)
	call close_adodb(conn)
	'call close_adodb(conn)
' security
end if

' set form defaults
sub add_mode()
	' load the first site
	strComments = ""
	intLanguage = ""
end sub

sub load_values(intClass)	
	if intClass = 0 then 
		introw = 0 
	else
		for introw = 0 to ubound(aData,2)
			if clng(intClass) = aData(0,introw) then 
				exit for
			end if
		next

		if intRow > ubound(aData,2) then 
			intRow = 0 
		end if 	
	end if 

	intLanguage = aData(1,intRow)
	strName = aData(2,intRow)
	strEmail = aData(3,intRow)
	strComments = aData(4,introw)
end sub
%>
