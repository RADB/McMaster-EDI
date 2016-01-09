<!-- #include virtual="/shared/admin_security.asp" -->
<%
' public variables
'on error resume next
' totals
dim intSites, intSchools, intTeachers, intClasses, intStudents
' fields
dim strLocalID, intSex, dtmDob, strPostal, intLanguage, intDay,intMonth, intYear
' arrays
dim aData, aSites, aTeachers, aClasses, aLanguages, aClass

' initialize the variable
intLanguage = ""

' if the user has not logged in they will not be able to see the page
if blnSecurity then
	'call open_adodb(conn, "DATA")
	call open_adodb(conn, "MACEDI")
	
	set rstData = server.CreateObject("adodb.recordset")
	
	'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
	' Form Actions 
	' - December 17, 2002
	' - Andrew Renner
	'//////////////////////////////////////////////////////////////////////
	' delete record - February 8, 2003
	if Request.Form("Action") = "Delete" then 
		intEDIYear = Request.Form("frmEDIYear")
		intSite = Request.Form("site")
		intSchool = Request.Form("school")
		intTeacher = Request.Form("teacher")
		intClass = Request.Form("classtime") 
		intChild = Request.Form("code")
		
		' delete the unique Child  - will cascade down the other tables
		strSql = "DELETE FROM children WHERE strEdiID = '" & intEDIYear & intSite & intSchool & intTeacher & intClass & intChild & "'"
		'response.write strSql
		' execute the sql
		conn.execute strSql 	
		
		' build the error string
		if conn.errors.count > 0 then 
			strError = "<font class=""regtextred"">Error Number : " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/></font>"
		end if

	' Add a site - loads an empty form 
	'		     - set all values = ""

	elseif Request.Form("Action") = "Add" then 
		intEDIYear = request.form("frmEDIYear")
		intSite = Request.Form("site")
		intSchool = Request.Form("school")
		intTeacher = Request.Form("teacher")
		intClass = Request.Form("classtime")
		intStudent = Request.Form("code")
		
		call add_mode()
	' update records - December 16
	elseif Request.Form("Action") = "Update" then
		intEDIYear = Request.Form("frmEDIYear")
		intSite = Request.Form("site") 
		intSchool = Request.Form("school")
		intTeacher = Request.Form("teacher")
		intClass = Request.Form("classtime")
		intStudent = Request.Form("code")
		strLocal = Request.Form("localID")
		intSex = Request.Form("sex")
		strPostal = Request.Form("postal")
		dtmDOB = Request.Form("DOBday") & "/" & monthname(Request.Form("DOBmonth")) & "/" & Request.Form("DOByear")
				
		' build the SQL statement
		strSql = "UPDATE children " & _
  				 "SET intClassID = " & intSite & intSchool & intTeacher & intClass & ", intChild = " & intStudent & ", strLocalID = " & checkNull(strLocal) & ", intsex = " & intSex & ", dtmDOB = " & checkNull(dtmDOB) & ", strPostal = " & checknull(strPostal) & _
				 " WHERE strEDIID = " & checkNull(Request.Form("hiddenStudent"))
		
		'response.write strSQl
		' update the record
		conn.execute strSql
		
		
		' build the error string
		if conn.errors.count > 0 then 
			strError = "<font class=""regtextred"">Error Number : " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/></font>"
		end if
	' insert record - December 17	
	elseif Request.Form("Action") = "Save" then
		intEDIYear = Request.Form("frmEDIYear")
		intSite = Request.Form("site") 
		intSchool = Request.Form("school")
		intTeacher = Request.Form("teacher")
		intClass = Request.Form("classtime")
		intStudent = Request.Form("code")
		strLocal = Request.Form("localID")
		intSex = Request.Form("sex")
		strPostal = Request.Form("postal")
		dtmDOB = Request.Form("DOBday") & "/" & monthname(Request.Form("DOBmonth")) & "/" & Request.Form("DOByear")
				
		strSql = "INSERT INTO children (intClassID, intChild, strEDIID, strLocalID, intSex, dtmDOB, strPostal) " & _
					"VALUES(" & intSite & intSchool & intTeacher & intClass & "," & intStudent & "," & checknull(right("000" & intEDIYear & intSite & intSchool & intTeacher & intClass & intStudent,15)) & "," & checkNull(strLocal) & "," & intSex & "," & checknull(dtmDOb) & "," & checknull(strPostal) & ")"
		
		'Response.Write strSQL
		' insert the record
		conn.execute strSql
		
		if conn.errors.count > 0 then 
			strError = "<font class=""regtextred"">Error Number : " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/></font>"
		else
			aTables = array("demographics", "sectionA", "sectionB","sectionC","sectionD","sectionE") 
			for intRow = 0 to 5 
				strSql = "INSERT INTO " & aTables(introw) & " (strEDIID) VALUES(" &  checknull(right("000" & intEDIyear & intSite & intSchool & intTeacher & intClass & intStudent,15)) & ")"
				conn.execute strSql
			next 
		end if 
	else	
		if month(date()) > 6 then 
			intEDIYear = year(date()+1)
		else	
			intEDIYear = year(date())
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

	' get all classes
	rstData.Open "SELECT DISTINCT c.intClassID, t.intTeacherID, t.intSchoolID, s.intSiteID FROM classes c, teachers t, schools s WHERE c.intTeacherID = t.intteacherid AND  s.intSchoolID=t.intSchoolID ORDER BY c.intClassID", conn
	
	if not rstData.EOF then 
		' store info in array
		aClasses = rstData.GetRows 
					
		' get the total number of classes
		intClasses = ubound(aClasses,2) + 1							
	else
		intClasses = 0 
	end if 
						
	' close the recordset 
	rstData.Close 
	
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
	' Details:	A) Get total number of students  - intTotalStudents if their are classes then continue
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
	'				B) if no children but there are classes then enter add mode
	'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
	if Request.Form("Action") <> "Add" AND intClasses > 0 Then 
		' get the total number of classes		
		rstData.Open "SELECT COUNT(strEDIID) FROM children", conn
	
		' if more than 0 children/students
		if not rstData.EOF then 
			' get the total number of students
			intTotalStudents = rstData(0)
			'response.write intTotalStudents
			' close the recordset
			rstData.Close
		
			' check querystring for site
			if Request.QueryString("site").Count = 0 then 
				' this will have the current site if updated
				if intSite = "" then 
					' get the site from the classes array
					intSite = right("000" & aClasses(3,0),3)
				end if 
			else
				intSite = Request.QueryString("site")
			end if 
			
			strSql = "SELECT DISTINCT t.intSchoolID FROM classes c, teachers t, schools s WHERE c.intTeacherID = t.intTeacherID AND t.intSchoolID = s.intSchoolID AND s.intSiteID = " & intSite & " ORDER BY t.intSchoolID"
			'response.write strSql
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
				
	
				strSql = "SELECT DISTINCT c.intTeacherID FROM classes c, teachers t WHERE c.intTeacherID = t.intTeacherID AND t.intSchoolID = " & intSite & intSchool & " ORDER BY c.intTeacherID"
				'response.write strSql
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
						
					strSql = "SELECT DISTINCT c.intClassID, t.strName, t.strEmail, c.strComments " & _
								"FROM classes c LEFT JOIN Teachers t ON c.intTeacherID = t.intTeacherID " & _
								"WHERE c.intTeacherID = " & intSite & intSchool & intTeacher & _
								" ORDER BY c.intClassID"
									
					' get the school specific teachers
					'response.write strSql
					rstData.Open strSql, conn	
									
					if not rstData.EOF then
						' store info in array
						aClasses = rstData.GetRows 
														
						' get the number of classes the specified teacher has
						intClasses = ubound(aClasses,2) + 1							
						
						' get the class
						if Request.QueryString("class").Count = 0 then		
							' this will have the current class if updated
							if intClass = ""  then 
								intClass = right(aClasses(0,0),1)
							end if 
						else
							intClass = Request.QueryString("class")
						end if
						
						' close the recordset
						rstData.Close	
						
						strSql = "SELECT DISTINCT strEDIID,strLocalID,intSex,dtmDOB,strPostal FROM children WHERE intClassID = " & intSite & intSchool & intTeacher & intClass & " ORDER BY strEDIID"
						'response.write "SELECT DISTINCT strEDIID,strLocalID,intSex,dtmDOB,strPostal FROM children WHERE intClassID = " & intSite & intSchool & intTeacher & intClass & " ORDER BY strEDIID"
						
						' get the school specific teachers
						rstData.Open strSql, conn	

						if not rstData.EOF then
							' store info in array
							aData = rstData.GetRows 
							
							'get the edi year
							intEDIyear = left(aData(0,0),4)
							
							' get the number of students in this class
							intStudents = ubound(aData,2) + 1															
							' get the teacher
							if Request.QueryString("child").Count = 0 then		
								' this will have the current teacher if updated
								if intStudent = ""  then 
									intStudent = right(aData(0,0),2)
								end if 
							else
								intStudent = Request.QueryString("child")
							end if				
							
							' load the values 
							call load_values(intEDIyear & intSite & intSchool & intTeacher & intClass & intStudent)
					   else
							intStudents = 0 
							call add_mode
					   end if 
					else
						intClasses = 0
						intStudents = 0
					end if 
				else
					intTeachers = 0
					intClasses = 0
					intStudents = 0 
				end if 
			else
				intSchools = 0
				intTeachers = 0
				intClasses = 0
				intStudents = 0
			end if 
		' if 0 teachers
		else
			' 0 classes in the database
			intTotalStudents = 0 
			'call add_mode
		end if		 	
		
		' close the recordset
		rstData.Close 
	' add mode
	else
		' get total number of classes
		intTotalStudents = Request.Form("classes")
		
		' get all schools at the selected site
		'format function only works with sql 2012
		'rstData.Open "SELECT DISTINCT t.intSchoolID FROM classes c, teachers t, schools s WHERE c.intTeacherID = t.intTeacherID AND t.intSchoolID = s.intSchoolID AND left(format(c.intTeacherID,'00000000'),3) = " & Request.Form("site") & " ORDER BY t.intSchoolID", conn
        rstData.Open "SELECT DISTINCT t.intSchoolID FROM classes c, teachers t, schools s WHERE c.intTeacherID = t.intTeacherID AND t.intSchoolID = s.intSchoolID AND left(right('00000000' + c.intTeacherID,8),3) = " & Request.Form("site") & " ORDER BY t.intSchoolID", conn
		
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
		If Not Err.Number = 0 Then
        Response.Write Err.Description
        Error.Clear
    End If

		' get all classes for the selected teacher
		rstData.Open "SELECT DISTINCT intClassID FROM classes WHERE intTeacheriD = " & Request.Form("site") & Request.Form("school") & Request.Form("teacher") & " ORDER BY intClassID", conn
		'response.write "SELECT DISTINCT intClassID FROM classes WHERE intTeacheriD = " & Request.Form("site") & Request.Form("school") & Request.Form("teacher") & " ORDER BY intClassID"
		if not rstData.EOF then 
			' store info in array
			aClasses = rstData.GetRows 
		end if 
											
		' close the recordset 
		rstData.Close
	end if  
	If Not Err.Number = 0 Then
        Response.Write Err.Description
        Error.Clear
    End If

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
	<script language="javascript" type="text/javascript" src="js/window.js"></script>
</head>
<body>
	<!-- #include virtual="/shared/page_header.inc" -->
		
		<br />
		<a class="reglinkMaroon" href="edi_admin.asp">Home</a>&nbsp;<font class="regtextblack">></font>&nbsp;<font class="boldtextblack">Student Information</font>
		<table border="1" bordercolor="006600" cellpadding="0" cellspacing="0" width="760">
		<tr><td>
			<form name="Children" method="POST" action="edi_admin_student.asp"> 
			<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">
			<tr>
				<td align="right" width="500"><font class="headerBlue">Student Information (<%=intTotalStudents%>)</font></td>
				<td align="right">
					<input type="hidden" name="Action" value="">
					<input type="hidden" name="students" value="<%=intTotalStudents%>">
					<input type="hidden" name="hiddenstudent" value="<%=intEDIYear & intSite & intSchool & intTeacher & intClass & intStudent%>">
					<input type="hidden" name="Student" value="">
					<input type="hidden" name="XML" value="">
					<input type="hidden" name="rpt" value="">
					<input type="hidden" name="strLanguage" value="">
					<input type="hidden" name="frmEDIYear" value="">
					<input type="hidden" name="frmSite" value="">
					<input type="hidden" name="frmSchool" value="">
					<input type="hidden" name="frmTeacher" value="">
					<input type="hidden" name="frmClass" value="">					
					<input type="hidden" name="frmChild" value="">
				<%
				' checks to be sure that there are active sites
				' doesn't allow administration if no sites - works
				if intClasses = 0 then 					
					Response.Write "</td></tr><tr><td colspan=""2""><br /></td></tr><tr><td colspan=""2"" align=""left"">"
					Response.Write "<font class=""regtextred"">Please add a class before attempting to administer students.</font>" 		
					Response.Write "</td></tr><tr><td colspan=""2""><br /></td></tr></table>"
					Response.Write "</td></tr></table>"
					Response.Write "</form>"
				else
					' if there are no classes then automatically in add mode
					' also add mode if user chooses to add
					if Request.form("Action") <> "Add" AND intStudents > 0 then  
					%>
						<input type="button" value="Add" name="SubmitAction" title="ADD STUDENT" onClick="javascript:confirm_Student_Add(this.value,<%=intEDIyear%>);">
						<input type="button" value="Delete" name="SubmitAction" title="DELETE STUDENT" onClick="javascript:confirm_Student_Delete(this.value,<%=intEDIyear%>);">
						<%
						'if intclasses > 1 then 
						'	Response.Write "<input type=""button"" value=""Find"" name=""Find"" title=""FIND SITE"">"
						'end if 
						%>
						<input type="button" value="Update" name="SubmitAction" title="UPDATE STUDENT" onClick="javascript:update_Student_Check(this.value,<%=intEDIyear%>);">
					<%
					else
					%>
						<input type="button" value="Cancel" name="Cancel" title="Go to this classes students" onClick="javascript:window.location='edi_admin_student.asp?site=' + document.forms.Children.site.value + '&school=' + document.forms.Children.school.options(document.forms.Children.school.selectedIndex).value + '&teacher=' + document.forms.Children.teacher.options(document.forms.Children.teacher.selectedIndex).value + '&class=' + document.forms.Children.classtime.options(document.forms.Children.classtime.selectedIndex).value;">
						<!--<input type="button" value="Save" name="SubmitAction" title="SAVE STUDENT" onClick="javascript:alert();update_Student_Check(this.value);">-->
						<input type="button" value="Save" name="SubmitAction" title="SAVE STUDENT" onClick="javascript:update_Student_Check(this.value,'<%=intEDIYear%>');">
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
					<td align="right" width="100">
						<font class="boldtextblack">EDI ID :&nbsp;&nbsp;</font>
					</td>
					<td nowrap>
						<%		
						if Request.Form("Action") = "Add" or intStudents = 0 then 
							' site
							Response.Write "<select name=""site"" onChange=""javascript:window.location='edi_admin_student.asp?site=' + this.value;"">"
								
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
							Response.Write "<select name=""school"" onChange=""javascript:window.location='edi_admin_student.asp?site=' + document.forms.Children.site.options(document.forms.Children.site.selectedIndex).value + '&school=' + this.value;"">"
								
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
							Response.Write "<select name=""teacher"" onChange=""javascript:window.location='edi_admin_student.asp?site=' + document.forms.Children.site.options(document.forms.Children.site.selectedIndex).value + '&school=' + document.forms.Children.school.options(document.forms.Children.school.selectedIndex).value + '&teacher=' + this.value;"">"
							
							' show all teachers at this school
							for intRow = 0 to ubound(aTeachers,2)
								strTeacher = right("00" & aTeachers(0,intRow),2)
												
								Response.Write "<option value = """ & strTeacher & """"
									
								' if code is selected show it
								if intTeacher = strTeacher then 
									Response.write " selected"
								end if 
						
								Response.Write ">" & strTeacher & "</option>"
							next
							Response.Write "</select>"
							
					
							' class						
							Response.Write "<select name=""classtime"">"
	
							' show all classes for this teacher
							for intRow = 0 to ubound(aClasses,2)
								strClass = right("0" & aClasses(0,intRow),1)
												
								Response.Write "<option value = """ & strClass & """"
									
								' if code is selected show it
								if intClass = strClass then 
									Response.write " selected"
								end if 
						
								Response.Write ">" & strClass & "</option>"
							next
							Response.Write "</select>"
							
							' child code input
							Response.Write "<input type=""text"" size=""10"" name=""code"" maxlength=""2"" title=""Enter the 2 digit student code"">"
						' not add mode
						else
							' site
							Response.Write "<select name=""site"" onChange=""javascript:window.location='edi_admin_student.asp?site=' + this.value;"">"
								
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
							Response.Write "<select name=""school"" onChange=""javascript:window.location='edi_admin_student.asp?site=' + document.forms.Children.site.options(document.forms.Children.site.selectedIndex).value + '&school=' + this.value;"">"
								
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
							Response.Write "<select name=""teacher"" onChange=""javascript:window.location='edi_admin_student.asp?site=' + document.forms.Children.site.options(document.forms.Children.site.selectedIndex).value + '&school=' + document.forms.Children.school.options(document.forms.Children.school.selectedIndex).value + '&teacher=' + this.value;"">"
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
							Response.Write "<select name=""classtime"" onChange=""javascript:window.location='edi_admin_student.asp?site=' + document.forms.Children.site.options(document.forms.Children.site.selectedIndex).value + '&school=' + document.forms.Children.school.options(document.forms.Children.school.selectedIndex).value + '&teacher=' + document.forms.Children.teacher.options(document.forms.Children.teacher.selectedIndex).value + '&class=' + this.value;"">"
							for intRow = 0 to ubound(aClasses,2)						
								Response.Write "<option value = """ & right(aClasses(0,intRow),1) & """"
								' write the class
								if intClass = right(aClasses(0,intRow),1) then
									Response.Write " selected"
								end if
								Response.Write ">" & right(aClasses(0,intRow),1) & "</option>"
							next
							Response.Write "</select>"
				
							' children/students						
							Response.Write "<select name=""code"" onChange=""javascript:window.location='edi_admin_student.asp?site=' + document.forms.Children.site.options(document.forms.Children.site.selectedIndex).value + '&school=' + document.forms.Children.school.options(document.forms.Children.school.selectedIndex).value + '&teacher=' + document.forms.Children.teacher.options(document.forms.Children.teacher.selectedIndex).value + '&class=' + document.forms.Children.classtime.options(document.forms.Children.classtime.selectedIndex).value + '&child=' + this.value;"">"
							for intRow = 0 to ubound(aData,2)						
								Response.Write "<option value = """ & right(aData(0,intRow),2) & """"
								' write the class dsescription name
								if intStudent = right(aData(0,intRow),2) then 
									Response.Write " selected"
								end if
								Response.Write ">" & right(aData(0,intRow),2) & "</option>"
							next
							Response.Write "</select>"
							
							' display number of children in this class
							Response.Write "&nbsp;<font class=""regtextgreen"">This class has " & intStudents & " Student"
							
							' if more than one class - plural
							if intStudents > 1 then 
								Response.Write "s"
							end if 
							Response.Write "</font>"
						' end  non-add mode
						end if 
					%>
					</td>
					<td></td>
				</tr>
				
				<tr valign="top">
					<td align="right">
						<font class="boldtextblack">Local ID :&nbsp;&nbsp;</font>
					</td>
					<td>
						<input type="text" size="15" name="localID" value="<%=strLocalID%>" maxlength="20">
					</td>
					<td width="250" rowspan="2">
						<%
						if Request.Form("Action") <> "Add" AND intStudents > 0 then 
						%>						
						<a href="javascript:goEDI('edi_admin_questionnaire.asp','<%=intEDIYear %>','<%=intSite %>','<%=intSchool %>','<%=intTeacher %>','<%=intClass %>','<%=intStudent %>');" class="bigLinkBlue"><img border="0" src="images/download.gif"> EDI Questionnaire</a>
						<%
						end if
						%>
					</td>
				</tr>
	
				<tr valign="top">
					<td align="right">
						<font class="boldtextblack">Sex :&nbsp;&nbsp;</font>
					</td>
					<td>
						<select name="sex">
							<option value="-1"></option>
							<%
								Response.Write "<option value=""1"""
								if intSex = 1 then Response.Write " selected"
								Response.Write ">Male</option>"
								Response.Write "<option value=""2"""
								if intSex = 2 then Response.Write " selected"
								Response.Write ">Female</option>"
							%>
						</select>
					</td>
				</tr>
	
				<tr valign="top">
					<td align="right">
						<font class="boldtextblack">DOB :&nbsp;&nbsp;</font>
					</td>
					<td>
						<select name="DOBday">
							<option value="-1"></option>
						<%
						for introw = 1 to 31
							Response.Write "<option value = """ & intRow & """"
							if intDay = intRow then 
								Response.write " selected"
							end if 
							' write the day
							Response.Write ">" & intRow & "</option>"
						next
						%>
						</select>
						<select name="DOBmonth">
							<option value="-1"></option>
						<%
						for introw = 1 to 12
							Response.Write "<option value = """ & intRow & """"
							if intMonth = intRow then 
								Response.write " selected"
							end if 
							' write the day
							Response.Write ">" & monthname(intRow,false) & "</option>"
						next
						%>
						</select>
						<select name="DOByear">
							<option value="-1"></option>
						<%
						' include all years - 4
						'for introw = 1 to year(date)-1998
						for introw = 1 to 10
							Response.Write "<option value = """ & intRow + year(date)-14 & """"
							if intYear = intRow + year(date)-14 then 
								Response.write " selected"
							end if 
							' write the day
							Response.Write ">" & intRow + year(date) - 14 & "</option>"
						next
						%>
						</select>
					</td>
					<td rowspan="2">
						<%
						if Request.Form("Action") <> "Add" AND intStudents > 0 then 
						%>
						<a href="javascript:goAdminEDIReport('<%=intEDIYear & intSite & intSchool & intTeacher & intClass & intStudent %>');" class="bigLinkBlue"><img border="0" src="images/details.gif"> View Student Summary</a>
						<%
						end if 
						%>
					</td>
				</tr>
		
				<tr valign="top">
					<td align="right">
						<font class="boldtextblack">Postal Code :&nbsp;&nbsp;</font>
					</td>
					<td>
						<input type="text" size="10" name="postal" value="<%=strPostal%>" maxlength="7" title="Enter postal code without dashes or spaces"> 
					</td>
				</tr>
				</table>
			</form>
			</td>
		</tr>			
		</table>
	<%end if%>	
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

sub add_mode()
	intMonth = 0
	intyear = 0
	intDay = 0
end sub

'strEDIID,strLocalID,intSex,dtmDOB,strPostal,intLanguage
sub load_values(strEDIID)	
	if intStudent = 0 then 
		introw = 0 
	else
	'strPostal = ubound(aData,2)		
		for introw = 0 to ubound(aData,2)
			if cstr(strEDIID) = cstr(aData(0,introw)) then 
				exit for
			end if
		next
		
		if intRow > ubound(aData,2) then 
			intRow = 0 
		end if 	
	end if 
	
	strLocalID = aData(1,introw)
	intSex = aData(2,intRow)
	dtmDob = aData(3,intRow)
	intmonth = month(dtmDob)
	intday = day(dtmDob)
	intyear = year(dtmDob)
	strPostal = aData(4,introw)
	'intLanguage = aData(5,intRow)
end sub
%>
