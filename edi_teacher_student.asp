<!-- #include virtual="/shared/security.asp" -->
<%
' public variables
on error resume next
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
    call open_adodb(conn, "MACEDI")
	'call open_adodb(conn, "DATA")
	'call open_adodb(conn, "EDI")
	
	set rstData = server.CreateObject("adodb.recordset")
	
	'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
	' Form Actions 
	' - December 17, 2002
	' - Andrew Renner
	'//////////////////////////////////////////////////////////////////////
	if Request.Form("Action") = "Add" then 
		intTeacher = Request.Form("teacher")
		intClass = Request.Form("classtime")
		intStudent = Request.Form("code")
		
		call add_mode()
	' update records - December 16
	elseif Request.Form("Action") = "Update" then
		intTeacher = Request.Form("teacher")
		intClass = Request.Form("classtime")
		intStudent = Request.Form("code")
		strLocal = Request.Form("localID")
		intSex = Request.Form("sex")
		strPostal = Request.Form("postal")
		dtmDOB = Request.Form("DOBday") & "/" & monthname(Request.Form("DOBmonth")) & "/" & Request.Form("DOByear")

				
		' build the SQL statement
		strSql = "UPDATE children " & _
  				 "SET intClassID = " & intTeacher & intClass & ", intChild = " & intStudent & ", strLocalID = " & checkNull(strLocal) & ", intsex = " & intSex & ", dtmDOB = " & checkNull(dtmDOB) & ", strPostal = " & checknull(strPostal) & _
				 " WHERE strEDIID = " & checkNull(Request.Form("hiddenStudent"))
		
		' update the record
		conn.execute strSql
		
		
		' build the error string
		if conn.errors.count > 0 then 
			strError = "<font class=""regtextred"">Error Number : " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/></font>"
		end if
	' insert record - December 17	
	elseif Request.Form("Action") = "Save" then
		intTeacher = Request.Form("teacher")
		intClass = Request.Form("classtime")
		intStudent = Request.Form("code")
		strLocal = Request.Form("localID")
		intSex = Request.Form("sex")
		strPostal = Request.Form("postal")
		dtmDOB = Request.Form("DOBday") & "/" & monthname(Request.Form("DOBmonth")) & "/" & Request.Form("DOByear")
		
		strSql = "INSERT INTO children (intClassID, intChild, strEDIID, strLocalID, intSex, dtmDOB, strPostal) " & _
					"VALUES(" & intTeacher & intClass & "," & intStudent & "," & checknull(right("000" & intTeacher & intClass & intStudent,11)) & "," & checkNull(strLocal) & "," & intSex & "," & checknull(dtmDOb) & "," & checknull(strPostal) & ")"
		
		'Response.Write strSQL
		' insert the record
		conn.execute strSql
		
		if conn.errors.count > 0 then 
			strError = "<font class=""regtextred"">Error Number : " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/></font>"
		else
			aTables = array("demographics", "sectionA", "sectionB","sectionC","sectionD","sectionE") 
			for intRow = 0 to 5 
				strSql = "INSERT INTO " & aTables(introw) & " (strEDIID) VALUES(" &  checknull(right("000" & intTeacher & intClass & intStudent,11)) & ")"
				conn.execute strSql
			next 
			
			htmltext ="<html><head><title>Child Addition</title></head><body><center><img src=""http://www.e-edi.ca/images/e-edi.gif"" alt=""e-EDI"" name=""e-edi.gif""><br><br><font color=""black"">A new student has been added to class <b>" & intTeacher & intClass & "</b>.<br /><br /><b>Student: </b>" & right("000" & intTeacher & intClass & intStudent,11) & "<br /><b>Local ID: </b>" & strLocal & "<br /><b>DOB: </b>" & dtmDOB & "</font></center></body></html>"

			set objmail = server.CreateObject("CDONTS.NewMail")
				objmail.From = "webmaster@e-edi.ca"
				objmail.To = "webmaster@e-edi.ca"
				objmail.Subject = "e-EDI Child Addition by Teacher " & intTeacher
				objmail.BodyFormat = 0
				objmail.MailFormat = 0
				objmail.Body = htmlText
				objmail.Send 
			set objmail = nothing
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

	' get all classes for this teacher
	rstData.Open "SELECT DISTINCT c.intClassID, t.intTeacherID, t.intSchoolID, s.intSiteID FROM classes c, teachers t, schools s WHERE c.intTeacherID = t.intteacherid AND  s.intSchoolID=t.intSchoolID AND t.strEmail ='" & session("id") & "' ORDER BY c.intClassID", conn
	
	if not rstData.EOF then 
		' store info in array
		aClasses = rstData.GetRows 
					
		' get the total number of classes that this teacher has
		intClasses = ubound(aClasses,2) + 1							
	else
		intClasses = 0 
	end if 
						
	' close the recordset 
	rstData.Close 

	' check if not add and this teacher has at least one class
	if Request.Form("Action") <> "Add" AND intClasses > 0 Then 
		' get the total number of classes
		rstData.Open "SELECT COUNT(strEDIID) FROM children", conn
	
		' if more than 0 children/students
		if not rstData.EOF then 
			' get the total number of students
			intTotalStudents = rstData(0)
	
			' close the recordset
			rstData.Close
		
			strSql = "SELECT DISTINCT c.intTeacherID FROM classes c, teachers t WHERE c.intTeacherID = t.intTeacherID AND t.strEmail ='" & session("id") & "' ORDER BY c.intTeacherID"
	
			' get the school specific teachers
			rstData.Open strSql, conn	
									
			if not rstData.EOF then
				' store info in array
				aTeachers = rstData.GetRows 
												
				' get the number of teachers at this school
				intTeachers = ubound(aTeachers,2) + 1							
									
				' get and verify the teacher id
				if intTeacher = "" AND Request.QueryString("teacher").Count = 0 then 
					intTeacher = right("000" & aTeachers(0,0),8)
				else
					if intTeacher = "" then 
						' check to see if the teacherid in the querystring is one of this teachers
						for introw = 0 to ubound(aTeachers,2)
							if Request.QueryString("teacher") = right("000" & aTeachers(0,introw),8) then 
								intTeacher = Request.QueryString("teacher")
								exit for 
							end if
						next
								
						if intTeacher = "" then 
							' user entered a value that is not valid for this user
							Response.Redirect "edi_teacher_student.asp"
						end if 
					end if 
				end if
	
				' close the recordset
				rstData.Close						
						
				strSql = "SELECT DISTINCT c.intClassID " & _
						"FROM classes c LEFT JOIN Teachers t ON c.intTeacherID = t.intTeacherID " & _
						"WHERE c.intTeacherID = " & intTeacher & _
						" ORDER BY c.intClassID"
	
				' get the school specific teachers
				rstData.Open strSql, conn	
						
				if not rstData.EOF then
					' store info in array
					aClasses = rstData.GetRows 
													
					' get the number of classes the specified teacher has
					intClasses = ubound(aClasses,2) + 1							
					
					' get the class
					if Request.QueryString("class").Count = 0 AND intClass = "" then		 
						intClass = right(aClasses(0,0),1)
					else
						if intClass = "" then 
							for introw = 0 to ubound(aClasses,2)
								if Request.QueryString("class") = right(aClasses(0,introw),1) then 
									intClass = Request.QueryString("class")
									exit for 
								end if
							next
						
							if intClass = "" then 
								' user entered a value that is not valid for this user
								Response.Redirect "edi_teacher_student.asp"
							end if 
						end if 
					end if 					
						
					' close the recordset
					rstData.Close						
						
					strSql = "SELECT DISTINCT strEDIID,strLocalID,intSex,dtmDOB,strPostal FROM children WHERE intClassID = " & intTeacher & intClass & " ORDER BY strEDIID"
					
					' get the school specific teachers
					rstData.Open strSql, conn	

					if not rstData.EOF then
						' store info in array
						aData = rstData.GetRows 
														
						' get the number of students in this class
						intStudents = ubound(aData,2) + 1															
						
						' get the child
						if Request.QueryString("child").Count = 0 AND intStudent = "" then		 
							intStudent = right("0" & aData(0,0),2)
						else
							if intStudent = "" then 
								for introw = 0 to ubound(aData,2)
									if Request.QueryString("child") = right("0" & aData(0,introw),2) then 
										intStudent = Request.QueryString("child")
										exit for 
									end if
								next
								
								if intStudent = "" then 
									' user entered a value that is not valid for this user
									Response.Redirect "edi_teacher_student.asp"
								end if 
							end if 
						end if 					
						
						' load the values 
						call load_values(intTeacher & intClass & intStudent)
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
		
		' get all teachers
		rstData.Open "SELECT DISTINCT intTeacherID FROM [teachers] WHERE strEmail ='" & session("id") & "' ORDER BY intTeacherID", conn
		
		if not rstData.EOF then 
			' store info in array
			aTeachers = rstData.GetRows 
		end if 
											
		' close the recordset 
		rstData.Close  
		
		' get all classes for the selected teacher
		rstData.Open "SELECT DISTINCT c.intClassID FROM classes c, teachers t WHERE c.intTeacherID = t.intTeacherID AND t.intTeacherID = " & intTeacher & " ORDER BY c.intClassID"
		if not rstData.EOF then 
			' store info in array
			aClasses = rstData.GetRows 
		end if 
											
		' close the recordset 
		rstData.Close
	end if  
%>
<html>
<head>
	<!-- Start CSS files-->
		<link rel="stylesheet" type="text/css" href="Styles/edi.css">
	<!-- End CSS files -->
	<script language="javascript" type="text/javascript" src="js/form.js"></script>
	<script language="javascript" type="text/javascript" src="js/window.js"></script>
</head>
<body>
	<!-- #include virtual="/shared/page_header.inc" -->
		
		<br />
		<a class="reglinkMaroon" href="edi_teacher.asp"><%=strHome%></a>&nbsp;<font class="regtextblack">></font>&nbsp;<font class="boldtextblack"><%=lblStudent%></font>
		<table border="1" bordercolor="006600" cellpadding="0" cellspacing="0" width="760">
		<tr><td>
			<form name="Children" method="POST" action="edi_teacher_student.asp"> 
			<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">
			<tr>
				<td align="right" width="500"><font class="headerBlue"><%=lblStudent%> (<%=intTeacher & intClass & intStudent%>)</font></td>
				<td align="right">
					<input type="hidden" name="Action" value="">
					<input type="hidden" name="students" value="<%=intTotalStudents%>">
					<input type="hidden" name="hiddenstudent" value="<%=intTeacher & intClass & intStudent%>">
					<input type="hidden" name="Student" value="">
					<input type="hidden" name="XML" value="">
					<input type="hidden" name="rpt" value="">
					<input type="hidden" name="strLanguage" value="">
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
					Response.Write "<font class=""regtextred"">This teacher has no classes to administer.</font>" 		
					Response.Write "</td></tr><tr><td colspan=""2""><br /></td></tr></table>"
					Response.Write "</td></tr></table>"
					Response.Write "</form>"
				else
					' if there are no classes then automatically in add mode
					' also add mode if user chooses to add
					if Request.form("Action") <> "Add" AND intStudents > 0 then  
					%>
						<input type="button" value="<%=lblAdd%>" name="Add"  onClick="javascript:confirm_Student_Add(this.name);">
						<!--<input type="button" value="Delete" name="SubmitAction" title="DELETE STUDENT" onClick="javascript:confirm_Delete(this.value);">-->
						<%
						'if intclasses > 1 then 
						'	Response.Write "<input type=""button"" value=""Find"" name=""Find"" title=""FIND SITE"">"
						'end if 
						%>
						<input type="button" value="Update" name="SubmitAction" title="UPDATE STUDENT" onClick="javascript:update_Student_Check(this.value);">
					<%
					else
					%>
						<input type="button" value="<%=lblCancel%>" name="Cancel" onClick="javascript:window.location='edi_teacher_student.asp?teacher=' + document.forms.Children.teacher.options(document.forms.Children.teacher.selectedIndex).value + '&class=' + document.forms.Children.classtime.options(document.forms.Children.classtime.selectedIndex).value;">
						<input type="button" value=<%=strSave%> name="Save"  onClick="javascript:update_Student_Check(this.name);">
					<%
					end if 
					%>
					<input type="button" value="<%=strExit%>" name="Exit"  onClick="javascript:window.location='edi_teacher.asp';">
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
						<font class="boldtextblack"><%=lblEDIID%> :&nbsp;&nbsp;</font>
					</td>
					<td nowrap>
						<%		
						if Request.Form("Action") = "Add" or intStudents = 0 then 
							' teacher						
							Response.Write "<select name=""teacher"" onChange=""javascript:window.location='edi_teacher_student.asp?teacher=' + this.value;"">"
							
							' show all teachers at this school
							for intRow = 0 to ubound(aTeachers,2)
								strTeacher = right("00" & aTeachers(0,intRow),8)
												
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
							Response.Write "<input type=""text"" size=""10"" name=""code"" maxlength=""2"" title=""Enter the 2 digit child code"">"
						' not add mode
						else
							' teacher						
							Response.Write "<select name=""teacher"" onChange=""javascript:window.location='edi_teacher_student.asp?teacher=' + this.value;"">"
							for intRow = 0 to ubound(aTeachers,2)						
								Response.Write "<option value = """ & right("00" & aTeachers(0,intRow),8) & """"
								' write the teacher
								if intTeacher = right("00" & aTeachers(0,intRow),8) then 
									Response.Write " selected"
								end if
								Response.Write ">" & right("00" & aTeachers(0,intRow),8) & "</option>"
							next
							Response.Write "</select>"
							
							' classes						
							Response.Write "<select name=""classtime"" onChange=""javascript:window.location='edi_teacher_student.asp?teacher=' + document.forms.Children.teacher.options(document.forms.Children.teacher.selectedIndex).value + '&class=' + this.value;"">"
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
							Response.Write "<select name=""code"" onChange=""javascript:window.location='edi_teacher_student.asp?teacher=' + document.forms.Children.teacher.options(document.forms.Children.teacher.selectedIndex).value + '&class=' + document.forms.Children.classtime.options(document.forms.Children.classtime.selectedIndex).value + '&child=' + this.value;"">"
							for intRow = 0 to ubound(aData,2)						
								Response.Write "<option value = """ & right(aData(0,intRow),2) & """"
								' write the class dsescription name
								if intStudent = right(aData(0,intRow),2) then 
									Response.Write " selected"
								end if
								Response.Write ">" & right(aData(0,intRow),2) & "</option>"
							next
							Response.Write "</select>"
							
							if session("Language") = "English" then 
								' display number of children in this class
								Response.Write "&nbsp;<font class=""regtextgreen"">This class has " & intStudents & " Student"
							
								' if more than one class - plural
								if intStudents > 1 then 
									Response.Write "s"
								end if 
								Response.Write "</font>"
							else
								' display number of children in this class
								Response.Write "&nbsp;<font class=""regtextgreen"">Cette classe a " & intStudents & " étudiant"
							
								' if more than one class - plural
								if intStudents > 1 then 
									Response.Write "es"
								end if 
								Response.Write "</font>"
							end if 
						' end  non-add mode	
						end if 
					%>
					</td>
					<td></td>
				</tr>
				
				<tr valign="top">
					<td align="right">
						<font class="boldtextblack"><%=lblLocal%> :&nbsp;&nbsp;</font>
					</td>
					<td>
						<input type="text" size="15" name="localID" value="<%=strLocalID%>" maxlength="10">
					</td>
					<td width="250" rowspan="2">
						<%
						if Request.Form("Action") <> "Add" AND intStudents > 0 then 
						%>
						<a href="javascript:goEDI('edi_teacher_questionnaire.asp','<%=left(intTeacher,3) %>','<%=mid(intTeacher,4,3) %>','<%=mid(intTeacher,7,2) %>','<%=intClass %>','<%=intStudent %>');" class="bigLinkBlue"><img border="0" src="images/download.gif"> <%=lblEDI%> Questionnaire</a>
						<%
						end if
						%>
					</td>
				</tr>
	
				<tr valign="top">
					<td align="right">
						<font class="boldtextblack"><%=lblSex%> :&nbsp;&nbsp;</font>
					</td>
					<td>
						<select name="sex">
							<option value="-1"></option>
							<%
								Response.Write "<option value=""1"""
								if intSex = 1 then Response.Write " selected"
								Response.Write ">" & lblMale & "</option>"
								Response.Write "<option value=""2"""
								if intSex = 2 then Response.Write " selected"
								Response.Write ">" & lblFemale & "</option>"
							%>
						</select>
					</td>
				</tr>
	
				<tr valign="top">
					<td align="right">
						<font class="boldtextblack"><%=lblDOB%> :&nbsp;&nbsp;</font>
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
							' write the month
							if session("language") = "English" then 
								Response.Write ">" & monthname(intRow,false) & "</option>"
							else
								Response.Write ">" & French_Month(intRow) & "</option>"
							end if 
						next
						%>
						</select>
						<select name="DOByear">
							<option value="-1"></option>
						<%
						for introw = 1 to year(date)-1998
							Response.Write "<option value = """ & intRow + 1994 & """"
							if intYear = intRow + 1994 then 
								Response.write " selected"
							end if 
							' write the day
							Response.Write ">" & intRow + 1994 & "</option>"
						next
						%>
						</select>
					</td>
					<td rowspan="2">
						<%
						if Request.Form("Action") <> "Add" AND intStudents > 0 then 
						%>
						<a href="javascript:goTeacherEDIReport('<%=intSite & intSchool & intTeacher & intClass & intStudent %>');" class="bigLinkBlue"><img border="0" src="images/details.gif"> <%=lblSummary%></a>
						<%
						end if 
						%>
					</td>
				</tr>
		
				<tr valign="top">
					<td align="right">
						<font class="boldtextblack"><%=lblPostal%> :&nbsp;&nbsp;</font>
					</td>
					<td>
						<input type="text" size="10" name="postal" value="<%=strPostal%>" maxlength="7"> 
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
	intLanguage = aData(5,intRow)
end sub
%>
