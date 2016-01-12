<!-- #include virtual="/shared/security.asp" -->
<%
' public variables

' totals
dim intSites, intSchools, intTeachers, intClasses, intLanguage
' fields
dim strName, strEmail, strComments
' arrays
dim aData, aSites, aTeachers, aClasses, aLanguages, aClass
on error resume next

' initialize the variable
intLanguage = ""

' if the user has not logged in they will not be able to see the page
if blnSecurity then
	'call open_adodb(conn, "DATA")
	call open_adodb(conn, "MACEDI")
	
	set rstData = server.CreateObject("adodb.recordset")

	'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
	' get the languages for the drop down box
	'//////////////////////////////////////////////////////////////////////	
	set rstLanguages = server.CreateObject("adodb.recordset")
	
	' open all languages
	rstLanguages.Open "SELECT LID, " & session("language") & " as strText FROM [LU Languages] ORDER BY english", conn
	
	' store all languages in array
	aLanguages = rstLanguages.GetRows 
	
	' close and kill the langauges recordset
	call close_adodb(rstLanguages)
	
	
	'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
	' get the class times for the drop down box
	'//////////////////////////////////////////////////////////////////////	
	set rstClass = server.CreateObject("adodb.recordset")
	
	' open all languages
	rstClass.Open "SELECT intClassID,  " & session("language") & " as strText FROM [LU Classes]", conn
	
	' store all languages in array
	aClass = rstClass.GetRows 
	
	' close and kill the langauges recordset
	call close_adodb(rstClass)
	
	'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
	' Form Actions 
	' - December 17, 2002
	' - Andrew Renner
	'//////////////////////////////////////////////////////////////////////
	if Request.Form("Action") = "Update" then
		intTeacher = Request.Form("teacher")
		intClass = Request.Form("code")
				
		' build the SQL statement
		strSql = "UPDATE classes " & _
  				 "SET strComments = " & checkNull(Request.Form("comments")) & _
				 " WHERE intCLASSID = " & Request.Form("hiddenclass")
		
		' update the record
		conn.execute strSql
		
		
		' build the error string
		if conn.errors.count > 0 then 
			strError = "<font class=""regtextred"">Error Number : " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/></font>"
		end if
	else 
		
		if Request.Form("Action") = "Add" then
			strSql = "INSERT INTO children (intClassID,intChild,strEDIID,strLocalID,createdBy) VALUES(" & mid(request.form("ediID"),5,9) & "," & right(request.form("ediID"),2) & ",'" & request.form("ediID") & "','" & request.form("LocalID") & "','" &session("id") & "')"	
			'response.write strSql
			conn.execute strSql
			
			if conn.errors.count > 0 then 
				strError = "<font class=""regtextred"">Error Number : " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/></font>"
			else
				aTables = array("demographics", "sectionA", "sectionB","sectionC","sectionD","sectionE") 
				for intRow = 0 to 5 
					strSql = "INSERT INTO " & aTables(introw) & " (strEDIID) VALUES('" &  request.form("ediID") & "')"
					conn.execute strSql
				next 
				
				htmltext ="<html><head><title>Child Addition</title></head><body><center><img src=""http://www.e-edi.ca/images/e-edi.gif"" alt=""e-EDI"" name=""e-edi.gif""><br><br><font color=""black"">A new student has been added to class <b>" & left(request.form("ediID"),9) & "</b>.<br /><br /><b>Student: </b>" & request.form("ediID") & "<br /><b>Local ID: </b>" & request.form("localid") & "<br /></font></center></body></html>"
				set objmail = server.CreateObject("CDONTS.NewMail")
					objmail.From = "webmaster@e-edi.ca"
					objmail.To = "webmaster@e-edi.ca"
					objmail.Subject = "e-EDI Child Addition by Teacher " & left(request.form("ediID"),8)
					objmail.BodyFormat = 0
					objmail.MailFormat = 0
					objmail.Body = htmlText
					'objmail.Send 
				set objmail = nothing            	            
			end if			
		elseif Request.form("frmAction") = "lock" then 
			' put the EDI ID together
			strEDIID = ""
			'for each item in Request.Form
			'	strEDIID = strEDIID & Request.form(item)
			'next
	
			' get individual variables		
			if Request.Form("frmNextChild") <> "" then 
				strChild = Request.Form("frmNextChild")
			else
				strChild = Request.Form("frmChild")
			end if
			strClass = Request.Form("frmClass")
			strTeacher = Request.Form("frmTeacher")
			strSchool = Request.Form("frmSchool")
			strSite = Request.Form("frmSite")
			strEDIYear = request.form("frmEDIYear")
			strEdiID = strEDIYear & strSite & strSchool & strTeacher & strClass & strChild 
            strSql = "UPDATE children SET chkCompleted = 1, dtmDate = '" & date & "' WHERE strEDIID = '" & strEDIID & "'"	            
			conn.execute strSql
		elseif Request.form("frmAction") = "consent" then
			strChild = Request.Form("frmChild")
			strClass = Request.Form("frmClass")
			strTeacher = Request.Form("frmTeacher")
			strSchool = Request.Form("frmSchool")
			strSite = Request.Form("frmSite")
			strEDIYear = request.form("frmEDIYear")
			strEdiID = strEDIYear & strSite & strSchool & strTeacher & strClass & strChild 
			
			if request.form("frmConsent") = "true" then 
				strSql = "UPDATE children SET intConsent = 1 WHERE strEDIID = '" & strEDIID &"'"			
			else	
				strSql = "UPDATE children SET intConsent = 0 WHERE strEDIID = '" & strEDIID &"'"
			end if 
			
			conn.execute strSql
		end if 
	end if
	
	
	' get the total number of classes
	rstData.Open "SELECT COUNT(intClassID) FROM classes", conn
	
	' if more than 0 classes
	if not rstData.EOF then 
		' get the total number of classes
		intTotalClasses = rstData(0)
	
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
						Response.Redirect "edi_teacher_class.asp"
					end if 
				end if 
			end if
	
			' close the recordset
			rstData.Close			
					
			' had to change to DISTINCTROW after switching to SP2
			'		cannot select distinct with MEMO field
			'strSql = "SELECT DISTINCTROW c.intClassID, c.intLanguage, t.strName, t.strEmail, c.strComments " & _
				'		"FROM classes c LEFT JOIN Teachers t ON c.intTeacherID = t.intTeacherID " & _
				'		"WHERE c.intTeacherID = " & intTeacher & _
					'" ORDER BY c.intClassID"
            strSql = "SELECT DISTINCT c.intClassID, c.intLanguage, t.strName, t.strEmail, c.strComments " & _
			"FROM classes c LEFT JOIN Teachers t ON c.intTeacherID = t.intTeacherID " & _
			"WHERE c.intTeacherID = " & intTeacher & _
			" ORDER BY c.intClassID"
			' get the school specific teachers
			'response.write strSql
			rstData.Open strSql, conn	
					
			if not rstData.EOF then
				' store info in array
				aData = rstData.GetRows 
												
				' get the number of classes the specified teacher has
				intClasses = ubound(aData,2) + 1							
				
				' get the class
				if Request.QueryString("class").Count = 0 AND intClass = "" then		 
					intClass = right(aData(0,0),1)
				else
					if intClass = "" then 
						for introw = 0 to ubound(aData,2)
							if Request.QueryString("class") = right(aData(0,introw),1) then 
								intClass = Request.QueryString("class")
								exit for 
							end if
						next
					
						if intClass = "" then 
							' user entered a value that is not valid for this user
							Response.Redirect "edi_teacher_class.asp"
						end if 
					end if 
				end if 					
				' load the values
			   call load_values(intTeacher & intClass)
			else
				intClasses = 0
				'call add_mode
				strError = "<font class=""regtextred"">There are no classes for this teacher.</font>"
			end if 
		else
			strError = "<font class=""regtextred"">There are no classes for this teacher.</font>"
			intClasses = 0
		end if 
	' if 0 teachers
	else
		' 0 classes in the database
		intTotalClasses = 0 
		'call add_mode
	end if		 	
		
	' close the recordset
	rstData.Close 
	
	
%>
<html>
<head>
    <!-- added UTF8 Encoding to get rid of funny characters -->
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" /> 
	<!-- Start CSS files-->
		<link rel="stylesheet" type="text/css" href="Styles/edi.css">
		<!-- Bootstrap -->
		<link href="css/bootstrap.min.css" type="text/css" rel="stylesheet">
		<link href="css/bootstrap-toggle.min.css" type="text/css" rel="stylesheet">
		
	<!-- End CSS files -->
	<!-- jQuery (necessary for Bootstrap's JavaScript plugins) -->
    <script src="js/jquery-1.11.3.min.js"></script>
    <!-- Include all compiled plugins (below), or include individual files as needed -->
    <script src="js/bootstrap.min.js"></script>
	<script src="js/bootstrap-toggle.min.js"></script>
	<script language="javascript" type="text/javascript" src="js/form.js"></script>
</head>
<body>
	<!-- #include virtual="/shared/page_header.inc" -->
		
		<br />
		<a class="reglinkMaroon" href="edi_teacher.asp"><%=strHome%></a>&nbsp;<font class="regtextblack">></font>&nbsp;<font class="boldtextblack"><%=lblClassInfo%></font>
		<table border="1" bordercolor="006600" cellpadding="0" cellspacing="0" width="760">
		<tr><td>
			<form name="Screens" method="POST" action="edi_teacher_class.asp"> 
			<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">
			<tr>
				<td align="right" width="430"><font class="headerBlue"><%=lblClassInfo%></font></td>
				<td align="right">
					<input type="hidden" name="Action" value="">
					<input type="hidden" name="classes" value="<%=intTotalClasses%>">
					<input type="hidden" name="strLanguage" value="">
					<input type="hidden" name="hiddenclass" value="<%=intTeacher & intClass%>">
					<%if strError = "" then %>
					<input type="button" value="<%=lblUpdate%>" name="Update" onClick="javascript:update_Class_Comments(this.name);">
					<%end if %>
					<input type="button" value="<%=strExit%>" name="Exit" onClick="javascript:window.location='edi_teacher.asp';">
					&nbsp;
				</td>
			</tr>
			<!-- show error if any -->
			<tr><td colspan="2"><%=strError%></td></tr>
			<!-- end error-->
			</table>
			<%if strError = "" then %>
			<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">
				<tr valign="top">
					<td align="right" nowrap>
						<font class="boldtextblack"><%=lblClass%> :&nbsp;&nbsp;</font>
					</td>
					<td>
						<%		
						' teacher						
						Response.Write "<select name=""teacher"" onChange=""javascript:window.location='edi_teacher_class.asp?teacher=' + this.value;"">"
						for intRow = 0 to ubound(aTeachers,2)						
							Response.Write "<option value = """ & right("000" & aTeachers(0,intRow),8) & """"
							' write the teacher
							if intTeacher = right("000" & aTeachers(0,intRow),8) then 
								Response.Write " selected"
							end if
							Response.Write ">" & right("000" & aTeachers(0,intRow),8) & "</option>"
						next
						Response.Write "</select>"
							
						' classes						
						Response.Write "<select name=""code"" onChange=""javascript:window.location='edi_teacher_class.asp?teacher=' + document.forms.Screens.teacher.options(document.forms.Screens.teacher.selectedIndex).value + '&class=' + this.value;"">"
						for intRow = 0 to ubound(aData,2)						
							Response.Write "<option value = """ & right(aData(0,intRow),1) & """"
							' write the class dsescription name
							if intClass = right(aData(0,intRow),1) then 
								Response.Write " selected"
							end if
							Response.Write ">" & right(aData(0,intRow),1) & "</option>"
						next
						Response.Write "</select>"
						
						'if session("language") = "English" then 	
							' display number of classes at this school for the teacher
							'Response.Write "&nbsp;<font class=""regtextgreen"">This teacher has " & intClasses & " Class"
								
							' if more than one class - plural
							'if intClasses > 1 then 
							'	Response.Write "es"
							'end if 
							'Response.Write " at this school.</font>"
						'else
							' display number of classes at this school for the teacher
							'Response.Write "&nbsp;<font class=""regtextgreen"">Ce professeur a " & intClasses & " classe"
								
							' if more than one class - plural
							'if intClasses > 1 then 
						'		Response.Write "s"
							'end if 
							'Response.Write "  cette cole.</font>"
						'end if 
					%>
					</td>
				</tr>
					<% 
					if Request.Form("Action") <> "Add" AND intclasses > 0 then 
					%>
				<tr valign="top">
					<td align="right">
						<font class="boldtextblack"><%=lblName%> :&nbsp;&nbsp;</font>
					</td>
					<td>
						<font class="regtextblack"><%=strName%></font>
						<!--<input type="text" size="70" name="name" value="<%=strName%>" readonly>-->
					</td>
				</tr>
	
				<tr valign="top">
					<td align="right">
						<font class="boldtextblack"><%=lblTime%> :&nbsp;&nbsp;</font>
					</td>
					<td>
						<%
						'Response.Write "<input type=""text"" size=""70"" name=""classtime"" value="
						Response.Write "<font class=""regtextblack"">" & getClassTime(intClass) & "</font>"
						'Response.Write ">" 
						%>
					</td>
				</tr>
					<% 
					end if 
					%>

				<tr valign="top">
					<td align="right">
						<font class="boldtextblack"><%=lblLanguage%> :&nbsp;&nbsp;</font>
					</td>
					<td>
						<%
						'Response.write "<input type=""text"" size=""70"" name=""language"""						
						response.write "<font class=""regtextblack"">"
						for intRow = 0 to ubound(aClassLanguage)													
							' if that language is selected than show it
							if intLanguage = intRow then 
								'Response.write " value=""" &  aClassLanguage(intRow) & """ readonly"
								response.write aClassLanguage(intRow)
								exit for
							end if 
						next
						response.write "</font>"
						'Response.Write ">"
						%>
					</td>
				</tr>
					<% 
					if Request.Form("Action") <> "Add" AND intclasses > 0 then 
					%>

				<tr valign="top">
					<td align="right">
						<font class="boldtextblack"><%=lblEmail%> :&nbsp;&nbsp;</font>
					</td>
					<td>
						<!--<input type="text" size="70" name="email" value="<%=strEmail%>" readonly>  -->
						<font class="regtextblack"><%=strEmail%></font>
					</td>
				</tr>
					<% 
					end if
					%>				
				<tr valign="top">
					<td align="right">
						<font class="boldtextblack"><%=lblComments%> :&nbsp;&nbsp;</font>
					</td>
					<td>
						<textarea rows="3" cols="61" name="comments"><%=strComments%></textarea>
					</td>
					</tr>
					<%
					' email webmaster  
					Response.Write "<tr><td align=""center"" colspan = ""2""><br />"
					Response.Write "<a href=""mailto:webmaster@e-edi.ca"" class=""reglinkblue"">" & strQuestions & ": webmaster@e-edi.ca</a><br />"
					Response.Write "<a href=""mailto:webmaster@e-edi.ca"" class=""reglinkblue"">" & strNote2009 & ": webmaster@e-edi.ca</a>"
					Response.Write "</td></tr>"
					%>			
				</table>
			</form>
			</td>
			</tr>
							
			
				<%
				' only show classes if on a site not in add mode
				'if Request.Form("Action") <> "Add" AND intclasses > 0 then 
				if intclasses > 0 then 
				%>
		
			<tr>
				<td>
					<form name="Children" method="POST" action="">
						<input type="hidden" name="frmEDIYear" value="">
						<input type="hidden" name="frmSite" value="">
						<input type="hidden" name="frmSchool" value="">
						<input type="hidden" name="frmTeacher" value="">
						<input type="hidden" name="frmClass" value="">					
						<input type="hidden" name="frmChild" value="">
						<input type="hidden" name="frmConsent" value="">
						<input type="hidden" name="frmAction" value="" />
					</form>
						<br />
						<%
							dim previousClass 
							previousClass = ""
							' select all children in the CLASS 				
							if session("Language") = "English" then 
								'strSql = "SELECT strEDIID, strLocalID, IIf([intSex]=1,'M',IIf([intsex]=2,'F','')) AS gender, dtmDOB, strPostal, chkCompleted, dtmDate FROM children WHERE intClassID = " & intSite & intSchool & intTeacher & intClass & " ORDER BY strEDIID"
								strSql = "SELECT s.strName,c2.intClassID,c.strEDIID, c.strLocalID, CASE c.intSex WHEN 1 THEN 'M' WHEN 2 THEN 'F' ELSE '' END AS gender, c.dtmDOB, c.strPostal, c.chkCompleted, c.dtmDate, d.intStatus, c.intConsent FROM  ((children c LEFT JOIN Classes c2 on c2.intClassid = c.intClassid)LEFT JOIN teachers t ON t.intTeacherID = c2.intTeacherID) LEFT JOIN schools s ON t.intSchoolID = s.intSchoolID LEFT JOIN demographics d ON c.strEDIID = d.strEDIID WHERE t.strEmail = '" & session("ID") & "' ORDER BY c.strEDIID"
							else
								'strSql = "SELECT strEDIID, strLocalID, IIf([intSex]=1,'M',IIf([intsex]=2,'F','')) AS gender, dtmDOB, strPostal, chkCompleted, dtmDate FROM children WHERE intClassID = " & intSite & intSchool & intTeacher & intClass & " ORDER BY strEDIID"
								strSql = "SELECT s.strName,c2.intClassID,c.strEDIID, c.strLocalID, CASE c.intSex WHEN 1 THEN 'M' WHEN 2 THEN 'F' ELSE '' END AS gender, c.dtmDOB, c.strPostal, c.chkCompleted, c.dtmDate, d.intStatus, c.intConsent FROM  ((children c LEFT JOIN Classes c2 on c2.intClassid = c.intClassid)LEFT JOIN teachers t ON t.intTeacherID = c2.intTeacherID) LEFT JOIN schools s ON t.intSchoolID = s.intSchoolID LEFT JOIN demographics d ON c.strEDIID = d.strEDIID WHERE t.strEmail = '" & session("ID") & "' ORDER BY c.strEDIID"
							end if 
							                           
							'intLanguage removed - feb 8
							'Response.Write strSQL
							
							' open list of classes and teachers at this CLASS
							rstData.Open strSql, conn		
							'response.write strSql
							dim Classid, previousediid, nextediid
							
							if rstData.EOF and rstData.bof then 
								' show the header
								Call ChildHeader()
								Response.Write "<tr><td colspan=""7"">&nbsp;<font class=""regtextmaroon"">There are no children in this class.</font></td></tr>"
								call ChildFooter()
							else
								do while not rstData.EOF 
									if previousClass <> rstData("intClassid") then 
										if previousClass <> "" then 
											call AddStudent(classid)
											call childFooter
											response.write "</form>"
										end if 

										classid = right("000" & rstData("intClassID"),9)
										response.write "<form name=""Class" & classid & """ method=""POST"" action="""">" 	
										response.write "<input type=""hidden"" name=""Action"" value=""Add"">"
										response.write "<input type=""hidden"" name=""frmEDIYear"" value="""">"
										response.write "<input type=""hidden"" name=""frmSite"" value="""">"
										response.write "<input type=""hidden"" name=""frmSchool"" value="""">"
										response.write "<input type=""hidden"" name=""frmTeacher"" value="""">"
										response.write "<input type=""hidden"" name=""frmClass"" value="""">"					
										response.write "<input type=""hidden"" name=""frmChild"" value="""">"	

									
										response.write "<table border=""1"" cellpadding=""0"" cellspacing=""0"" width=""750"" align=""center"">"							

										if session("province") = 6 or session("province") = 3 then 
											cols = 8
										else 
											cols = 7
										end if 
										
										' show the class information 
										response.write 	"<tr><td colspan=""" & cols & """>&nbsp;<font class=""subheadermaroon"">" & classid & ": " & rstData("strName") '& " - " & GetClassTime(right(rstData("intClassID"),1)) & "</font></td></tr>"
										
										' show the header
										Call ChildHeader()

										previousClass = rstData("intClassid") 
									end if 
									
									previousediid = right("000" & rstData("strEDIID"),15)

									' removed link 20061106
									'Response.Write "<tr><td><a href=""javascript:window.location='edi_teacher_student.asp?teacher=" & left(rstData("strEDIID"),8) & "&class=" & mid(rstData("strEDIID"),9,1) & "&child=" & right(rstData("strEDIID"),2) & "';"" class=""reglinkBlue"">" & right("000" & rstData("strEDIID"),11) & "</a></td>"
									Response.Write "<tr><td><font class=""regtextblack"">" & right("000" & rstData("strEDIID"),15) & "</font></td>"
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
									if session("language") = "English" then 
										Response.Write "<td align=""center""><font class=""regtextblack"">" & day(rstData("dtmDOB")) & "-" & monthname(datepart("m",rstData("dtmDOB")),true) & "-" & year(rstData("dtmDOB")) & "</font></td>"
										
										' changed 2006-02-27
										' Andrew Renner
										if err.number = 94 then 'Invalid use of null - when dob not available
											Response.Write "<td align=""center""><font class=""regtextblack"">NA</font></td>"
											err.Clear 
										end if
									else
										Response.Write "<td align=""center""><font class=""regtextblack"">" & day(rstData("dtmDOB")) & "-" & left(french_month(datepart("m",rstData("dtmDOB"))),3) & "-" & year(rstData("dtmDOB")) & "</font></td>"
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
										if session("Language") = "English" then 
											Response.Write "checked>&nbsp;" & day(rstData("dtmDate")) & "-" & monthname(datepart("m",rstData("dtmDate")),true) & "-" & year(rstData("dtmDate"))
										else
											Response.Write "checked>&nbsp;" & day(rstData("dtmDate")) & "-" & left(French_Month(datepart("m",rstData("dtmDate"))),3) & "-" & year(rstData("dtmDate"))
										end if 
									else
										Response.Write ">"
									end if 
									Response.Write "</font></td>"
									
									'2016
									if session("province") = 3 then 										
										response.write "<td align=""center""><input id=""consent" & rstData("strediid") & """ name=""consent" & rstData("strediid") & """ type=""checkbox"" "
										
										' check consent
										if rstData("intConsent") =1 then 
											response.write "checked disabled data-toggle=""toggle"" data-off=""No"" data-on=""Yes"" data-onstyle=""success"" data-offstyle=""danger"" data-size=""mini""></td>"
											Response.Write "<td><img src=""images/blinkingarrow.gif"" alt=""Blinking Arrow"" /><a href=""javascript:goEDI('edi_teacher_questionnaire.asp','" & left(rstData("strEDIID"),4) & "','" & mid(rstData("strEDIID"),5,3) & "','" & mid(rstData("strEDIID"),8,3) & "','" & mid(rstData("strEDIID"),11,2) & "','" & mid(rstData("strEDIID"),13,1) & "','" & right(rstData("strEDIID"),2) & "');"" class=""boldlinkBlue"">&nbsp" & lblEDI & "</a><!--<img src=""images/blinkingarrowRight.gif"" alt=""Blinking Arrow Right"" />--></td>"
										elseif rstData("intConsent") <> 0 or checkNull(rstData("intConsent")) ="null" then 
											response.write "checked data-toggle=""toggle"" data-off=""No"" data-on=""Yes"" data-onstyle=""success"" data-offstyle=""danger"" data-size=""mini""></td>"
											response.write "<td><button type=""button"" class=""btn btn-primary btn-xs"" onclick=""if(document.getElementById('consent"&rstData("strEdiid")&"').checked){goConfirm_Consent('"& left(rstData("strEDIID"),4) & "','" & mid(rstData("strEDIID"),5,3) & "','" & mid(rstData("strEDIID"),8,3) & "','" & mid(rstData("strEDIID"),11,2) & "','" & mid(rstData("strEDIID"),13,1) & "','" & right(rstData("strEDIID"),2) &"',document.getElementById('consent"&rstData("strEdiid")&"').checked,'consent');}else{alert('" & strAlbertaWarning & "');goConfirm_Lock('"& left(rstData("strEDIID"),4) & "','" & mid(rstData("strEDIID"),5,3) & "','" & mid(rstData("strEDIID"),8,3) & "','" & mid(rstData("strEDIID"),11,2) & "','" & mid(rstData("strEDIID"),13,1) & "','" & right(rstData("strEDIID"),2) &"','lockClassList');};"">Submit</button></td>"
										else
											response.write "disabled data-toggle=""toggle"" data-off=""No"" data-on=""Yes"" data-onstyle=""success"" data-offstyle=""danger"" data-size=""mini""></td>"
											response.write "<td></td>"
										end if 
										
										
										
										'onchange=""javascript:alert('" & strAlbertaWarning & "');goConfirm_Lock('"& left(rstData("strEDIID"),4) & "','" & mid(rstData("strEDIID"),5,3) & "','" & mid(rstData("strEDIID"),8,3) & "','" & mid(rstData("strEDIID"),11,2) & "','" & mid(rstData("strEDIID"),13,1) & "','" & right(rstData("strEDIID"),2) &"','lockClassList');""
										
									else 
										Response.Write "<td><img src=""images/blinkingarrow.gif"" alt=""Blinking Arrow"" /><a href=""javascript:goEDI('edi_teacher_questionnaire.asp','" & left(rstData("strEDIID"),4) & "','" & mid(rstData("strEDIID"),5,3) & "','" & mid(rstData("strEDIID"),8,3) & "','" & mid(rstData("strEDIID"),11,2) & "','" & mid(rstData("strEDIID"),13,1) & "','" & right(rstData("strEDIID"),2) & "');"" class=""boldlinkBlue"">&nbsp" & lblEDI & "</a><!--<img src=""images/blinkingarrowRight.gif"" alt=""Blinking Arrow Right"" />--></td>"
									end if 
									
									
									if session("province") = 6 then 							
										Response.Write "<td><img src=""images/blinkingarrow.gif"" alt=""Blinking Arrow"" /><a href=""javascript:goEDI('edi_teacher_identity.asp','" & left(rstData("strEDIID"),4) & "','" & mid(rstData("strEDIID"),5,3) & "','" & mid(rstData("strEDIID"),8,3) & "','" & mid(rstData("strEDIID"),11,2) & "','" & mid(rstData("strEDIID"),13,1) & "','" & right(rstData("strEDIID"),2) & "');"" class=""boldlinkBlue"">&nbsp" & lblIdentity & "</a><!--<img src=""images/blinkingarrowRight.gif"" alt=""Blinking Arrow Right"" />--></td>"
									end if 
									'Response.Write "<td><a href=""javascript:goFormStudentDelete(" & checkNull(right("000" & rstData("strEDIID"),11)) & ");"" class=""regLinkMaroon"">Delete</a></td>"
									Response.Write "</tr>"
									rstData.MoveNext 

                                    if Request.Form("Action") = "Add" AND request.form("ediID") = rstData("strEDIID") then 
                                        if session("province") <> 3 then 
											Response.Write "<SCRIPT LANGUAGE=""JavaScript"">"
											Response.write "goEDI('edi_teacher_questionnaire.asp','" & left(rstData("strEDIID"),4) & "','" & mid(rstData("strEDIID"),5,3) & "','" & mid(rstData("strEDIID"),8,3) & "','" & mid(rstData("strEDIID"),11,2) & "','" & mid(rstData("strEDIID"),13,1) & "','" & right(rstData("strEDIID"),2) & "');"                                        
											Response.Write "</SCRIPT>"
										end if 
                                    end if 
								loop
								if previousClass <> "" then 
									call AddStudent(classid)
									Call childFooter
									response.write "</form>"
								end if 
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
		<%
		else
			Response.Write "<br />"
		end if 
		%>		
		</table>
		
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

sub childHeader
		response.write "<tr>"
			response.write "<td align=""center""><font class=""boldtextblack"">" & lblEDIID& "</font></td>"
			response.write "<td align=""center""><font class=""boldtextblack"">" & lblLocal& "</font></td>"
			response.write "<td align=""center""><font class=""boldtextblack"">" & lblSex& "</font></td>"
			response.write "<td align=""center""><font class=""boldtextblack"">" & lblDOB& "</font></td>"
			response.write "<td align=""center""><font class=""boldtextblack"">" & lblPostal& "</font></td>"
			response.write "<td align=""center""><font class=""boldtextblack"">" & lblStatus& "</font></td>"		
			if session("province") = 3 then 
				response.write "<td align=""center""><font class=""boldtextblack"">" & lblConsent & "</font></td>"
			end if
			response.write "<td align=""center""><font class=""boldtextblack"">" & lblEDI & "</font></td>"
			if session("province") = 6 then 
				response.write "<td align=""center""><font class=""boldtextblack"">" & lblIdentity & "</font></td>"
			end if 
		response.write "</tr>"
end sub 

sub ChildFooter
	response.write "</table>"
end sub 

sub AddStudent(ClassID)
	nextediid = left(previousediid,4) & classid & right(previousediid+1,2)
	response.write "<tr>"
	response.write "<td colspan=""1""><input type=""text"" value=""" & nextediid & """ size=""16"" name=""ediid"" readonly /></td>"
	response.write "<td colspan=""1""><input type=""text"" size=""10"" name=""localid"" /></td>"
	response.write "<td colspan=""" & cols - 2 & """><input type=""button"" value=""" & lblAddStudent & """ name=""Add"" onClick=""javascript:goAddStudent('" & classid & "');""></td>"
	response.write "</tr>"	
end Sub

function GetClassTime(intClass)
	for intRow = 0 to ubound(aClass,2)						
		' show the selected class time
		if intClass = right(aClass(0,intRow),1) then
			'	Response.Write """" & aClass(1,intRow) & """ readonly"
			'Response.Write "<font class=""regtextblack"">" & aClass(1,intRow) & "</font>"
			' if NFL 
			if session("province") = 7 then 
				GetClassTime = "Other" 	
			else
				GetClassTime = aClass(1,intRow)
			end if 
			exit for
		end if 
	next
end function
%>