<!-- #include virtual="/shared/admin_security.asp" -->
<%
' if the user has not logged in they will not be able to see the page
if blnSecurity then
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
	<%
	dim aSites, aSchools, aTeachers, aClasses', aChildren, aChild
	dim strTable
	' open edi connection
	'call open_adodb(conn, "EDI")
	'call open_adodb(conn, "DATA")
    call open_adodb(conn, "MACEDI")
	set rstData = server.CreateObject("adodb.recordset")
	
	' put the EDI ID together
	strEDIID = ""
	for each item in Request.Form
		strEDIID = strEDIID & Request.form(item)
	next

	' get individual variables		
	strChild = Request.Form("frmChild")
	strClass = Request.Form("frmClass")
	strTeacher = Request.Form("frmTeacher")
	strSchool = Request.Form("frmSchool")
	strSite = Request.Form("frmSite")
	
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Check for child ID 
	'	- if child id is there then show the child
	'			- build EDI string
	'	- else check for class ID
	'		- if class ID then show its students
	'		- else check for teacher ID
	'			- if teacher ID then show their classes
	'			- else check for school ID
	'				- if school ID then show their teachers
	'				- else check for site ID
	'					- if site ID then show their teachers
	'					- else show all sites WITH classes!!!
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	' start child section
	'if strChild <> "" then 
	'	strMap = "&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""edi_admin_reports_questionnaire.asp"">Site Selection</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strSite & "','','','','');"">School Selection</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strSite & "','" & strSchool & "','','','');"">Teacher Selection</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strSite & "','" & strSchool & "','" & strTeacher  & "','','');"">Class Selection</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strSite & "','" & strSchool & "','" & strTeacher  & "','" & strClass  & "','');"">Child Selection</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<font class=""boldtextblack"">Child Questionnaire</font>"
	
		' get the childs data
	'	strSql = "SELECT * FROM children WHERE strEDIID = '" & strEDIID & "' ORDER BY strEDIID"
		
	'	rstData.Open strSql, conn
	'	if not rstData.eof then 
	'		aChild = rstData.getrows
	'	else
	'		strError = "<font class=""regtextred"">No data on child - " & strEDIID & "</font>"
	'	end if
		
		' close the recordset
	'	rstData.close
	' no child in the form data
	'else
		' start class section
		if strClass <> "" then 
			strMap = "&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""edi_admin_reports_questionnaire.asp"">Site Selection</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strSite & "','','','','');"">School Selection</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strSite & "','" & strSchool & "','','','');"">Teacher Selection</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strSite & "','" & strSchool & "','" & strTeacher  & "','','');"">Class Selection</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<font class=""boldtextblack"">Child Selection for EDI Summary Report</font>"
	
			' get the class list 
			strSql = "SELECT * FROM children WHERE intClassID = " & strEDIID & " ORDER BY intChild"
			
			rstData.Open strSql, conn
			if not rstData.eof then 
				aChildren = rstData.getrows
			else
				strError = "<font class=""regtextred"">No data on class - " & strEDIID & "</font>"
			end if
			
			' close the recordset
			rstData.close
		' no class in the form data
		else
			' start teacher section
			if strTeacher <> "" then 
				strMap = "&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""edi_admin_reports_questionnaire.asp"">Site Selection</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strSite & "','','','','');"">School Selection</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strSite & "','" & strSchool & "','','','');"">Teacher Selection</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<font class=""boldtextblack"">Class Selection for EDI Summary Report</font>"
	
				' get the class data - this teachers classes
				strSql = "SELECT DISTINCT c.intClassID FROM classes AS c RIGHT JOIN children AS ch ON c.intClassID = ch.intClassID WHERE c.intTeacherID = " & strEDIID & " ORDER BY c.intClassID"
				
				rstData.Open strSql, conn
				if not rstData.eof then 
					aClasses = rstData.getrows
				else
					strError = "<font class=""regtextred"">No class data on teacher - " & strEDIID & "</font>"
				end if
				
				' close the recordset
				rstData.close
			' no teacher in the form data
			else
				' start school section
				if strSchool <> "" then 
					strMap = "&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""edi_admin_reports_questionnaire.asp"">Site Selection</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strSite & "','','','','');"">School Selection</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<font class=""boldtextblack"">Teacher Selection for EDI Summary Report</font>"
	
					' get the teacher data - this schools teachers
					strSql = "SELECT DISTINCT t.intTeacherID, t.intSchoolID, t.strName FROM teachers AS t RIGHT JOIN (classes AS c RIGHT JOIN children AS ch ON c.intClassID = ch.intClassID) ON t.intTeacherID = c.intTeacherID WHERE t.intSchoolID = " & strEDIID & " ORDER BY t.intTeacherID"
					
					rstData.Open strSql, conn
					if not rstData.eof then 
						aTeachers = rstData.getrows
					else
						strError = "<font class=""regtextred"">No teacher data on school - " & strEDIID & "</font>"
					end if
					
					' close the recordset
					rstData.close
				' no school in the form data
				else
					' start site section
					if strSite <> "" then 
						strMap = "&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""edi_admin_reports_questionnaire.asp"">Site Selection</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<font class=""boldtextblack"">School Selection for EDI Summary Report</font>"
	
						' get the school data - this sites schools
						strSql = "SELECT DISTINCT sc.intSchoolID, sc.intSiteID, sc.strName FROM ((schools AS sc RIGHT JOIN teachers AS t ON sc.intSchoolID = t.intSchoolID) RIGHT JOIN classes AS c ON t.intTeacherID = c.intTeacherID) RIGHT JOIN children ch ON c.intClassID = ch.intClassID WHERE sc.intSiteID = " & strEDIID & " ORDER BY sc.intSchoolID"
							
						rstData.Open strSql, conn
						if not rstData.eof then 
							aSchools = rstData.getrows
						else
							strError = "<font class=""regtextred"">No class data on teacher - " & strEDIID & "</font>"
						end if
							
						' close the recordset
						rstData.close
					' no site in the form data
					else
						strMap = "&nbsp;<font class=""regtextblack"">></font>&nbsp;<font class=""boldtextblack"">Site Selection for EDI Summary Report</font>"
	
						' get the site data - sites that have valid classes
						strSql = "SELECT DISTINCT s.intSiteID, s.strName, s.strCity FROM ((sites AS s RIGHT JOIN (schools AS sc RIGHT JOIN teachers AS t ON sc.intSchoolID = t.intSchoolID) ON s.intSiteID = sc.intSiteID) RIGHT JOIN classes AS c ON t.intTeacherID = c.intTeacherID) RIGHT JOIN children ch ON c.intClassID = ch.intClassID ORDER BY s.intSiteID"
								
						rstData.Open strSql, conn
						if not rstData.eof then 
							aSites = rstData.getrows
						else
							strError = "<font class=""regtextred"">No site with available classes.</font>"
						end if	
						' close the recordset
						rstData.close
					' end site section
					end if 
				' end school section
				end if 
			' end teacher section
			end if 
		' end class section
		end if 
	' end child section
	'end if
	%>
	<form name="Children" method="POST" action="edi_admin_reports_questionnaire.asp"> 
		<input type="hidden" name="frmSite" value="">
		<input type="hidden" name="frmSchool" value="">
		<input type="hidden" name="frmTeacher" value="">
		<input type="hidden" name="frmClass" value="">					
		<input type="hidden" name="frmChild" value="">
	</form>
	<form name="Screens" method="POST" action="edi_admin_reports.asp" target="Reports"> 
		<input type="hidden" name="Student" value="">
		<input type="hidden" name="rpt" value="">
		<input type="hidden" name="XML" value="">		
		<a class="reglinkMaroon" href="edi_admin.asp">Home</a><%=strMap%>
		<table border="1" bordercolor="006600" cellpadding="0" cellspacing="0" width="760">
			<tr>
				<td>
					<table border="0" cellpadding="0" cellspacing="0" width="760" align="center">
						<tr>
							<td align="right" width="480">
								<font class="headerBlue">EDI Summary(<%=strEDIID%>)</font>
							</td>
							<td align="right">
								<input type="button" value="Exit" name="Exit" title="EXIT Screen" onClick="javascript:window.location='edi_admin.asp';">
								&nbsp;
							</td>
						</tr>
						<tr><td><%="<br/>" & strError%></td></tr>
					</table>
					
					<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">						
						<tr>
							<td>
																
							</td>
						</tr>
					</table>
					
						<%
						select case len(strMap)
							case 675 ' child
								Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""500"" align=""center"">"						
								Response.Write "<tr><td>"	
								if strError = "" then 
									for intRow = 0 to ubound(aChildren,2)
										Response.Write "<font class=""regTextBlack"">" & right("00" & aChildren(1,intRow),2) & ") </font>"
										Response.Write "<a class=""reglinkMaroon"" href=""javascript:goEDIReport('" & strEDIID & "');"">" & day(aChildren(5,intRow)) & "-" & monthname(datepart("m",aChildren(5,intRow)),true) & "-" & year(aChildren(5,intRow)) & " - " & aChildren(3,intRow) & "</a>"
										Response.Write "<br/>"
									next
								end if 
								Response.Write "</td></tr></table>"	
							case 533 ' class
								'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
								' get the class times
								'//////////////////////////////////////////////////////////////////////	
								set rstClass = server.CreateObject("adodb.recordset")
	
								' open all languages
								rstClass.Open "SELECT intClassID, English FROM [LU Classes]", conn
								
								' store all languages in array
								aClass = rstClass.GetRows 
	
								' close and kill the langauges recordset
								call close_adodb(rstClass)
							
								Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""500"" align=""center"">"						
								Response.Write "<tr><td><img src=""images/greenbb.gif"">&nbsp;<a class=""reglinkMaroon"" href=""javascript:goEDIReport('" & strEDIID & "');"">All Students for teacher " & strEdIID & "</a>&nbsp;<img src=""images/greenbb.gif""><br /><br /></td><tr>"
								Response.Write "<tr><td>"	
								if strError = "" then 
									for intRow = 0 to ubound(aClasses,2)
										Response.Write "<font class=""regTextBlack"">" & right("0" & aClasses(0,intRow),1) & ") </font>"
										Response.Write "<a class=""reglinkMaroon"" href=""javascript:goChild('" & left(strEDIID,3) & "','" & mid(strEDIID,4,3) & "','" & right(strEDIID,2) & "','" & right(aClasses(0,intRow),1) & "','');"">"  & aClass(1,right(aClasses(0,intRow),1)) & "</a>"
										' show class description
										Response.Write "<br/>"
									next
								end if 
								Response.Write "</td></tr></table>"	
							case 393	' teacher
								Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""500"" align=""center"">"						
								Response.Write "<tr><td><img src=""images/greenbb.gif"">&nbsp;<a class=""reglinkMaroon"" href=""javascript:goEDIReport('" & strEDIID & "');"">All Students at school " & strEdIID & "</a>&nbsp;<img src=""images/greenbb.gif""><br /><br /></td><tr>"
								Response.Write "<tr><td>"	
								if strError = "" then 
									for intRow = 0 to ubound(aTeachers,2)
										Response.Write "<font class=""regTextBlack"">" & right(aTeachers(0,intRow),2) & ") </font>"
										Response.Write "<a class=""reglinkMaroon"" href=""javascript:goChild('" & left(strEDIID,3) & "','" & mid(strEDIID,4,3) & "','" & right(aTeachers(0,intRow),2) & "','','');"">" & aTeachers(2,intRow) & "</a>"
										Response.Write "<br/>"
									next
								end if 
								Response.Write "</td></tr></table>"
							case 254 ' school
								Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""500"" align=""center"">"						
								'Response.Write "<tr><td><img src=""images/greenbb.gif"">&nbsp;<a class=""reglinkMaroon"" href=""javascript:goEDIReport('" & strEDIID & "');"">All Students at site " & strEdIID & "</a>&nbsp;<img src=""images/greenbb.gif""><br /><br /></td><tr>"
								'Response.Write "<tr><td><br /><br /></td><tr>"
								Response.Write "<tr><td>"	
								if strError = "" then 
									for intRow = 0 to ubound(aSchools,2)
										Response.Write "<font class=""regTextBlack"">" & right("000" & aSchools(0,intRow),3) & ") </font>"
										Response.Write "<a class=""reglinkMaroon"" href=""javascript:goChild('" & left(strEDIID,3) & "','" & right(aSchools(0,intRow),3) & "','','','');"">" & aSchools(2,intRow) & "</a>"
										Response.Write "<br/>"
									next
								end if 
								Response.Write "</td></tr></table>"
							case 119	' site
								Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""500"" align=""center"">"						
								'Response.Write "<tr><td><img src=""images/greenbb.gif"">&nbsp;<a class=""reglinkMaroon"" href=""javascript:goEDIReport('0');"">All Students</a>&nbsp;<img src=""images/greenbb.gif""><br /><br /></td><tr>"
								'Response.Write "<tr><td><br /><br /></td><tr>"
								Response.Write "<tr><td>"	
								if strError = "" then 
									for intRow = 0 to ubound(aSites,2)
										Response.Write "<font class=""regTextBlack"">" & right("000" & aSites(0,intRow),3) & ") </font>"
										Response.Write "<a class=""reglinkMaroon"" href=""javascript:goChild('" & right("000" & aSites(0,intRow),3) & "','','','','');"">" & aSites(1,intRow) & "</a>"
										Response.Write "<br/>"
									next
								end if 
								Response.Write "</td></tr></table>"
							case else
								Response.Write "Incorrect Length - " & len(strMap)
						end select
						%>	
					<br/> 
				</td>
			</tr>
		</table>
		<br/> 
	</form>
	
	<!-- #include virtual="/shared/page_footer.inc" -->
</body>
</html>
<%
	call close_adodb(rstData)
	'call close_adodb(conn)
	call close_adodb(conn)
end if



sub buildSites()
	
end sub
%>
