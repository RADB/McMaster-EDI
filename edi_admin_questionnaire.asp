<!-- #include virtual="/shared/admin_security.asp" -->
<%
' if the user has not logged in they will not be able to see the page
if blnSecurity then 
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
	<%
	dim aSites, aSchools, aTeachers, aClasses, aChildren, aChild, aStudent
	dim strTable
	
	if session("Language") = "English" then 
		strLanguage = "English"
		intLanguage = 1
	else
		strLanguage = "French"
		intLanguage = 2
	end if
	
	' open edi connection
	'call open_adodb(conn, "EDI")
	'call open_adodb(conn, "DATA")
	'call open_adodb(conn_tables, "TABLES")
	call open_adodb(conn, "MACEDI")

	set rstData = server.CreateObject("adodb.recordset")
	
	'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
	' get the languages for the drop down box
	'//////////////////////////////////////////////////////////////////////	
	set rstLanguages = server.CreateObject("adodb.recordset")
	
	' open all languages
	rstLanguages.Open "SELECT LID, english, french FROM [LU Languages] ORDER BY LID", conn
	
	' store all languages in array
	aLanguages = rstLanguages.GetRows 
	
	' close and kill the langauges recordset
	call close_adodb(rstLanguages)
	
	' put the EDI ID together
	strEDIID = ""
	'for each item in Request.Form
	'	strEDIID = strEDIID & Request.form(item)
	'next
	
	' get individual variables		
	
	if request.form("frmEDIYear") then 
		strEDIYear = request.form("frmEDIYear")
	else
		if month(date) > 8 then 
			strEDIYear = year(date)
		else
			strEDIYear = year(date) - 1
		end if 
	end if 
	strChild = Request.Form("frmChild")
	strClass = Request.Form("frmClass")
	strTeacher = Request.Form("frmTeacher")
	strSchool = Request.Form("frmSchool")
	strSite = Request.Form("frmSite")
	strEdiID = strEDIYear & strSite & strSchool & strTeacher & strClass & strChild
	 	
	if Request.form("frmAction") = "lock" then  
		conn.execute "UPDATE children SET chkCompleted = 1, dtmDate = '" & date & "' WHERE strEDIID = '" & strEDIID & "'"	
	elseif Request.form("frmAction") = "unlock" then
		conn.execute "UPDATE children SET chkCompleted = 0, dtmDate = null WHERE strEDIID = '" & strEDIID & "'"			
	end if 
	

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
	if strChild <> "" then 
		strMap = "&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""edi_admin_questionnaire.asp"">Site</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strEDIYear & "','" & strSite & "','','','','');"">School</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strEDIYear & "','" & strSite & "','" & strSchool & "','','','');"">Teacher</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher  & "','','');"">Class</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher  & "','" & strClass  & "','');"">Student</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<font class=""boldtextblack"">EDI Questionnaire</font>"
	
		' get the childs data
		strSql = "SELECT * FROM Students WHERE strEDIID = '" & strEDIID & "' ORDER BY strEDIID"
		
		rstData.Open strSql, conn
		if not rstData.eof then 
			aStudent = rstData.getrows
		else
			strError = "<font class=""regtextred"">No data on child - " & strEDIID & "</font>"
		end if
		
		' close the recordset
		rstData.close
	' no child in the form data
	else
		' start class section
		if strClass <> "" then 
			strMap = "&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""edi_admin_questionnaire.asp"">Site</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strEDIYear & "','" & strSite & "','','','','');"">School</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strEDIYear & "','" & strSite & "','" & strSchool & "','','','');"">Teacher</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher  & "','','');"">Class</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<font class=""boldtextblack"">Student</font>"
	
			' get the class list 
			strSql = "SELECT * FROM children WHERE intClassID = " & strSite & strSchool & strTeacher & strClass & " ORDER BY intChild"
			
			rstData.Open strSql, conn
			if not rstData.eof then 
				aChildren = rstData.getrows
			else
				strError = "<font class=""regtextred"">No data on class - " & strSite & strSchool & strTeacher & strClass & "</font>"
			end if
			
			' close the recordset
			rstData.close
		' no class in the form data
		else
			' start teacher section
			if strTeacher <> "" then 
				strMap = "&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""edi_admin_questionnaire.asp"">Site</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strEDIYear & "','" & strSite & "','','','','');"">School</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strEDIYear & "','" & strSite & "','" & strSchool & "','','','');"">Teacher</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<font class=""boldtextblack"">Class</font>"
	
				' get the class data - this teachers classes
				strSql = "SELECT * FROM classes WHERE intTeacherID = " & strSite & strSchool & strTeacher  & " ORDER BY intClassID"
				
				rstData.Open strSql, conn
				if not rstData.eof then 
					aClasses = rstData.getrows
				else
					strError = "<font class=""regtextred"">No class data on teacher - " & strSite & strSchool & strTeacher & "</font>"
				end if
		
				' close the recordset
				rstData.close
			' no teacher in the form data
			else
				' start school section
				if strSchool <> "" then 
					strMap = "&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""edi_admin_questionnaire.asp"">Site</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strEDIYear & "','" & strSite & "','','','','');"">School</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<font class=""boldtextblack"">Teacher</font>"
	
					' get the teacher data - this schools teachers
					strSql = "SELECT * FROM teachers WHERE intSchoolID = " & strSite & strSchool  & " ORDER BY intTeacherID"
					
					rstData.Open strSql, conn
					if not rstData.eof then 
						aTeachers = rstData.getrows
					else
						strError = "<font class=""regtextred"">No teacher data on school - " & strSite & strSchool & "</font>"
					end if
					
					' close the recordset
					rstData.close
				' no school in the form data
				else
					' start site section
					if strSite <> "" then 
						strMap = "&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""edi_admin_questionnaire.asp"">Site</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<font class=""boldtextblack"">School</font>"
	
						' get the school data - this sites schools
						strSql = "SELECT * FROM schools WHERE intSiteID = " & strSite & " ORDER BY intSchoolID"
							
						rstData.Open strSql, conn
						if not rstData.eof then 
							aSchools = rstData.getrows
						else
							strError = "<font class=""regtextred"">No class data on teacher - " & strSite & "</font>"
						end if
							
						' close the recordset
						rstData.close
					' no site in the form data
					else
						strMap = "&nbsp;<font class=""regtextblack"">></font>&nbsp;<font class=""boldtextblack"">Site</font>"
	
						' get the site data - sites that have valid classes
						strSql = "SELECT DISTINCT s.intSiteID, s.strName, s.strCity FROM sites s RIGHT JOIN ((schools sc RIGHT JOIN teachers t ON sc.intSchoolID = t.intSchoolID) RIGHT JOIN classes c ON t.intTeacherID = c.intTeacherID) ON s.intSiteID = sc.intSiteID ORDER BY s.intSiteID"
								
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
	end if
	%>
	<form name="Children" method="POST" action="edi_admin_questionnaire.asp"> 
		<input type="hidden" name="Student" value="">
		<input type="hidden" name="XML" value="">
		<input type="hidden" name="rpt" value="">
		<input type="hidden" name="frmEDIYear" value="">
		<input type="hidden" name="frmSite" value="">
		<input type="hidden" name="frmSchool" value="">
		<input type="hidden" name="frmTeacher" value="">
		<input type="hidden" name="frmClass" value="">					
		<input type="hidden" name="frmChild" value="">
		<input type="hidden" name="frmAction" value="">
	</form>
	<form name="Screens" method="POST" action="edi_admin_questionnaire.asp"> 
		<a class="reglinkMaroon" href="edi_admin.asp">Home</a><%=strMap%>
		<table border="1" bordercolor="006600" cellpadding="0" cellspacing="0" width="760">
			<tr>
				<td>
					<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">
						<tr>
							<td align="right" width="520"><font class="headerBlue">EDI Questionnaire(<%=strEDIID%>)</font></td>
							<td align="right">
								<input type="button" value="Exit" name="Exit" title="EXIT Screen" onClick="javascript:window.location='edi_admin.asp';">
								&nbsp;
							</td>	
						</tr>
						<tr><td colspan="2"><%="<br/>" & strError%></td></tr>
					</table>
					
					<%
					
					select case len(strMap)
						case 769,741 ' Questionnaire
							' aChild
							if strError = "" then 
								' summary 
								Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""750"" align=""left"">"														
									' status
									Response.Write "<tr>"	
									Response.Write "<td align=""right""><font class=""boldTextBlack"">Status:&nbsp;</font></td>"
									Response.Write "<td align=""left""><font class=""regTextBlack"">"
										if aStudent(7,0) = true  then 
											Response.Write "<font class=""regTextGreen"">&nbsp;&nbsp;Completed and Locked"
										else
											Response.Write "<font class=""regTextRed"">&nbsp;&nbsp;Incomplete and Unlocked"
											end if 
									Response.write "</font></td>"
									Response.Write "<td width = ""350"" rowspan=""3"" align=""left"">&nbsp;&nbsp;<a href=""javascript:confirm_Unlock('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','unlock');"" class=""bigLinkBlue""><img src=""images/unlock.gif"" border=""0"">Unlock Child</a></td>"
									Response.Write "</tr>"
									
									' spacer
									Response.Write "<tr><td colspan=""2"">&nbsp;</td></tr>"
									
									' site
									Response.Write "<tr>"	
									Response.Write "<td align=""right""><font class=""boldTextBlack"">Site:&nbsp;</font></td>"
									Response.Write "<td align=""left""><font class=""lrgregTextBlack"">&nbsp;&nbsp;" & aStudent(10,0) & "</font></td>"
									Response.Write "</tr>"
										
									' school
									Response.Write "<tr>"	
									Response.Write "<td align=""right""><font class=""boldTextBlack"">School:&nbsp;</font></td>"
									Response.Write "<td align=""left""><font class=""lrgregTextBlack"">&nbsp;&nbsp;" & aStudent(11,0) & "</font></td>"
									Response.Write "<td width = ""350"" rowspan=""3"" align=""left"">&nbsp;&nbsp;<a href=""javascript:confirm_Lock('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','lock','" & strConfirmLanguage & "','" & strConfirmLanguage2 & "');"" class=""bigLinkBlue""><img src=""images/lock.gif"" border=""0"">Lock Child</a></td>"
									Response.Write "</tr>"
										
									' teacher
									Response.Write "<tr>"	
									Response.Write "<td align=""right""><font class=""boldTextBlack"">Teacher:&nbsp;</font></td>"
									Response.Write "<td align=""left""><font class=""lrgregTextBlack"">&nbsp;&nbsp;" & aStudent(12,0) & "</font></td>"
									Response.Write "</tr>"
										
									' local ID
									Response.Write "<tr>"	
									Response.Write "<td align=""right""><font class=""boldTextBlack"">Local ID:&nbsp;</font></td>"
									Response.Write "<td align=""left""><font class=""lrgregTextBlack"">&nbsp;&nbsp;" & aStudent(3,0) & "</font></td>"
									'Response.Write "<td width = ""350"" rowspan=""2"">&nbsp;</td>"
									Response.Write "</tr>"
										
									' gender
									Response.Write "<tr>"	
									Response.Write "<td align=""right""><font class=""boldTextBlack"">Gender:&nbsp;</font></td>"
									Response.Write "<td align=""left""><font class=""lrgregTextBlack"">&nbsp;&nbsp;" & aStudent(4,0) & "</font></td>"
									Response.Write "<td width = ""350"" rowspan=""3"" align=""left"">&nbsp;&nbsp;&nbsp;<a href=""javascript:goAdminEDIReport('" & strEDIID & "');"" class=""bigLinkBlue""><img border=""0"" src=""images/details.gif"">View Student Summary</a></td>"
									Response.Write "</tr>"
										
									' DOB
									Response.Write "<tr>"	
									Response.Write "<td align=""right""><font class=""boldTextBlack"">DOB:&nbsp;</font></td>"
                                    if not isnull(aStudent(5,0)) then 												
									    Response.Write "<td align=""left""><font class=""lrgregTextBlack"">&nbsp;&nbsp;" & right("00" & day(aStudent(5,0)),2) & "-" & monthname(datepart("m",aStudent(5,0)),true) & "-" & year(aStudent(5,0)) & "</font></td>"
                                    else
                                        Response.Write "<td align=""left""><font class=""lrgregTextBlack"">&nbsp;&nbsp;</font></td>"
                                    end if 
									'Response.Write "<td width = ""350"" rowspan=""2"">&nbsp;</td>"
									Response.Write "</tr>"
										
									' Postal Code
									Response.Write "<tr>"	
									Response.Write "<td align=""right""><font class=""boldTextBlack"">Postal Code:&nbsp;</font></td>"
									Response.Write "<td align=""left""><font class=""lrgregTextBlack"">&nbsp;&nbsp;" & aStudent(6,0) & "</font></td>"
									'Response.Write "<td width = ""350"">&nbsp;</td>"
									Response.Write "</tr>"
									
									' spacer 
									Response.Write "<tr><td><br /></td></tr>"
								Response.Write "</table>"	
							' end first column	
							Response.Write "</td></tr>"
							
							
							' start second row
							Response.Write "<tr><td>"	
								Response.Write "<br />&nbsp;&nbsp;<font class=""subheaderBlue"">Demographics</font><br />"
									
								' get all the demographic questions
								strSql = "SELECT question, english, french FROM Page_Section_Demographics WHERE Question>0 AND [Option]=0 ORDER BY Question"
								
								'open the demographic questions 
								rstData.Open strSql, conn
								' Menu
								Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""750"" align=""center"">"																					
									Response.Write "<tr><td colspan=""2""><br /></td></tr>"
									
									do while not rstData.EOF 
										Response.Write "<tr>"	
										Response.Write "<td align=""right"" width=""100"">&nbsp;</td><td align=""left""><font class=""boldTextBlack"">" 
										
										if rstData("question") < 10 then 
											Response.Write "&nbsp;&nbsp;"
										end if 
										Response.Write rstData("question") & " ) &nbsp;&nbsp;" &  rstData(strLanguage) & " - </font>"
										Response.Write "<font class=""lrgRegTextBlack"">"
			'previous							
' Date of completion		case 1
' DOB						case 2
' Language Section			case 3
' Class Assignment			case 4
' Class Type				case 5
' Student Status			case 6
' Sex						case 7
' Postal Code				case 8
' ESL						case 9
 'Special Needs				case 10
' Aboriginal				case 11
' French Immersion			case 12
' Other Immersion			case 13
		' 2004
' Class Assignment			case 1
' DOB						case 2
' Sex						case 3
' Postal Code				case 4
' Class Type				case 5
' Date of completion		case 6
' Special Needs				case 7
' ESL						case 8
' French Immersion			case 9
' Other Immersion			case 10
' Aboriginal				case 11
' Language Section			case 12
' Communicates				case 13	
' Student Status			case 14
' Repeat the grade			case 15			
										select case rstData("question")
											' Class Assignment
											case 1
												Response.write aStudent(16,0)
											' DOB
											case 2
                                                if not isnull(aStudent(5,0)) then 	    											
		    										Response.Write right("00" & day(aStudent(5,0)),2) & "-" & monthname(datepart("m",aStudent(5,0)),true) & "-" & year(aStudent(5,0))		
                                                else
                                                    
                                                end if 
											' Sex	
											case 3
												Response.Write aStudent(4,0)
											' Postal Code
											case 4
												Response.Write aStudent(6,0)
											' Class Type
											case 5
												Response.write aStudent(17,0)
											' Date of completion
											case 6
												if not isnull(aStudent(8,0)) then 
													Response.Write right("00" & day(aStudent(8,0)),2) & "-" & monthname(datepart("m",aStudent(8,0)),true) & "-" & year(aStudent(8,0)) 		
												else
													Response.write "Incomplete"
												end if 
											' Special Needs
											case 7
												Response.write aStudent(20,0)
											' ESL
											case 8
												Response.write aStudent(19,0)
											' French Immersion
											case 9
												Response.write aStudent(22,0)
											' Other Immersion
											case 10
												Response.write aStudent(23,0)
											' Aboriginal
											case 11
												Response.write aStudent(21,0)
											' Language 1
											case 12
												for intRow = 0 to ubound(aLanguages,2)															
												' show the language
													if aStudent(14,0) = aLanguages(0,intRow) then 
														Response.write aLanguages(intLanguage,introw)
														exit for
													end if 
												next
												' Language 2
												for intRow = 0 to ubound(aLanguages,2)															
												' show the language
													if aStudent(15,0) = aLanguages(0,intRow) then 
														Response.write " & " & aLanguages(intLanguage,introw)
														exit for
													end if 
												next
											' Communicates
											case 13
												Response.write aStudent(24,0)
											' Student Status
											case 14
												Response.write aStudent(18,0)
											' Repeat the grade			
											case 15			
												Response.write aStudent(25,0)
											
										end select
										Response.Write "</font></td>"
										Response.Write "</tr>"
										rstData.movenext
									loop
								Response.Write "</table>"		
							end if 
							
						case 617,596 ' child
							Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""500"" align=""center"">"						
							Response.Write "<tr><td>"	
							if strError = "" then 
								for intRow = 0 to ubound(aChildren,2)
									Response.Write "<font class=""regTextBlack"">" & right("00" & aChildren(1,intRow),2) & ") </font>"
									Response.Write "<a class=""reglinkMaroon"" href=""javascript:goChild('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & right("00" & aChildren(1,intRow),2) & "');"">" & day(aChildren(5,intRow)) & "-" & monthname(datepart("m",aChildren(5,intRow)),true) & "-" & year(aChildren(5,intRow)) & " - " & aChildren(3,introw) & "</a>"
									Response.Write "<br/>"
								next
							end if 
							Response.Write "</td></tr></table>"	
						case 476, 462 ' class
							'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
							' get the class times
							'//////////////////////////////////////////////////////////////////////	
							set rstClass = server.CreateObject("adodb.recordset")
	
							' open all languages
							rstClass.Open "SELECT intClassID, " & session("Language") & " as strDescription FROM [LU Classes]", conn
								
							' store all languages in array
							aClass = rstClass.GetRows 
	
							' close and kill the langauges recordset
							call close_adodb(rstClass)
							
							Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""500"" align=""center"">"						
							Response.Write "<tr><td>"	
							if strError = "" then 
								for intRow = 0 to ubound(aClasses,2)
									Response.Write "<font class=""regTextBlack"">" & right(aClasses(0,intRow),1) & ") </font>"
									Response.Write "<a class=""reglinkMaroon"" href=""javascript:goChild('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & right(aClasses(0,intRow),1) & "','');"">" & aClass(1,right(aClasses(0,intRow),1)) & "</a>"
									' show class description
									Response.Write "<br/>"
								next
							end if 
							Response.Write "</td></tr></table>"	
						case 339, 332	' teacher
							Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""500"" align=""center"">"						
							Response.Write "<tr><td>"	
							if strError = "" then 
								for intRow = 0 to ubound(aTeachers,2)
									Response.Write "<font class=""regTextBlack"">" & right(aTeachers(0,intRow),2) & ") </font>"
									Response.Write "<a class=""reglinkMaroon"" href=""javascript:goChild('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & right(aTeachers(0,intRow),2) & "','','');"">" & aTeachers(2,intRow) & "</a>"
									Response.Write "<br/>"
								next
							end if 
							Response.Write "</td></tr></table>"
						case 203 ' school
							Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""500"" align=""center"">"						
							Response.Write "<tr><td>"	
							if strError = "" then 
								for intRow = 0 to ubound(aSchools,2)
									Response.Write "<font class=""regTextBlack"">" & right("000" & aSchools(0,intRow),3) & ") </font>"
									Response.Write "<a class=""reglinkMaroon"" href=""javascript:goChild('" & strEDIYear & "','" & strSite & "','" & right(aSchools(0,intRow),3) & "','','','');"">" & aSchools(2,intRow) & "</a>"
									Response.Write "<br/>"
								next
							end if 
							Response.Write "</td></tr></table>"
						case 86	' site
							Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""500"" align=""center"">"						
							Response.Write "<tr><td>"	
							if strError = "" then 
								for intRow = 0 to ubound(aSites,2)
									Response.Write "<font class=""regTextBlack"">" & right("000" & aSites(0,intRow),3) & ") </font>"
									Response.Write "<a class=""reglinkMaroon"" href=""javascript:goChild('" & strEDIYear & "','" & right("000" & aSites(0,intRow),3) & "','','','','');"">" & aSites(1,intRow) & "</a>"
									Response.Write "<br/>"
								next
							end if 
							Response.Write "</td></tr></table>"
						case else
						Response.write " YOU changed the length to " & len(strMap)
					end select
					%>
					<br/> 
				</td>
			</tr>
		</table>
		<br/> 
		<input type="hidden" name="item" value="">	
	</form>
	<!-- #include virtual="/shared/page_footer.inc" -->
</body>
</html>
<%
	call close_adodb(rstData)
	'call close_adodb(conn_tables)
	'call close_adodb(conn)
	call close_adodb(conn)
end if
%>
