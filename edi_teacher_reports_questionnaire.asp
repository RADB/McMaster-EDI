<!-- #include virtual="/shared/security.asp" -->
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
	call open_adodb(conn, "MACEDI")
	'call open_adodb(conn, "DATA")
	set rstData = server.CreateObject("adodb.recordset")
	
	' put the EDI ID together
	strEDIID = ""
	for each item in Request.Form
		strEDIID = strEDIID & Request.form(item)
	next

	' get individual variables		
	strChild = Request.Form("frmChild")
	strClass = Request.Form("frmClass")
	strTeacher = session("id")
	
	if session("Language") = "English" then
		strTitle = "EDI Summary"
		strLink1 = "Class Selection for EDI Summary Report"
		strLink2 = "Child Selection for EDI Summary Report"
		strLink3 = "Class Selection"
		strAll = "All students for this teacher"
	else
		strTitle = "IMPDE Résumé"
		strLink1 = "La Sélection de classe pour le Rapport de Résumé d'IMPDE"
		strLink2 = "La Sélection d'enseignant pour le Rapport de Résumé d'IMPDE"
		strLink3 = "Sélection de classe"
		strAll = "Tous des élèves de l'enseignant"
	end if 
	
		' start class section
		if strClass <> "" then 
			strMap = "&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""edi_teacher_reports_questionnaire.asp"">Class Selection</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<font class=""boldtextblack"">Child Selection for EDI Summary Report</font>"
			strFrench = "&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""edi_teacher_reports_questionnaire.asp"">" & strLink3 & "</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<font class=""boldtextblack"">" & strLink2 & "</font>"
	
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
				strMap = "&nbsp;<font class=""regtextblack"">></font>&nbsp;<font class=""boldtextblack"">Class Selection for EDI Summary Report</font>"
				strFrench = "&nbsp;<font class=""regtextblack"">></font>&nbsp;<font class=""boldtextblack"">" & strLink1 & "</font>"
				' get the class data - this teachers classes
				'strSql = "SELECT DISTINCT c.intClassID FROM classes AS c RIGHT JOIN children AS ch ON c.intClassID = ch.intClassID WHERE c.intTeacherID = " & strEDIID & " ORDER BY c.intClassID"
				strSql = "SELECT DISTINCT c.intClassID FROM teachers AS t RIGHT JOIN (classes AS c RIGHT JOIN children AS ch ON c.intClassID = ch.intClassID) ON t.intTeacherID = c.intTeacherID WHERE t.strEmail = '" & strTeacher & "'"
				
				rstData.Open strSql, conn
				if not rstData.eof then 
					aClasses = rstData.getrows
				else
					strError = "<font class=""regtextred"">No class data on teacher - " & strEDIID & "</font>"
				end if
				
				' close the recordset
				rstData.close
			' end teacher section
			end if
		' end class section
		end if 
	' end child section
	'end if
	%>
	<form name="Children" method="POST" action="edi_teacher_reports_questionnaire.asp"> 
		<input type="hidden" name="frmSite" value="">
		<input type="hidden" name="frmSchool" value="">
		<input type="hidden" name="frmTeacher" value="">
		<input type="hidden" name="frmClass" value="">					
		<input type="hidden" name="frmChild" value="">
	</form>
	<form name="Screens" method="POST" action="edi_teacher_reports.asp" target="Reports"> 
		<input type="hidden" name="email" value="">
		<input type="hidden" name="Student" value="">
		<input type="hidden" name="rpt" value="">
		<input type="hidden" name="XML" value="">		
		<a class="reglinkMaroon" href="edi_teacher.asp"><%=strHome%></a><%=strFrench%>
		<table border="1" bordercolor="006600" cellpadding="0" cellspacing="0" width="760">
			<tr>
				<td>
					<table border="0" cellpadding="0" cellspacing="0" width="760" align="center">
						<tr>
							<td align="right" width="460">
								<font class="headerBlue"><%=strTitle%></font>
							</td>
							<td align="right">
								<input type="button" value="<%=strExit%>" name="Exit" onClick="javascript:window.location='edi_teacher.asp';">
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
							case 256 ' child
								Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""500"" align=""center"">"						
								Response.Write "<tr><td>"	
								if strError = "" then 
									for intRow = 0 to ubound(aChildren,2)
										Response.Write "<font class=""regTextBlack"">" & right("00" & aChildren(1,intRow),2) & ") </font>"
										if session("language") = "English" then 
											Response.Write "<a class=""reglinkMaroon"" href=""javascript:goEDIReport('" & strEDIID & right("00" & aChildren(1,intRow),2) & "');"">" & day(aChildren(5,intRow)) & "-" & monthname(datepart("m",aChildren(5,intRow)),true) & "-" & year(aChildren(5,intRow)) & " - " & aChildren(3,intRow) & "</a>"
										else
											Response.Write "<a class=""reglinkMaroon"" href=""javascript:goEDIReport('" & strEDIID & right("00" & aChildren(1,intRow),2) & "');"">" & day(aChildren(5,intRow)) & "-" & left(french_month(datepart("m",aChildren(5,intRow))),3) & "-" & year(aChildren(5,intRow)) & " - " & aChildren(3,intRow) & "</a>"
										end if 
										Response.Write "<br/>"
									next
								end if 
								Response.Write "</td></tr></table>"	
							case 120 ' class
								'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
								' get the class times
								'//////////////////////////////////////////////////////////////////////	
								set rstClass = server.CreateObject("adodb.recordset")
	
								' open all languages
								rstClass.Open "SELECT intClassID, " & session("language") & " FROM [LU Classes]", conn
								
								' store all languages in array
								aClass = rstClass.GetRows 
	
								' close and kill the langauges recordset
								call close_adodb(rstClass)
							
								Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""500"" align=""center"">"						
								Response.Write "<tr><td><img src=""images/greenbb.gif"">&nbsp;<a class=""reglinkMaroon"" href=""javascript:goTeacherClassEDIReport('0', '" & strTeacher & "');"">" & strAll & "</a>&nbsp;<img src=""images/greenbb.gif""><br /><br /></td><tr>"
								Response.Write "<tr><td>"	
								if strError = "" then 
									for intRow = 0 to ubound(aClasses,2)
										Response.Write "<font class=""regTextBlack"">" & right("0" & aClasses(0,intRow),9) & ") </font>"
										Response.Write "<a class=""reglinkMaroon"" href=""javascript:goChild('" & left(right("000" & aClasses(0,intRow),9),3) & "','" & mid(right("000" & aClasses(0,intRow),9),4,3) & "','" & mid(right("000" & aClasses(0,intRow),9),7,2) & "','" & right(right("000" & aClasses(0,intRow),9),1) & "','');"">"  & aClass(1,right(aClasses(0,intRow),1)) & "</a>"
										' show class description
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
%>
