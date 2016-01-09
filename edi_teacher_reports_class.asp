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
	
	if session("Language") = "English" then
		strTitle = "Class Summary "
		strLink1 = "Class Selection for Class Summary Report"
		strAll = "All classes for teacher"
	else
		strTitle = "Résumé de classe"
		'strLink1 = "La Sélection de classe pour le Rapport de Résumé de Classe"
		strLink1 = "La Sélection de la classe pour le Rapport Résumé d'IMDPE"
		strAll = "Toutes classes pour l'enseignant"
	end if 
	
	' get individual variables		
	'strChild = Request.Form("frmChild")
	strClass = Request.Form("frmClass")
	strTeacher = session("ID")
	
		if strTeacher <> "" then 
			strMap = "&nbsp;<font class=""regtextblack"">></font>&nbsp;<font class=""boldtextblack"">Class Selection for Class Summary Report</font>"
			strFrench = "&nbsp;<font class=""regtextblack"">></font>&nbsp;<font class=""boldtextblack"">" & strLink1 & "</font>"
			' get the class data - this teachers classes
			strSql = "SELECT DISTINCT c.intClassID FROM teachers AS t RIGHT JOIN (classes AS c RIGHT JOIN children AS ch ON c.intClassID = ch.intClassID) ON t.intTeacherID = c.intTeacherID WHERE t.strEmail = '" & strTeacher & "'"
			
			rstData.Open strSql, conn
			if not rstData.eof then 
				aClasses = rstData.getrows
			else
				strError = "<font class=""regtextred"">No class data on teacher - " & strEDIID & "</font>"
			end if
				
			' close the recordset
			rstData.close
		end if 
	%>
	<form name="Children" method="POST" action="edi_teacher_reports_class.asp"> 
		<input type="hidden" name="frmSite" value="">
		<input type="hidden" name="frmSchool" value="">
		<input type="hidden" name="frmTeacher" value="">
		<input type="hidden" name="frmClass" value="">					
		<input type="hidden" name="frmChild" value="">
	</form>
	<form name="Screens" method="POST" action="edi_teacher_reports.asp" target="Reports"> 
		<input type="hidden" name="classes" value="">
		<input type="hidden" name="email" value="">
		<input type="hidden" name="rpt" value="">
		<input type="hidden" name="XML" value="">		
		<a class="reglinkMaroon" href="edi_teacher.asp"><%=strHome%></a><%=strFrench%>
		<table border="1" bordercolor="006600" cellpadding="0" cellspacing="0" width="760">
			<tr>
				<td>
					<table border="0" cellpadding="0" cellspacing="0" width="760">
						<tr>
							<td align="right" width="460">
								<font class="headerBlue"><%=strTitle%></font>
							</td>
							<td align="right">
								<input type="button" value="<%=strExit%>" name="Exit" onClick="javascript:window.location='edi_teacher.asp';">
								&nbsp;
							</td>
						<tr><td colspan="2"><%="<br/>" & strError%></td></tr>
					</table>
					<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">						
						<tr>
							<td>
																
							</td>
						</tr>
					</table>
					
						<%
						select case len(strMap)
							case 122 ' class
								'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
								' get the class times
								'//////////////////////////////////////////////////////////////////////	
								set rstClass = server.CreateObject("adodb.recordset")
	
								' open all languages
								rstClass.Open "SELECT intClassID, " & session("Language") & " FROM [LU Classes]", conn
								
								' store all languages in array
								aClass = rstClass.GetRows 
	
								' close and kill the langauges recordset
								call close_adodb(rstClass)
							
								Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""500"" align=""center"">"						
								Response.Write "<tr><td><img src=""images/greenbb.gif"">&nbsp;<a class=""reglinkMaroon"" href=""javascript:goTeacherReport('0', '" & strTeacher & "');"">" & strAll & " " & strEdIID & "</a>&nbsp;<img src=""images/greenbb.gif""><br /><br /></td><tr>"
								Response.Write "<tr><td>"	
								if strError = "" then 
									for intRow = 0 to ubound(aClasses,2)
										Response.Write "<font class=""regTextBlack"">" & right("0" & aClasses(0,intRow),9) & ") </font>"
										Response.Write "<a class=""reglinkMaroon"" href=""javascript:goTeacherReport('" & right("0" & aClasses(0,intRow),9) & "', '" & strTeacher & "');"">"  & aClass(1,right(aClasses(0,intRow),1)) & "</a>"
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



sub buildSites()
	
end sub
%>
