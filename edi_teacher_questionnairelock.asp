<!-- #include virtual="/shared/security.asp" -->
<%
' if the user has not logged in they will not be able to see the page
on error resume next
if blnSecurity then 
	' open edi connection
	call open_adodb(conn, "MACEDI")
	'call open_adodb(conn, "DATA")
	'call open_adodb(conn_tables, "TABLES")
	dim aHeader(6)
	'Response.Write "<font class=""boldtextwhite"">" & Request.form & "</font>"
	
%>
<html>
<head>
	<!-- added UTF8 Encoding to get rid of funny characters -->
	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" /> 
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
	strChild = Request.Form("frmChild")	
	strClass = Request.Form("frmClass")
	strTeacher = Request.Form("frmTeacher")
	strSchool = Request.Form("frmSchool")
	strSite = Request.Form("frmSite")
	strEDIYear = Request.Form("frmEDIYear")
	strEdiID = strEDIYear & strSite & strSchool & strTeacher & strClass & strChild

	if Request.form("frmAction") = "lockClassList" then  
		strSql = "UPDATE demographics SET intStatus = 6, strLanguageCompleted = '" & session("language") & "' WHERE strEDIID = '" & strEDIID & "'"
		'response.write strSql
		' update the status to indicate no consent
		conn.execute strSql
	end if 
	 	
	if Request.form("frmAction") = "lock" then  

	'	conn.execute "UPDATE children SET chkCompleted = true, dtmDate = '" & date & "' WHERE strEDIID = '" & strEDIID & "'"
		select case Request.Form("CurrentSection")
		case "_Demographics"
				strSQL = "UPDATE demographics SET "
				strSql2 = "UPDATE children SET "

				if not (Request.Form("DOBday") = "-1" or Request.Form("DOBmonth") = "-1" or Request.Form("DOByear") = "-1") then 
					strSql2 = strSql2 & "dtmDob = '" & Request.Form("DOBday") & "/" & monthname(Request.Form("DOBmonth")) & "/" & Request.Form("DOByear") & "', "
				end if 
                

				for each item in Request.Form 
					if NOT (item = "frmAction" OR item = "frmSection" OR item = "frmEDIYear" OR item = "frmSite" OR item ="frmSchool" OR item ="frmTeacher" OR item ="frmClass" OR item = "frmChild" OR item = "frmNextChild" OR item = "intSex" OR item = "strPostal" OR item = "DOBday" OR item = "DOByear" OR item = "DOBmonth" OR item = "btnSave" OR item = "hdnLock" OR item = "hdnCheckBoxes" OR item = "Student" OR item = "classes" OR item = "email" OR  item = "hdnRadioButtons" OR item ="rpt" OR item ="XML" or item="CurrentSection") then 
						blnUpdate = true
						if left(item ,3) = "str" then
							strSql = strSql & item & " = " & checknull(Request.Form(item)) & ","
						else
							strSql = strSql & item & " = " & checkValue(Request.Form(item)) & ","
						end if 
					elseif NOT (item = "frmAction" OR item = "frmSection" OR item = "frmEDIYear" OR item = "frmSite" OR item ="frmSchool" OR item ="frmTeacher" OR item ="frmClass" OR item = "frmChild" OR item = "frmNextChild" OR item = "DOBday" OR item = "DOByear" OR item = "DOBmonth" OR item = "btnSave" OR item = "hdnLock" OR item = "hdnCheckBoxes" OR item = "Student" OR item = "classes" OR item = "email" OR  item = "hdnRadioButtons" OR item ="rpt" OR item ="XML" or item="CurrentSection") then 
						blnUpdate2 = true
						if left(item ,3) = "str" then
							strSql2 = strSql2 & item & " = " & checknull(Request.Form(item)) & ","
						else
							strSql2 = strSql2 & item & " = " & checknumber(Request.Form(item)) & ","
						end if
					end if 
				next
				
				' deal with the checkboxes
		        for each item in split(request.form("hdnCheckBoxes"),",")
		            'response.write item
		            if request.form(item).count = 0 then 
		                strSql = strSql & item & " = 0,"
		            end if 
		        next 
				
				' remove the last comma
				strSql = left(strSql,len(strSql)-1) & " WHERE strEDIID = '" & strEDIID & "'" 'Request.Form("frmSite") & Request.Form("frmSchool") & Request.Form("frmTeacher") & Request.Form("frmClass") & Request.Form("frmChild") & "'"
				strSql2 = left(strSql2,len(strSql2)-1) & " WHERE strEDIID = '" & strEDIID & "'" 'Request.Form("frmSite") & Request.Form("frmSchool") & Request.Form("frmTeacher") & Request.Form("frmClass") & Request.Form("frmChild") & "'"
				
				if blnUpdate then 
					'Response.Write strSql
					conn.execute strSql 
				end if 
							
				' if no errors then update other
				if conn.errors.count = 0 then 	
					if blnUpdate2 then 
                        'Response.Write strSql2
						' updates the children table
						conn.execute strSql2
					end if 
				else
				   response.Write  "<font class=""regtextred"">Error Number : " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/></font>"
				end if 
			case "A"
				if cint(Request.Form("days")) > -1 then 
			        if cint(Request.Form("days2")) > -1 then 
    				    strSQL =  "UPDATE sectionA SET intQ1 = " & Request.Form("days") & "." & Request.Form("days2") & ","
    				else
    				    strSQL = "UPDATE sectionA SET intQ1 = " & Request.Form("days") & ".0,"
    				end if 
    				blnUpdate = True
				else
				    strSQL = "UPDATE sectionA SET "
				end if 
				
				for each item in Request.Form 
					if NOT (item = "frmAction" OR  item = "hdnRadioButtons" OR item = "frmSection" OR item = "frmEDIYear" OR item = "frmSite" OR item ="frmSchool" OR item ="frmTeacher" OR item ="frmClass" OR item = "frmChild" OR item = "frmNextChild" OR item = "days2" OR item = "days" OR item = "btnSave" OR item = "hdnLock" OR item = "hdnCheckBoxes" OR item = "Student" OR item = "classes" OR item = "email" OR item ="rpt" OR item ="XML" or item="CurrentSection") then 
						' added May 26 2004
						' removed feb 2 2006 - see below
						blnUpdate = true
						if left(item ,3) = "str" then
							strSql = strSql & item & " = " & checknull(Request.Form(item)) & ","
						else
							strSql = strSql & item & " = " & Request.Form(item) & ","
						end if 
					end if 
				next
			
			case "B"
				strSQL = "UPDATE sectionB SET " 
				for each item in Request.Form 
					if NOT (item = "frmAction" OR  item = "hdnRadioButtons" OR item = "frmSection" OR item = "frmEDIYear" OR item = "frmSite" OR item ="frmSchool" OR item ="frmTeacher" OR item ="frmClass" OR item = "frmChild" OR item = "frmNextChild" OR item = "btnSave" OR item = "hdnLock" OR item = "hdnCheckBoxes" OR item = "Student" OR item = "classes" OR item = "email" OR item ="rpt" OR item ="XML" or item="CurrentSection") then 
						blnUpdate = true
						if left(item ,3) = "str" then
							strSql = strSql & item & " = " & checknull(Request.Form(item)) & ","
						else
							strSql = strSql & item & " = " & Request.Form(item) & ","
						end if 
					end if 
				next

			case "C"
				strSQL = "UPDATE sectionC SET " 
				for each item in Request.Form 
					if NOT (item = "frmAction" OR  item = "hdnRadioButtons" OR item = "frmSection" OR item = "frmEDIYear" OR item = "frmSite" OR item ="frmSchool" OR item ="frmTeacher" OR item ="frmClass" OR item = "frmChild" OR item = "frmNextChild" OR item = "btnSave" OR item = "hdnLock" OR item = "hdnCheckBoxes" OR item = "Student" OR item = "classes" OR item = "email" OR item ="rpt" OR item ="XML" or item="CurrentSection") then 
						blnUpdate = true
						if left(item ,3) = "str" then
							strSql = strSql & item & " = " & checknull(Request.Form(item)) & ","
						else
							strSql = strSql & item & " = " & Request.Form(item) & ","
						end if 
					end if 
				next

			case "D"                
				strSQL = "UPDATE sectionD SET " 
	
                for each item in Request.Form 
					if NOT (item = "frmAction" OR  item = "hdnRadioButtons" OR item = "frmSection" OR item = "frmEDIYear" OR item = "frmSite" OR item ="frmSchool" OR item ="frmTeacher" OR item ="frmClass" OR item = "frmChild" OR item = "frmNextChild" OR item = "btnSave" OR item = "hdnLock" OR item = "hdnCheckBoxes" OR item = "Student" OR item = "classes" OR item = "email" OR item ="rpt" OR item ="XML" or item="CurrentSection") then 
						blnUpdate = true						
						if left(item ,3) = "str" then
							strSql = strSql & item & " = " & checknull(Request.Form(item)) & ","
						else
							strSql = strSql & item & " = " & checkValue(Request.Form(item)) & ","
						end if 
					end if 
				next			             
			case "E"
				strSQL = "UPDATE sectionE SET " 
				for each item in Request.Form 
					if NOT (item="frmEDIYear" OR  item = "hdnRadioButtons" OR item="intQ2a" OR item="intQ2b" OR item="intQ2c" OR item="intQ2d" OR item="intQ2e" OR item="intQ2f" OR item="intQ2g" OR item="intQ2h" OR item = "frmAction" OR item = "frmSection" OR item = "frmSite" OR item ="frmSchool" OR item ="frmTeacher" OR item ="frmClass" OR item = "frmChild" OR item = "frmNextChild" OR item = "btnSave" OR item = "hdnLock" OR item = "hdnCheckBoxes" OR item = "Student" OR item = "classes" OR item = "email" OR item ="rpt" OR item ="XML" or item="CurrentSection") then 
						blnUpdate = true
						if left(item ,3) = "str" then
							strSql = strSql & "[" & item & "] = " & checknull(Request.Form(item)) & ","
						else
							strSql = strSql & "[" & item & "] = " & Request.Form(item) & ","
						end if 
					end if 
				next
		end select
		
		if blnUpdate AND Request.Form("CurrentSection") <> "_Demographics" then
			' deal with nullified radio buttons and checkboxes
			call splitValues(split(request.form("hdnCheckBoxes"),","), "0")			
			call splitValues(split(request.form("hdnRadioButtons"),","), "")
			' remove the last comma
			'strSql = left(strSql,len(strSql)-1) & " WHERE strEDIID = '" & Request.Form("frmSite") & Request.Form("frmSchool") & Request.Form("frmTeacher") & Request.Form("frmClass") & Request.Form("frmChild") & "'" 
			strSql = left(strSql,len(strSql)-1) & " WHERE strEDIID = '" & strEDIID & "'"
			'Response.Write strSql & "<br />"
			'response.write request.form
			
			conn.execute strSql 
		end if 
		
		if conn.errors.count > 0 and blnUpdate <> False then 
			strError = "<font class=""regtextred"">Error Number : " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/></font>"
		end if 	
	end if 
	

	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	' Check for child ID 
	'	- if child id is there then show the child
	'			- build EDI string
	'	- else check for class ID
	'		- if class ID then show its students
	'		- else check for teacher ID
	'		-  show teachers classes!!!
	'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	
	' start child section
	if strChild <> "" then 
		strMap = "&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strSite & "','" & strSchool & "','','','');"">"&lblTeacher&"</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strSite & "','" & strSchool & "','" & strTeacher  & "','','');"">Class</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strSite & "','" & strSchool & "','" & strTeacher  & "','" & strClass  & "','');"">Student</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<font class=""boldtextblack"">EDI Questionnaire</font>"
		strFrench = "&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strSite & "','" & strSchool & "','','','');"">"&lblTeacher&"</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strSite & "','" & strSchool & "','" & strTeacher  & "','','');"">classe</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strSite & "','" & strSchool & "','" & strTeacher  & "','" & strClass  & "','');"">élève</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<font class=""boldtextblack"">IMDPE Questionnaire</font>"
	
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
			strMap = "&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strSite & "','" & strSchool & "','','','');"">Teacher</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strSite & "','" & strSchool & "','" & strTeacher  & "','','');"">Class</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<font class=""boldtextblack"">Student</font>"
			strFrench = "&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strSite & "','" & strSchool & "','','','');"">enseignant</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strSite & "','" & strSchool & "','" & strTeacher  & "','','');"">classe</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<font class=""boldtextblack"">étudiant</font>"
	
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
				strMap = "&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strSite & "','" & strSchool & "','','','');"">"&lblTeacher&"</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<font class=""boldtextblack"">Class</font>"
				strFrench = "&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strSite & "','" & strSchool & "','','','');"">"&lblTeacher&"</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<font class=""boldtextblack"">classe</font>"
	
				' get the class data - this teachers classes
				strSql = "SELECT * FROM classes WHERE intTeacherID = " & strEDIID & " ORDER BY intClassID"
				
				rstData.Open strSql, conn
				if not rstData.eof then 
					aClasses = rstData.getrows
				else
					strError = "<font class=""regtextred"">No class data on teacher - " & strEDIID & "</font>"
				end if
		
				' close the recordset
				rstData.close
			else
				strMap = "&nbsp;<font class=""regtextblack"">></font>&nbsp;<font class=""boldtextblack"">"&lblTeacher&"</font>"
				strFrench = "&nbsp;<font class=""regtextblack"">></font>&nbsp;<font class=""boldtextblack"">"&lblTeacher&"</font>"
				
				strSql = "SELECT * FROM teachers WHERE strEmail = '" & session("id") & "'"
				
				' get the teacher ID
				rstData.Open strSql, conn
				if not rstData.eof then 
					aTeachers = rstData.getrows
					strEDIID = right("00" & aTeachers(0,0),8)	
				else
					strError = "<font class=""regtextred"">No teacher data!!</font>"
				end if
				
				' close the recordset
				rstData.close
				
			' end teacher section
			end if 
		' end class section
		end if 
	' end child section
	end if
	%>
	<form name="Children" method="POST" action="edi_teacher_class.asp"> 
		<input type="hidden" name="Student" value="">
		<input type="hidden" name="classes" value="">
		<input type="hidden" name="email" value="">
		<input type="hidden" name="rpt" value="">
		<input type="hidden" name="XML" value="">	
		<input type="hidden" name="frmEDIYear" value="">					
		<input type="hidden" name="frmSite" value="">					
		<input type="hidden" name="frmSchool" value="">
		<input type="hidden" name="frmTeacher" value="">
		<input type="hidden" name="frmClass" value="">					
		<input type="hidden" name="frmChild" value="">
		<input type="hidden" name="frmNextChild" value="">
		<input type="hidden" name="frmAction" value="">
		<input type="hidden" name="frmSection" value="">	
	<!--</form>
	<form name="Screens" method="POST" action="edi_teacher_questionnaire.asp"> -->
		<a class="reglinkMaroon" href="edi_teacher.asp"><%=strHome%></a>
		<%
		if session("Language") = "English" then 
			Response.write strMap
		else
			Response.Write strFrench
		end if 
		%>	
		<table border="1" bordercolor="006600" cellpadding="0" cellspacing="0" width="760">
			<tr>
				<td>
					<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">
						<tr>
							<td align="right" width="520"><font class="headerBlue"><%=lblEDI%> Questionnaire <%=strLock%>(<%=strEDIID%>)</font></td>
							<td align="right">
								<input type="button" value="<%=lblCancel%>" name="Cancel" onClick="javascript:goChild(<%="'" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "'"%>);">
								&nbsp;
							</td>	
						</tr>
						<tr><td colspan="2"><%="<br/>" & strError%></td></tr>
					</table>
<%
					select case len(strMap)
						case 504,498 ' Questionnaire
							' aChild
							if strError = "" then 
								' summary 
								Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""750"" align=""left"">"														
									' status
									Response.Write "<tr>"	
									Response.Write "<td align=""right""><font class=""boldTextBlack"">" & lblStatus & ":&nbsp;</font></td>"
									Response.Write "<td align=""left""><font class=""regTextBlack"">"
										if aStudent(7,0) = true  then 
											Response.Write "<font class=""regTextGreen"">&nbsp;&nbsp;" & lblComplete
											blnLock = true
										else
											Response.Write "<font class=""regTextRed"">&nbsp;&nbsp;" & lblIncomplete
											blnLock = false
										end if 
									Response.write "</font></td>"
									Response.Write "<td width = ""350"" rowspan=""3"" align=""left"">&nbsp;&nbsp;"
									if not blnLock then 
										Response.Write "<a href=""javascript:confirm_Lock('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','lock','" & strConfirmLanguage & "','" & strConfirmLanguage2 & "');"" class=""bigLinkBlue""><img src=""images/lock.gif"" border=""0"">" & lblFinished & "</a>"
									end if 
									
									' determine the screen that the user is on
									if Request.Form("frmSection").Count = 0 OR Request.Form("frmSection") = "_Demographics" OR Request.Form("frmSection") = "" then 
										strCurrent = "_Demographics"
									else
										strCurrent = Request.Form("frmSection")
									end if 
									

									set rstChildren = server.CreateObject("adodb.recordset")
									
									
									Response.Write "</td>"
									Response.Write "</tr>"
									
									' spacer
									Response.Write "<tr><td colspan=""2"">&nbsp;</td></tr>"


									'********************************************									
									' Changed Oct 31, 2005
									' Changed Nov 19, 2006 - Added Status return
									'********************************************
									' set the SQL query to get then next child in the class
									'strSql = "SELECT Completed, intStatus FROM Demographics_Complete_" & session("province") & " WHERE strEDIID = '" & strEDIID & "'" 
                                    strSql = "Exec completionCheck '" & strEDIID & "'," & session("province")
                                   
									' open the recordset
									rstChildren.Open strSql, conn
																		
									dim blnStatus

									if not isnull(rstChildren("intStatus")) then                                    
										if rstChildren("intStatus") > 1 then 
											blnStatus = True
										else
											' complete questionnaire
											blnStatus = False
										end if 
									else
										' complete questionnaire
										blnStatus = False
									end if 									
									

									' Demographics
									Response.Write "<tr>"	
									Response.Write "<td align=""right""><font class=""boldTextBlack"">" & strDemographics& ":&nbsp;</font></td>"
									if blnStatus then 
										Response.Write "<td align=""left""><font class=""lrgregTextBlack"">&nbsp;&nbsp;<a " & StatusLink("","Complete") & "</font></td>"
									else
										Response.Write "<td align=""left""><font class=""lrgregTextBlack"">&nbsp;&nbsp;<a " & StatusLink("",rstChildren("Demographics")) & "</font></td>"
									end if 
									Response.Write "</tr>"

									'rstChildren.close
									' set the SQL query to get then next child in the class
									'strSql = "SELECT Completed FROM SectionA_Complete WHERE strEDIID = '" & strEDIID & "'" 

									' open the recordset
									'rstChildren.Open strSql, conn

                                    ' - not sure what this does!!!
									'blnStatus = checkStatus(rstChildren(0))

									' Section A
									Response.Write "<tr>"	
									Response.Write "<td align=""right""><font class=""boldTextBlack"">Section A:&nbsp;</font></td>"
									if blnStatus then 
										Response.Write "<td align=""left"">&nbsp;&nbsp;" & StatusLink("A","Complete") & "</td>"
									else
										Response.Write "<td align=""left"">&nbsp;&nbsp;" & StatusLink("A",rstChildren("SectionA")) & "</td>"
									end if 
									'Response.Write "<td width = ""350"" rowspan=""3"" align=""left"">&nbsp;&nbsp;<a href=""javascript:goTeacherClassReport('" & right("0" & strSite & strSchool & strTeacher & strClass,9) & "', '" & session("id") & "');"" class=""bigLinkBlue""><img src=""images/download.gif"" border=""0""> " & lblClassSummary & "</a></td>"
									Response.Write "</tr>"
									
									'rstChildren.close
									' set the SQL query to get then next child in the class
									'strSql = "SELECT Completed FROM SectionB_Complete WHERE strEDIID = '" & strEDIID & "'" 

									' open the recordset
									'rstChildren.Open strSql, conn
									
									' Section B
									Response.Write "<tr>"	
									Response.Write "<td align=""right""><font class=""boldTextBlack"">Section B:&nbsp;</font></td>"
									if blnStatus then
										Response.Write "<td align=""left""><font class=""lrgregTextBlack"">&nbsp;&nbsp;" & StatusLink("B","Complete") & "</font></td>"
									else
										Response.Write "<td align=""left""><font class=""lrgregTextBlack"">&nbsp;&nbsp;" & StatusLink("B",rstChildren("SectionB")) & "</font></td>"
									end if
									Response.Write "</tr>"

									'rstChildren.close
									' set the SQL query to get then next child in the class
									'strSql = "SELECT Completed FROM SectionC_Complete WHERE strEDIID = '" & strEDIID & "'" 

									' open the recordset
									'rstChildren.Open strSql, conn
									
										
									' Section C
									Response.Write "<tr>"	
									Response.Write "<td align=""right""><font class=""boldTextBlack"">Section C:&nbsp;</font></td>"
									if blnStatus then
										Response.Write "<td align=""left""><font class=""lrgregTextBlack"">&nbsp;&nbsp;" & StatusLink("C","Complete") & "</font></td>"
									else
										Response.Write "<td align=""left""><font class=""lrgregTextBlack"">&nbsp;&nbsp;" & StatusLink("C",rstChildren("SectionC")) & "</font></td>"
									end if 
									'Response.Write "<td width = ""350"" rowspan=""2"">&nbsp;</td>"
									Response.Write "</tr>"

									'rstChildren.close
									' set the SQL query to get then next child in the class ' _" & session("province") & "
									'strSql = "SELECT Completed FROM SectionD_Complete WHERE strEDIID = '" & strEDIID & "'" 

									' open the recordset
									'rstChildren.Open strSql, conn
									
									' Section D
									Response.Write "<tr>"	
									Response.Write "<td align=""right""><font class=""boldTextBlack"">Section D:&nbsp;</font></td>"
									if blnStatus then
										Response.Write "<td align=""left""><font class=""lrgregTextBlack"">&nbsp;&nbsp;" & StatusLink("D","Complete") & "</font></td>"
									else
										Response.Write "<td align=""left""><font class=""lrgregTextBlack"">&nbsp;&nbsp;" & StatusLink("D",rstChildren("SectionD")) & "</font></td>"
									end if 
									'Response.Write "<td width = ""350"" rowspan=""3"" align=""left"">&nbsp;&nbsp;&nbsp;<a href=""javascript:goTeacherEDIReport('" & strEDIID & "');"" class=""bigLinkBlue""><img border=""0"" src=""images/details.gif"">" & lblSummary & "</a></td>"
									Response.Write "</tr>"
						
									'rstChildren.close
									' set the SQL query to get then next child in the class '_" & session("province") & "
									'strSql = "SELECT Completed FROM SectionE_Complete WHERE strEDIID = '" & strEDIID & "'" 

									' open the recordset
									'rstChildren.Open strSql, conn
									
									' Section E
									Response.Write "<tr>"	
									Response.Write "<td align=""right""><font class=""boldTextBlack"">Section E:&nbsp;</font></td>"
									if blnStatus then
										Response.Write "<td align=""left""><font class=""lrgregTextBlack"">&nbsp;&nbsp;" & StatusLink("E","Complete") & "</font></td>"
									else
										Response.Write "<td align=""left""><font class=""lrgregTextBlack"">&nbsp;&nbsp;" & StatusLink("E",rstChildren("SectionE")) & "</font></td>"
									end if 
									'Response.Write "<td width = ""350"" rowspan=""2"">&nbsp;</td>"
									Response.Write "</tr>"

									' close and kill the children recordset
									call close_adodb(rstChildren)
									
									' spacer 
									Response.Write "<tr><td><br /></td></tr>"
								Response.Write "</table>"	
							' end first column	
							Response.Write "</td></tr>"
						end if
					end select
%>			

					<br/>
				</td>
			</tr>
		</table>
		<br/> 
		<input type="hidden" name="hdnLock" value="<%=blnLock%>">	
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

function StatusLink(Section,Status)
	if Status= "Incomplete" then 
		StatusLink = "<a class=""reglinkMaroon"" href=""javascript:goChildSection('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher  & "','" & strClass  & "','" & strChild & "','" & Section & "');"">" & strIncomplete & "</a>"
	else
		StatusLink = "<font class=""regTextGreen"">" & strComplete & "<font>"
	end if 
End Function

function splitValues(items, valueIfNull)
    for each item in items
		'response.write item
		if request.form(item).count = 0 then 
			strSql = strSql & item & " = " & checkValue(valueIfNull) & ","
		end if 
	next	
end function
%>