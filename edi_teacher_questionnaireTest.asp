<!-- #include virtual="/shared/security.asp" -->

<%
' if the user has not logged in they will not be able to see the page
'on error resume next
if blnSecurity then 
	' open edi connection
	call open_adodb(conn, "MACEDI")
	'call open_adodb(conn, "DATA")
	'call open_adodb(conn_tables, "TABLES")
	
	dim aHeader(6)
	dim strSql
	dim strCheckBoxes
	'Response.Write "<font class=""boldtextwhite"">" & Request.form & "</font>"
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
	
	
	if Request.Form("frmAction").Count > 0 AND Request.form("hdnLock") = "False" then 
		
		select case Request.Form("frmAction")
			case "_Demographics"				
				strSQL = "UPDATE demographics SET "
				strSql2 = "UPDATE children SET "

				if not (Request.Form("DOBday") = "-1" or Request.Form("DOBmonth") = "-1" or Request.Form("DOByear") = "-1") then 
					strSql2 = strSql2 & "dtmDob = '" & Request.Form("DOBday") & "/" & monthname(Request.Form("DOBmonth")) & "/" & Request.Form("DOByear") & "', "
				end if 
                

				for each item in Request.Form 
					if NOT (item = "frmEDIYear" OR item = "frmAction" OR item = "frmSection" OR item = "frmSite" OR item ="frmSchool" OR item ="frmTeacher" OR item ="frmClass" OR item = "frmChild" OR item = "frmNextChild" OR item = "intSex" OR item = "strPostal" OR item = "DOBday" OR item = "DOByear" OR item = "DOBmonth" OR item = "btnSave" OR item = "hdnLock" OR item = "hdnCheckBoxes" OR item = "Student" OR item = "classes" OR item = "email" OR item ="rpt" OR item ="XML" or item="CurrentSection") then 
						blnUpdate = true
						if left(item ,3) = "str" then
							strSql = strSql & item & " = " & checknull(Request.Form(item)) & ","
						else
							'response.write item & Request.Form(item)
							strSql = strSql & item & " = " & checkValue(Request.Form(item)) & ","
						end if 
					elseif NOT (item = "frmEDIYear" OR item = "frmAction" OR item = "frmSection" OR item = "frmSite" OR item ="frmSchool" OR item ="frmTeacher" OR item ="frmClass" OR item = "frmChild" OR item = "frmNextChild" OR item = "DOBday" OR item = "DOByear" OR item = "DOBmonth" OR item = "btnSave" OR item = "hdnLock" OR item = "hdnCheckBoxes" OR item = "Student" OR item = "classes" OR item = "email" OR item ="rpt" OR item ="XML" or item="CurrentSection") then 
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
				'strSql = left(strSql,len(strSql)-1) & " WHERE strEDIID = '" & Request.Form("frmSite") & Request.Form("frmSchool") & Request.Form("frmTeacher") & Request.Form("frmClass") & Request.Form("frmChild") & "'"
				strSql = left(strSql,len(strSql)-1) & " WHERE strEDIID = '" & strEDIID & "'"
				'strSql2 = left(strSql2,len(strSql2)-1) & " WHERE strEDIID = '" & Request.Form("frmSite") & Request.Form("frmSchool") & Request.Form("frmTeacher") & Request.Form("frmClass") & Request.Form("frmChild") & "'"
				strSql2 = left(strSql2,len(strSql2)-1) & " WHERE strEDIID = '" & strEDIID & "'"
										
				if blnUpdate then 
					'Response.Write strSql
					conn.execute strSql 
				end if 
							
				' if no errors then update other
				if conn.errors.count = 0 then 	
					if blnUpdate2 then 
						' updates the children table
						'Response.Write strSql2
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
					if NOT (item = "frmEDIYear" OR item = "frmAction" OR item = "frmSection" OR item = "frmSite" OR item ="frmSchool" OR item ="frmTeacher" OR item ="frmClass" OR item = "frmChild" OR item = "frmNextChild" OR item = "days2" OR item = "days" OR item = "btnSave" OR item = "hdnLock" OR item = "hdnCheckBoxes" OR item = "Student" OR item = "classes" OR item = "email" OR item ="rpt" OR item ="XML" or item="CurrentSection") then 
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
					if NOT (item = "frmEDIYear" OR item = "frmAction" OR item = "frmSection" OR item = "frmSite" OR item ="frmSchool" OR item ="frmTeacher" OR item ="frmClass" OR item = "frmChild" OR item = "frmNextChild" OR item = "btnSave" OR item = "hdnLock" OR item = "hdnCheckBoxes" OR item = "Student" OR item = "classes" OR item = "email" OR item ="rpt" OR item ="XML" or item="CurrentSection") then 
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
					if NOT (item = "frmEDIYear" OR item = "frmAction" OR item = "frmSection" OR item = "frmSite" OR item ="frmSchool" OR item ="frmTeacher" OR item ="frmClass" OR item = "frmChild" OR item = "frmNextChild" OR item = "btnSave" OR item = "hdnLock" OR item = "hdnCheckBoxes" OR item = "Student" OR item = "classes" OR item = "email" OR item ="rpt" OR item ="XML" or item="CurrentSection") then 
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
					if NOT (item = "frmEDIYear" OR item = "frmAction" OR item = "frmSection" OR item = "frmSite" OR item ="frmSchool" OR item ="frmTeacher" OR item ="frmClass" OR item = "frmChild" OR item = "frmNextChild" OR item = "btnSave" OR item = "hdnLock" OR item = "hdnCheckBoxes" OR item = "Student" OR item = "classes" OR item = "email" OR item ="rpt" OR item ="XML" or item="CurrentSection") then 
						blnUpdate = true						
						if left(item ,3) = "str" then
							strSql = strSql & item & " = " & checknull(Request.Form(item)) & ","
						else
							strSql = strSql & item & " = " & checkValue(Request.Form(item)) & ","
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
				
			case "E"
				strSQL = "UPDATE sectionE SET " 
				for each item in Request.Form 
					if NOT (item = "frmEDIYear" OR item="intQ2a" OR item="intQ2b" OR item="intQ2c" OR item="intQ2d" OR item="intQ2e" OR item="intQ2f" OR item="intQ2g" OR item="intQ2h" OR item = "frmAction" OR item = "frmSection" OR item = "frmSite" OR item ="frmSchool" OR item ="frmTeacher" OR item ="frmClass" OR item = "frmChild" OR item = "frmNextChild" OR item = "btnSave" OR item = "hdnLock"  OR item = "hdnCheckBoxes" OR item = "Student" OR item = "classes" OR item = "email" OR item ="rpt" OR item ="XML" or item="CurrentSection") then 
						blnUpdate = true
						if left(item ,3) = "str" then
							strSql = strSql & "[" & item & "] = " & checknull(Request.Form(item)) & ","
						else
							strSql = strSql & "[" & item & "] = " & Request.Form(item) & ","
						end if 
					end if 
				next
				
				if request.Form("intQ2a") = "on" then
				    strSql = strSql & "intQ2a=1,"
				else
				    strSql = strSql & "intQ2a=0,"
				end if 
				
				if request.Form("intQ2b") = "on" then
				    strSql = strSql & "intQ2b=1,"
				else
				    strSql = strSql & "intQ2b=0,"
				end if 
				
				if request.Form("intQ2c") = "on" then
				    strSql = strSql & "intQ2c=1,"
				else
				    strSql = strSql & "intQ2c=0,"
				end if 
				
				if request.Form("intQ2d") = "on" then
				    strSql = strSql & "intQ2d=1,"
				else
				    strSql = strSql & "intQ2d=0,"
				end if 
				
				if request.Form("intQ2e") = "on" then
				    strSql = strSql & "intQ2e=1,"
				else
				    strSql = strSql & "intQ2e=0,"
				end if 
				
				if request.Form("intQ2f") = "on" then
				    strSql = strSql & "intQ2f=1,"
				else
				    strSql = strSql & "intQ2f=0,"
				end if 
				
				if request.Form("intQ2g") = "on" then
				    strSql = strSql & "intQ2g=1,"
				else
				    strSql = strSql & "intQ2g=0,"
				end if 
				
				if request.Form("intQ2h") = "on" then
				    strSql = strSql & "intQ2h=1,"
				else
				    strSql = strSql & "intQ2h=0,"
				end if 					
		end select
		
		if blnUpdate AND Request.Form("frmAction") <> "_Demographics" then
			' remove the last comma
			'strSql = left(strSql,len(strSql)-1) & " WHERE strEDIID = '" & Request.Form("frmSite") & Request.Form("frmSchool") & Request.Form("frmTeacher") & Request.Form("frmClass") & Request.Form("frmChild") & "'" 
			strSql = left(strSql,len(strSql)-1) & " WHERE strEDIID = '" & strEDIID & "'"
			'Response.Write strSql 
			conn.execute strSql 
		end if 
		
		if conn.errors.count > 0 and blnUpdate <> False then 
			strError = "<font class=""regtextred"">Error Number : " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/></font>"
		end if 
	end if 
%>
<html>
<head> 
	<!-- added UTF8 Encoding to get rid of funny characters -->
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" /> 
    <title>EDI Teacher Questionnaire</title>
	<!-- Start CSS files-->
		<link rel="stylesheet" type="text/css" href="Styles/edi.css" />
	<!-- End CSS files -->
	<script language="javascript" type="text/javascript" src="js/form.js"></script>
	<script language="javascript" type="text/javascript" src="js/window.js"></script>
</head>
<body>
	<!-- #include virtual="/shared/page_header.inc" -->
	<%
	dim aSites, aSchools, aTeachers, aClasses, aChildren, aChild, aStudent
	dim strTable, childDOb
	
	set rstData = server.CreateObject("adodb.recordset")
	
	'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
	' get the languages for the drop down box
	'//////////////////////////////////////////////////////////////////////	
	set rstLanguages = server.CreateObject("adodb.recordset")
	
	' open all languages
	rstLanguages.Open "SELECT LID, english, french FROM [LU Languages] ORDER BY [Sequence]", conn

	' store all languages in array
	aLanguages = rstLanguages.GetRows 	

	' close and kill the langauges recordset
	call close_adodb(rstLanguages)
	
	 	
'	if Request.form("frmAction") = "lock" then  
'		conn.execute "UPDATE children SET chkCompleted = true, dtmDate = '" & date & "' WHERE strEDIID = '" & strEDIID & "'"	
	if Request.form("frmAction") = "unlock" then
		conn.execute "UPDATE children SET chkCompleted = false, dtmDate = null WHERE strEDIID = '" & strEDIID & "'"			
	end if 
	
	'rebuild the EDIID for the new EDIID
	if Request.Form("frmNextChild") <> "" then 
		strChild = Request.Form("frmNextChild")
	else
		strChild = Request.Form("frmChild")
	end if
	strEdiID = strEDIYear & strSite & strSchool & strTeacher & strClass & strChild

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
        strFrench = "&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strEDIYear & "','" & strSite & "','" & strSchool & "','','','');"">" & lblTeacher &"</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher  & "','','');"">Classe</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher  & "','" & strClass  & "','');"">Élève</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<font class=""boldtextblack"">IMDPE Questionnaire</font>"
        strMap = "&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strEDIYear & "','" & strSite & "','" & strSchool & "','','','');"">" & lblTeacher &"</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher  & "','','');"">Class</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher  & "','" & strClass  & "','');"">Student</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<font class=""boldtextblack"">EDI Questionnaire</font>" 	
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
			strMap = "&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strEDIYear & "','" & strSite & "','" & strSchool & "','','','');"">"&lblTeacher&"</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher  & "','','');"">Class</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<font class=""boldtextblack"">Student</font>"
			strFrench = "&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strEDIYear & "','" & strSite & "','" & strSchool & "','','','');"">"&lblTeacher&"</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher  & "','','');"">Classe</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<font class=""boldtextblack"">Élève</font>"
	
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
				strMap = "&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strEDIYear & "','" & strSite & "','" & strSchool & "','','','');"">"&lblTeacher&"</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<font class=""boldtextblack"">Class</font>"
				strFrench = "&nbsp;<font class=""regtextblack"">></font>&nbsp;<a class=""reglinkMaroon"" href=""javascript:goChild('" & strEDIYear & "','" & strSite & "','" & strSchool & "','','','');"">"&lblTeacher&"</a>&nbsp;<font class=""regtextblack"">></font>&nbsp;<font class=""boldtextblack"">Classe</font>"
	
				' get the class data - this teachers classes
				strSql = "SELECT * FROM classes WHERE intTeacherID = " & strSite & strSchool & strTeacher & " ORDER BY intClassID"
				
				rstData.Open strSql, conn
				if not rstData.eof then 
					aClasses = rstData.getrows
				else
					strError = "<font class=""regtextred"">No class data on teacher - " & strSite & strSchool & strTeacher  & "</font>"
				end if
		
				' close the recordset
				rstData.close
			else
				strMap = "&nbsp;<font class=""regtextblack"">></font>&nbsp;<font class=""boldtextblack"">Teacher</font>"
				strFrench = "&nbsp;<font class=""regtextblack"">></font>&nbsp;<font class=""boldtextblack"">enseignant</font>"
				
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
	<form name="Children" method="post" action="edi_teacher_questionnaire.asp"> 
		<input type="hidden" name="Student" value="" />
		<input type="hidden" name="classes" value="" />
		<input type="hidden" name="email" value="" />
		<input type="hidden" name="rpt" value="" />
		<input type="hidden" name="XML" value="" />	
		<input type="hidden" name="strLanguageCompleted" value="<%=session("language")%>" />		
		<input type="hidden" name="frmEDIYear" value="" />	
		<input type="hidden" name="frmSite" value="" />					
		<input type="hidden" name="frmSchool" value="" />
		<input type="hidden" name="frmTeacher" value="" />
		<input type="hidden" name="frmClass" value="" />					
		<input type="hidden" name="frmChild" value="" />
		<input type="hidden" name="frmNextChild" value="" />
		<input type="hidden" name="frmAction" value="" />
		<input type="hidden" name="frmSection" value="" />	
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
		<table border="1" style="border-color:#006600;" cellpadding="0" cellspacing="0" width="760">
			<tr>
				<td>
					<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">
						<tr>
							<td align="right" width="520"><font class="headerBlue"><%=lblEDI%> Questionnaire(<%=strEDIID%>)</font></td>
							<td align="right">
								<input type="button" value="<%=strExit%>" name="Exit" onclick="javascript:window.location='edi_teacher.asp';" id="Button1" />
								&nbsp;
							</td>	
						</tr>
						<tr><td colspan="2"><%="<br/>" & strError%></td></tr>
					</table>
					
					<%
					
					select case len(strMap)
						case 525, 519,498 ' Questionnaire
							' aChild
							if strError = "" then 
								' summary 
								Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""750"" align=""left"">"														
									' status
									Response.Write "<tr>"	
									Response.Write "<td align=""right""><font class=""boldTextBlack"">" & lblStatus & ":&nbsp;</font></td>"
									Response.Write "<td align=""left"">"
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
										Response.Write "<a href=""javascript:goConfirm_Lock('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','lock');"" class=""bigLinkBlue""><img src=""images/lock.gif"" border=""0"" alt=""Lock"" />" & lblCompletion & "</a>"
									end if 
									
									' determine the screen that the user is on
									if Request.Form("frmSection").Count = 0 OR Request.Form("frmSection") = "_Demographics" OR Request.Form("frmSection") = "" then 
										strCurrent = "_Demographics"
									else
										strCurrent = Request.Form("frmSection")
									end if 
									
									' set the SQL query to get then previous child in the class
									strSql = "SELECT right('0' + Max(intChild),2) as previousChild FROM children WHERE intClassID=" & strSite & strSchool & strTeacher & strClass & " AND intChild < " & strChild & " GROUP BY intClassID"
									set rstChildren = server.CreateObject("adodb.recordset")
									rstChildren.Open strSql, conn
									
									if not rstChildren.EOF then  
										Response.Write "&nbsp;"
										if  rstChildren("previousChild") <> "" then 
											strPrevious = "<a href=""javascript:goSaveEDIChild('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','" & right("0" & rstChildren("previousChild"),2) & "','" & strCurrent & "','" & strCurrent & "');"" class=""bigLinkBlue"">" & lblPrevious & "<img src=""images/student4.jpg"" border=""0"" width=""41"" height=""41"" title=""Previous Child"" alt=""Previous Child"" /></a>"
										else
											strPrevious = ""
										end if 
									end if 
									
									' close the recordset
									rstChildren.Close 
									
									' set the SQL query to get then next child in the class
									strSql = "SELECT right('0' + Min(intChild),2) as nextChild FROM children WHERE intClassID=" & strSite & strSchool & strTeacher & strClass & " AND intChild > " & strChild & " GROUP BY intClassID"
									
									' open the recordset
									rstChildren.Open strSql, conn
									
									if not rstChildren.EOF then					
										Response.Write "&nbsp;"
										if rstChildren("nextChild") <> "" then 
											strNext = "<a href=""javascript:goSaveEDIChild('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','" & right("0" & rstChildren("nextChild"),2) & "','" & strCurrent & "','" & strCurrent & "');"" class=""bigLinkBlue""><img src=""images/student.jpg"" height=""41"" width=""41"" border=""0"" title=""Next Child"" alt=""Next Child"" />" & lblNext & "</a>"
										else
											strNext = ""
										end if 
									end if 
									
									' close and kill the children recordset
									call close_adodb(rstChildren)
									
									Response.Write "</td>"
									Response.Write "</tr>"
									
									' spacer
									Response.Write "<tr><td colspan=""2"">&nbsp;</td></tr>"
									
									' site
									Response.Write "<tr>"	
									Response.Write "<td align=""right""><font class=""boldTextBlack"">" & lblSite & ":&nbsp;</font></td>"
									Response.Write "<td align=""left""><font class=""lrgregTextBlack"">&nbsp;&nbsp;" & aStudent(10,0) & "</font></td>"
									Response.Write "</tr>"
										
									' school
									Response.Write "<tr>"	
									Response.Write "<td align=""right""><font class=""boldTextBlack"">" & lblSchool & ":&nbsp;</font></td>"
									Response.Write "<td align=""left"" colspan=""2""><font class=""lrgregTextBlack"">&nbsp;&nbsp;" & aStudent(11,0) & "</font></td>"
									'Response.Write "<td width = ""350"" rowspan=""3"" align=""left"">&nbsp;&nbsp;<a href=""javascript:goTeacherClassReport('" & right("0" & strSite & strSchool & strTeacher & strClass,9) & "', '" & session("id") & "');"" class=""bigLinkBlue""><img src=""images/download.gif"" border=""0"" alt=""Download"" /> " & lblClassSummary & "</a></td>"								
									Response.Write "</tr>"
										
									' teacher
									Response.Write "<tr>"	
									Response.Write "<td align=""right""><font class=""boldTextBlack"">" & lblTeacher & ":&nbsp;</font></td>"
									Response.Write "<td align=""left""><font class=""lrgregTextBlack"">&nbsp;&nbsp;" & aStudent(12,0) & "</font></td>"
									Response.Write "</tr>"
										
									' local ID
									Response.Write "<tr>"	
									Response.Write "<td align=""right""><font class=""boldTextBlack"">" & lblLocal & ":&nbsp;</font></td>"
									Response.Write "<td align=""left""><font class=""lrgregTextBlack"">&nbsp;&nbsp;" & aStudent(3,0) & "</font></td>"
									'Response.Write "<td width = ""350"" rowspan=""2"">&nbsp;</td>"
									Response.Write "</tr>"
										
									' gender
									Response.Write "<tr>"	
									Response.Write "<td align=""right""><font class=""boldTextBlack"">" & lblSex & ":&nbsp;</font></td>"
									if strLanguage = "English" then 
										Response.Write "<td align=""left"" colspan=""2""><font class=""lrgregTextBlack"">&nbsp;&nbsp;" & aStudent(4,0) & "</font></td>"
									else
										Response.Write "<td align=""left"" colspan=""2""><font class=""lrgregTextBlack"">&nbsp;&nbsp;" & replace(replace(aStudent(4,0),"Male","masculin"),"Female","feminin") & "</font></td>"
									end if
									'Response.Write "<td width = ""350"" rowspan=""3"" align=""left"">&nbsp;&nbsp;&nbsp;<a href=""javascript:goTeacherEDIReport('" & strEDIID & "');"" class=""bigLinkBlue""><img border=""0"" src=""images/details.gif"" alt=""Details"" />" & lblSummary & "</a></td>"
									Response.Write "</tr>"
										
									' DOB
									Response.Write "<tr>"	
									Response.Write "<td align=""right""><font class=""boldTextBlack"">" & lblDOB & ":&nbsp;</font></td>"
									if isnull(aStudent(5,0))  then
									    Response.Write "<td align=""left""><font class=""lrgregTextBlack"">&nbsp;&nbsp;N/A</font></td>"
									else
									    if strLanguage = "English" then 
										    Response.Write "<td align=""left""><font class=""lrgregTextBlack"">&nbsp;&nbsp;" & right("00" & day(aStudent(5,0)),2) & "-" & monthname(datepart("m",aStudent(5,0)),true) & "-" & year(aStudent(5,0)) & "</font></td>"
									    else
										    Response.Write "<td align=""left""><font class=""lrgregTextBlack"">&nbsp;&nbsp;" & right("00" & day(aStudent(5,0)),2) & "-" & left(french_month(datepart("m",aStudent(5,0))),3) & "-" & year(aStudent(5,0)) & "</font></td>"
									    end if 
									end if 
									'Response.Write "<td width = ""350"" rowspan=""2"">&nbsp;</td>"
									Response.Write "</tr>"
										
									' Postal Code
									Response.Write "<tr>"	
									Response.Write "<td align=""right""><font class=""boldTextBlack"">" & lblPostal & ":&nbsp;</font></td>"
									Response.Write "<td align=""left""><font class=""lrgregTextBlack"">&nbsp;&nbsp;" & aStudent(6,0) & "</font></td>"
									'Response.Write "<td width = ""350"">&nbsp;</td>"
									Response.Write "</tr>"
									
									' spacer 
									Response.Write "<tr><td><br /></td></tr>"
								Response.Write "</table>"	
							' end first column	
							Response.Write "</td></tr>"
							
							'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
							' EDI SECTION - February 22, 2003
							'		- determine which section and show those questions
							'///////////////////////////////////////////////////////////////////
							' start second row
							Response.Write "<tr><td>"	
							
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							' DEMOGRAPHIC SECTION - February 22, 2003
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							if Request.Form("frmSection").Count = 0 OR Request.Form("frmSection") = "_Demographics" OR Request.Form("frmSection") = "" then 
							%>
								<!-- #include virtual="/shared/Section_Demographics.inc" -->
							<%
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							' SECTION A Physical Well Being - February 22, 2003
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							elseif Request.Form("frmSection") = "A" then
							%>
								<!-- #include virtual="/shared/SectionAtest.inc" -->
							<%
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							' SECTION B Language and Cognitive Skills - February 23, 2003
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							elseif Request.Form("frmSection") = "B" then
							%>
								<!-- #include virtual="/shared/SectionB.inc" -->
							<%
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							' SECTION C Physical Well Being - February 22, 2003
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							elseif Request.Form("frmSection") = "C" then
								%>
							<!-- #include virtual="/shared/SectionC.inc" -->
								<%								
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							' SECTION D Physical Well Being - February 22, 2003
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							elseif Request.Form("frmSection") = "D" then
								%>
								<!-- #include virtual="/shared/SectionD.inc" -->
								<%
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							' SECTION E Physical Well Being - February 22, 2003
							'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
							elseif Request.Form("frmSection") = "E" then
								%>
								<!-- #include virtual="/shared/SectionE.inc" -->
								<%
							' end EDI
							end if 	
						' end error
						end if 
							
						case 373, 367, 353, 359 ' child
							Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""500"" align=""center"">"						
							Response.Write "<tr><td>"	
							if strError = "" then 
								for intRow = 0 to ubound(aChildren,2)
									Response.Write "<font class=""regTextBlack"">" & right("00" & aChildren(1,intRow),2) & ") </font>"
									if isnull(aChildren(5,introw)) then 
										Response.Write "<a class=""reglinkMaroon"" href=""javascript:goChild('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & right("00" & aChildren(1,intRow),2) & "');""> N/A - " & aChildren(3,introw) & "</a>"	
									else
										childDob = achildren(5,introw)
										if session("language") = "English" then 
											Response.Write "<a class=""reglinkMaroon"" href=""javascript:goChild('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & right("00" & aChildren(1,intRow),2) & "');"">" & day(childDOB) & "-" & monthname(datepart("m",childDOB),true) & "-" & year(childDOb) & " - " & aChildren(3,introw) & "</a>"
										else
											Response.Write "<a class=""reglinkMaroon"" href=""javascript:goChild('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & right("00" & aChildren(1,intRow),2) & "');"">" & day(childDob) & "-" & left(French_Month(datepart("m",ChildDob)),3) & "-" & year(ChildDob) & " - " & aChildren(3,introw) & "</a>"
										end if 
									end if 
									Response.Write "<br/>"
								next
							end if 
							Response.Write "</td></tr></table>"	
						case 232, 226, 229, 230, 225 ' class
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
						case 89 	' teacher
							Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""500"" align=""center"">"						
							Response.Write "<tr><td>"	
							if strError = "" then 
								for intRow = 0 to ubound(aTeachers,2)
									Response.Write "<font class=""regTextBlack"">" & aTeachers(0,intRow) & ") </font>"
									Response.Write "<a class=""reglinkMaroon"" href=""javascript:goChild('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & right(aTeachers(0,intRow),2) & "','','');"">" & aTeachers(2,intRow) & "</a>"
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
		<input type="hidden" name="hdnLock" value="<%=blnLock%>" />	
		<input type="hidden" name="hdnCheckBoxes" value="<%=strCheckBoxes %>" />
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

function buildCheckBox(columnname,currentvalue)
    strCheck = "&nbsp;&nbsp;<input type=""checkbox"" value=""true"" id=""" & columnname & """ name=""" & columnname & """"
    if currentvalue = true then 
        strCheck = strCheck & " checked=""CHECKED"""
    end if                                      
    strCheck = strCheck & "/>"
    
    if len(strCheckBoxes) > 0 then 
        strCheckboxes = strCheckboxes & ","
    end if 
    
    strCheckboxes = strCheckboxes & columnname 
    
    buildcheckbox = strCheck    
End function
%>