<!-- #include virtual="/shared/security.asp" -->
<%
' public variables

' totals
dim intSites, intSchools, intTeachers, intClasses, intLanguage,intTeacher
' fields
dim strName, strEmail, strComments, Question1a, Question1b, Question1c
' arrays
dim aData, aFeedback
'on error resume next

dim aHeader(6)
' initialize the variable
intLanguage = ""
intTeacher = ""

' if the user has not logged in they will not be able to see the page
if blnSecurity then
	'call open_adodb(conn, "DATA")
	'call open_adodb(conn_tables, "TABLES")
	call open_adodb(conn, "MACEDI")
		
	set rstData = server.CreateObject("adodb.recordset")   		
	set rstFeedback = server.createobject("adodb.recordset")
    set rstQuestions = server.createobject("adodb.recordset")
  
	'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
	' Form Actions 
	' - November 8, 1008
	' - Andrew Renner
	'//////////////////////////////////////////////////////////////////////
	if Request.Form("Action") = "Update" then
        intTeacher = Request.Form("code") 
				
		' build the SQL statement
		strSQL = "UPDATE teacherTrainingFeedback SET "
		
		for each item in Request.Form 
			if NOT (item = "Action" OR item="strLanguage" OR  item ="code" OR item = "frmAction" OR item = "frmSection" OR item = "frmSite" OR item ="frmSchool" OR item ="frmTeacher" OR item ="frmClass" OR item = "frmChild" OR item = "frmNextChild" OR item = "days2" OR item = "days" OR item = "btnSave" OR item = "hdnLock" OR item = "Student" OR item = "classes" OR item = "email" OR item ="rpt" OR item ="XML" or item="CurrentSection") then 
				' added May 26 2004
				' removed feb 2 2006 - see below
				'blnUpdate = true
				if left(item ,3) = "str" then
					strSql = strSql & "[" & item & "] = " & checknull(Request.Form(item)) & ","
				else
					strSql = strSql & "[" & item & "] = " & Request.Form(item) & ","
				end if 
			end if 
		next
		
		' remove the last comma
		strSql = left(strSql,len(strSql)-1) & " WHERE intTeacherID = " & Request.Form("Code")
		
		'response.Write strSql
		' update the record
		conn.execute strSql
		
		
		' build the error string
		if conn.errors.count > 0 then 
			strError = "<font class=""regtextred"">Error Number : " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/></font>"
		end if
	end if
	
	strSql = "SELECT DISTINCT * FROM [teachers] WHERE strEmail ='" & session("id") & "' ORDER BY intTeacherID"
		
	' get the school specific teachers
	rstData.Open strSql, conn	
											
	if not rstData.EOF then
		' store info in array
		aData = rstData.GetRows 
											
		' get the number of teacher ID's on this email
		intTeachers = ubound(aData,2) + 1							
							
		' get the teacher
		' if the teacher is updated then the value will already be here
				
		if intTeacher = "" AND Request.QueryString("teacher").Count = 0 then 
			intTeacher = right("000" & aData(0,0),8)				
		else
			if intTeacher = "" then 
				' check to see if the teacherid in the querystring is one of this teachers
				for introw = 0 to ubound(adata,2)
					if Request.QueryString("teacher") = right("000" & aData(0,introw),8) then 
						intTeacher = Request.QueryString("teacher")
						exit for 
					end if
				next
			
				if intTeacher = "" then 
					' user entered a value that is not valid for this user
					Response.Redirect "edi_teacher_feedback.asp"
				end if 
			end if 
		end if

		' load the values
	   call load_values(intTeacher)
	else
		intTeachers = 0
		call add_mode
	end if 
	
	
%>
<html>
<head>
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
		<a class="reglinkMaroon" href="edi_teacher.asp"><%=strHome%></a>&nbsp;<font class="regtextblack">></font>&nbsp;<font class="boldtextblack"><%=lblTrainingFeedback%></font>
		<table border="1" bordercolor="006600" cellpadding="0" cellspacing="0" width="760">
		<tr><td>
			<form name="Screens" method="POST" action="edi_teacher_feedback.asp"> 
			<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">
			    <tr>
				    <td align="right" width="550"><font class="headerBlue"><%=lblTrainingFeedback%></font></td>
				    <td align="right">
					    <input type="hidden" name="Action" value="">					
					    <input type="hidden" name="strLanguage" value="">
						<input type="hidden" name="strLanguageCompleted" value="<%=session("language")%>" />						
					    <%if strError = "" then %>
					    <input type="button" value="<%=strSave%>" name="Update" onClick="javascript:update_TeacherFeedbackCheck(this.name);">
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
					<td align="left" width="200" >
						<font class="boldtextblack"><%=lblCode%> :&nbsp;&nbsp;</font>
					</td>
					<td>
					<%
						' teacher						
						Response.Write "<select name=""code"" onChange=""javascript:window.location='edi_teacher_feedback.asp?teacher=' + this.value;"">"
						for intRow = 0 to ubound(aData,2)						
							Response.Write "<option value = """ & right("000" & aData(0,intRow),8) & """"
							' write the teacher
							if intTeacher = right("000" &  aData(0,intRow),8) then 
								Response.Write " selected"
							end if
							Response.Write ">" & right("000" &  aData(0,intRow),8) & "</option>"
						next
						Response.Write "</select>" 
						%>
					</td>
				</tr>
			</table>		
			<br />
			<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">
				<tr valign="top">
					<td align="left">
						<br />
						<font class="subheaderBlue"><%=lblFeedbackFeedback%> :&nbsp;&nbsp;</font>
					</td>
				</tr>
			</table>
			<table border="1" cellpadding="0" cellspacing="0" width="750" align="center">
                <tr>
	                <td align="left">
		                &nbsp;&nbsp;<font class="boldtextblack">1a) </font>
		            </td>
		            <td align="left">
		                &nbsp;&nbsp;<font class="boldtextblack"><%=lblFeedbackQ1%> &nbsp;&nbsp;</font>
	                </td>
	                <td align="center">
		                <font class="boldtextblack">
			                <input type="radio" name="Question1a" value="1" <%if Question1a = 1 then Response.Write "checked"%>><%=lblFeedbackYesGoto%> 2 &nbsp;&nbsp;<input type="radio" name="Question1a" value="2" <%if Question1a = 2 then Response.Write "checked"%>><%=lblFeedbackNo%>
		                </font>
	                </td>
                </tr>						               
                </tr>	
                <tr>
	                <td align="left">
		                &nbsp;&nbsp;<font class="boldtextblack">1b) </font>
		            </td>
		            <td align="left">
		                &nbsp;&nbsp;<font class="boldtextblack"><%=lblFeedbackQ2%> &nbsp;&nbsp;</font>
	                </td>
	                <td align="center">
		                <font class="boldtextblack">
			                <input type="radio" name="Question1b" value="1" <%if Question1b = 1 then Response.Write "checked"%>><%=lblFeedbackYes%> &nbsp;&nbsp;<input type="radio" name="Question1b" value="2" <%if Question1b = 2 then Response.Write "checked"%>><%=lblFeedbackNoGoto%> 2
		                </font>
	                </td>
                </tr>									                
                <tr>
	                <td align="left">
		                &nbsp;&nbsp;<font class="boldtextblack">1c) </font>
		            </td>
		            <td align="left">
		                &nbsp;&nbsp;<font class="boldtextblack"><%=lblFeedbackQ3%> &nbsp;&nbsp;</font>
	                </td>
	                <td align="center">
		                <font class="boldtextblack">
			                <input type="radio" name="Question1c" value="1" <%if Question1c = 1 then Response.Write "checked"%>><%=lblFeedbackElectronic%> &nbsp;&nbsp;<input type="radio" name="Question1c" value="2" <%if Question1c = 2 then Response.Write "checked"%>><%=lblFeedbackPaper%>
		                </font>
	                </td>
                </tr>							                        
			</table>
			
			<br />		
			<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">
				<tr valign="top">
					<td colspan="2">
				        <!-- #include virtual="/shared/SectionTeacherFeedback.inc" -->
					</td>
				</tr>											
				</table>
			</form>
			</td>
			</tr>						
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
	call close_adodb(rstFeedback)
	call close_adodb(rstData)
	'call close_adodb(conn)
	'call close_adodb(conn_tables)
	call close_adodb(conn)
' security
end if

' set form defaults
sub add_mode()
	' load the first site
	Question1a = ""
	Question1b = ""
	Question1c = ""
end sub

sub load_values(intTeacher)	
' get the demographic data
	strSql = "SELECT * FROM teacherTrainingFeedback WHERE intTeacherID = " & intTeacher
  		
	rstFeedback.Open strSql, conn
	
	Question1a = rstFeedback("Question1a")
	Question1b = rstFeedback("Question1b")
	Question1c = rstFeedback("Question1c")
end sub

%>