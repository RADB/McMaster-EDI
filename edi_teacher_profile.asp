<!-- #include virtual="/shared/security.asp" -->
<%
' public variables
dim aData
dim Question1,Question2,Question3	,	Question4	,	Question5	,	Question6	,	Question7YH	,	Question7YNH	,	Question7NNH	,	Question7NNone	,	Question7NTime	,	Question7NFamiliar,		Question7Other	,	Question7OtherText	,StudentsInClass
dim strName, strEmail,intSex, intAge, strPhone, strFax, intQ5a, intQ5b,intQ5c, intQ6a, intQ6b, intQ6c, intQ6d, intQ6e, intQ6f, intQ6g, intQ6h, intQ6i, intQ6j, intQ6k
dim intMth1, intMth2, intMth3,intMth4, intMth5, intMth6,intMth7, intMth8, intMth9, intYr1, intYr2, intYr3,intYr4, intYr5, intYr6,intYr7, intYr8, intYr9
on error resume next

' if the user has not logged in they will not be able to see the page
if blnSecurity then
	'call open_adodb(conn, "DATA")
	call open_adodb(conn, "MACEDI")
	
	set rstData = server.CreateObject("adodb.recordset")
	set rstLanguages = server.CreateObject("adodb.recordset")
	
	' open all languages
	rstLanguages.Open "SELECT LID, english FROM [LU Languages] ORDER BY english", conn
	
	' store all languages in array
	aLanguages = rstLanguages.GetRows 
	
	' close and kill the langauges recordset
	call close_adodb(rstLanguages)
	
	if Request.Form("Action") = "Update" then
		intTeacher = Request.Form("code") 

        ' removed phone 2009
        'strPhone = " & checkNull(Request.Form("phone")) & ",			
		' build the SQL statement
		strSql = "UPDATE teachers " & _
  				 "SET strName = " & checkNull(Request.Form("name")) & ", strEmail = " & checkNull(Request.Form("email")) & ", strFax = " & checkNull(Request.Form("fax")) & ", intSex = " & checkNull(Request.Form("sex")) & ", intAge = " & checkNull(Request.Form("age")) & ", intQ5a = " & checkNull(Request.Form("intQ5a")) & ",intQ5b = " & checkNull(Request.Form("intQ5b")) & ",intQ5c = " & checkNull(Request.Form("intQ5c")) & ",intQ6a = " & checkNull(Request.Form("intQ6a")) & ", intQ6b = " & checkNull(Request.Form("intQ6b")) & ",intQ6c = " & checkNull(Request.Form("intQ6c")) & ",intQ6d = " & checkNull(Request.Form("intQ6d")) & ",intQ6e = " & checkNull(Request.Form("intQ6e")) & ",intQ6f = " & checkNull(Request.Form("intQ6f")) & ",intQ6g = " & checkNull(Request.Form("intQ6g")) & ",intQ6h = " & checkNull(Request.Form("intQ6h")) & ",intQ6i = " & checkNull(Request.Form("intQ6i")) & ",intQ6j = " & checkNull(Request.Form("intQ6j")) & ",intQ6k = " & checkNull(Request.Form("intQ6k")) & _
				 " WHERE intTeacherID = " & intTeacher
	
		'Response.Write strSql
		' update the record
		conn.execute strSql
		
		' build the error string
		if conn.errors.count > 0 then 
			strError = "<font class=""regtextred"">Error Number : " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/></font>"
		end if
			
		' build the SQL statement		
		 strSQL = "UPDATE  teacherParticipation SET Question1 = " & checkNull(Request.Form("Question1")) & ", Question2 = " & checkNull(Request.Form("Question2")) & ", Question3 = " & checkNull(Request.Form("Question3")) & ", Question4 = " & checkNull(Request.Form("Question4")) & ", Question5 = " & checkNull(Request.Form("Question5")) & ", Question6 = " & checkNull(Request.Form("Question6")) & ", Question7YH = " & checkNull(Request.Form("Question7YH")) & ", Question7YNH = " & checkNull(Request.Form("Question7YNH")) & ", Question7NNH = " & checkNull(Request.Form("Question7NNH")) & ", Question7NNone = " & checkNull(Request.Form("Question7NNone")) & ", Question7NTime = " & checkNull(Request.Form("Question7NTime")) & ", Question7NFamiliar = " & checkNull(Request.Form("Question7NFamiliar")) & ", Question7Other = " & checkNull(Request.Form("Question7Other")) & ", Question7OtherText = " & checkNull(Request.Form("Question7OtherText")) &", StudentsInClass = " & checkNull(Request.Form("StudentsInClass")) & ",strLanguageCompleted = '" & session("Language") & "'" &_
                " WHERE intTeacherID = " & intTeacher
		
		'Response.Write strSql
		' update the record
		conn.execute strSql
		
		' build the error string
		if conn.errors.count > 0 then 
			strError = strError & "<font class=""regtextred"">Error Number : " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/></font>"
		end if
		
		' build the error string
		if conn.errors.count > 0 then 
			strError = "<font class=""regtextred"">Error Number : " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/></font>"
		else
			for introw = 1 to Request.form("intClasses")
				' extract the months from the form 
				intInsertMonths = right(Request.Form("intQ5d" & intRow + 3),len(Request.Form("intQ5d" & intRow + 3))-9)
				
				' update the classes
				strSQL = "UPDATE classes SET intMonths =" & intInsertMonths & " WHERE intClassID = " &  left(Request.Form("intQ5d" & intRow + 3),9)
				conn.execute strSql
			next 
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
					Response.Redirect "edi_teacher_profile.asp"
				end if 
			end if 
		end if
				
		' load the values
	   call load_values(intTeacher)
	else
		intTeachers = 0
		call add_mode
	end if 

	' close the recordset
	rstData.Close 
	
	'********************************************************
	' participation data
	
	' get the teacher participation information
    strSql = "SELECT DISTINCT t.strName, tp.* FROM [teacherParticipation] tp, teachers t WHERE tp.intTeacherID = t.intTeacherID AND t.strEmail ='" & session("id") & "'"
		
	rstData.Open strSql, conn	
								
	if not rstData.EOF then
		' store info in array
		aData = rstData.GetRows 
		
		' get the number of teacher ID's on this email - used to show the update button
		intTeachers = ubound(aData,2) + 1	
				
		if intTeacher = "" AND Request.QueryString("teacher").Count = 0 then 
			intTeacher = right("000" & aData(1,0),8)
		else
			if intTeacher = "" then 
				' check to see if the teacherid in the querystring is one of this teachers
				for introw = 0 to ubound(adata,2)
					if Request.QueryString("teacher") = right("000" & aData(1,introw),8) then 
						intTeacher = Request.QueryString("teacher")
						exit for 
					end if
				next
			
				if intTeacher = "" then 
					' user entered a value that is not valid for this user
					Response.Redirect "edi_teacher_profile.asp"
				end if 
			end if 
		end if
																		
		' load the values
	   call load_ParticipationValues(intTeacher)
	else
		intTeachers = 0
		' only allow add to teachers...
		'call add_mode
	end if 

	' close the recordset
	rstData.Close 
	
	' select all classes that this teacher has  (0) None Selected (1) English (2) French (3) Other		
	
	'if session("Language") = "English" then			
    strSql = "Exec GetTeacherClassesByEmail '" & session("Language") & "','" & session("ID") & "'"
	'	strSql = "SELECT c.intClassID, iif(c.intLanguage=1, 'English',iif(c.intLanguage=2, 'French',iif(c.intLanguage=3, 'Other','Unknown'))) as strLanguage, count(ch.strEDIID) as intStudents, sum(iif(ch.chkCompleted=true,1,0)) AS Completed, int(c.intMonths/12) as years, (c.intMonths mod 12) as months " & _
	'				"FROM (classes c LEFT JOIN children ch ON c.intClassID = ch.intClassID) " & _
	'				"LEFT JOIN teachers t ON c.intTeacherID = t.intTeacherID " & _
	'				"WHERE t.strEmail = '" & Session("ID") & "'" & _
	'				" GROUP BY c.intClassID, iif(c.intLanguage=1, 'English',iif(c.intLanguage=2, 'French',iif(c.intLanguage=3, 'Other','Unknown'))), int(c.intMonths/12), (c.intMonths mod 12)"
	
	'else
	'	strSql = "SELECT c.intClassID, iif(c.intLanguage=1, 'Anglais',iif(c.intLanguage=2, 'Français ',iif(c.intLanguage=3, 'Autre','Unknown'))) as strLanguage, count(ch.strEDIID) as intStudents, sum(iif(ch.chkCompleted=true,1,0)) AS Completed, int(c.intMonths/12) as years, (c.intMonths mod 12) as months " & _
	'				"FROM (classes c LEFT JOIN children ch ON c.intClassID = ch.intClassID) " & _
	'				"LEFT JOIN teachers t ON c.intTeacherID = t.intTeacherID " & _
	'				"WHERE t.strEmail = '" & Session("ID") & "'" & _
	'				" GROUP BY c.intClassID, iif(c.intLanguage=1, 'Anglais',iif(c.intLanguage=2, 'Français ',iif(c.intLanguage=3, 'Autre','Unknown'))), int(c.intMonths/12), (c.intMonths mod 12)"
					
					'"WHERE c.intTeacherID = " & intTeacher & _
	'end if
	
	' open list of classes and teachers at this school
	rstData.Open strSql, conn,1
		
	if rstData.EOF then 
		intClasses = 0
	else
		'intClasses = rstData.RecordCount 
        intClasses = rstData("Classes")
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
	<%	
	' provinces recordset
	set rstProvinces = server.CreateObject("Adodb.recordset")
	%>
	<form name="Screens" method="POST" action="edi_teacher_profile.asp"> 
		<a class="reglinkMaroon" href="edi_teacher.asp"><%=strHome%></a>&nbsp;<font class="regtextblack">></font>&nbsp;<font class="boldtextblack"><%=lblTitle%></font>
		<table border="1" bordercolor="006600" cellpadding="0" cellspacing="0" width="760">
		<tr><td>
			<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">
			<tr>
				<td align="right" width="430"><font class="headerBlue"><%=lblTitle%></font></td>
				<td align="right">
					<input type="hidden" name="Action" value="">					
				<%
					if Request.form("Action") <> "Add" AND intTeachers > 0 then  
					%>
						<input type="button" value="<%=strSave%>" name="Update"  onClick="javascript:update_TeacherCheck(this.name);">
					<%
					end if 
					%>
					<input type="button" value="<%=strExit%>" name="Exit" onClick="javascript:window.location='edi_teacher.asp';">
					&nbsp;
				</td>
			</tr>
			<tr><td colspan="2"><%=strError%></td></tr>
			<!-- sections here-->
			</table>
			<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">
				<tr valign="top">
					<td align="right" nowrap>
						<font class="boldtextblack"><%=lblCode%> :&nbsp;&nbsp;</font>
					</td>
					<td>
					<%
						' teacher						
						Response.Write "<select name=""code"" onChange=""javascript:window.location='edi_teacher_profile.asp?teacher=' + this.value;"">"
						for intRow = 0 to ubound(aData,2)						
							Response.Write "<option value = """ & right("000" & aData(1,intRow),8) & """"
							' write the teacher
							if intTeacher = right("000" &  aData(1,intRow),8) then 
								Response.Write " selected"
							end if
							Response.Write ">" & right("000" &  aData(1,intRow),8) & "</option>"
						next
						Response.Write "</select>" 
						%>
					</td>
				</tr>
				<tr valign="top">
					<td align="right">
						<font class="boldtextblack"><%=lblName%> :&nbsp;&nbsp;</font>
					</td>
					<td>
						<input type="text" size="80" name="name" value="<%=strName%>">
					</td>
				</tr>
				
				<tr valign="top">
					<!--
					Remove Phone 2009
					<td align="right">
						<font class="boldtextblack"><%=lblPhone%> :&nbsp;&nbsp;</font>
					</td>
					<td>
						<input type="text" size="15" name="phone" value="<%=strPhone%>" maxlength="14"> 
						<font class="boldtextblack"><%=lblFax%> :&nbsp;&nbsp;</font>
						<input type="text" size="15" name="fax" value="<%=strFax%>" maxlength="14"> 
					</td>-->
					<td align="right">
						<font class="boldtextblack"><%=lblFax%> :&nbsp;&nbsp;</font>
					</td>
					<td>												
						<input type="text" size="15" name="fax" value="<%=strFax%>" maxlength="14"> 
					</td>
				</tr>
				<tr>
					<td align="right">
						<font class="boldtextblack"><%=lblEmail%> :&nbsp;&nbsp;</font>
					</td>
					<td>
						<input type="text" size="40" name="email" value="<%=strEmail%>">  
					</td>
				</tr>
			</table>

			
			<!-- Participation Data -->
			
				<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">
				<tr valign="top">
					<td align="left">
						<br />
						<font class="subheaderBlue"><%=lblTitleSubHeader%> :&nbsp;&nbsp;</font>
					</td>
				</tr>
			</table>
			
			<br />
			<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">
				<tr>
				    <td>
				        
				        <table border="1" cellpadding="0" cellspacing="0" width="710" align="center">
			                <tr>
				                <td>
					                <table border="0" cellpadding="0" cellspacing="0" width="700" align="center">
						                <tr>
							                <td>
								                <font class="boldtextblack">1) <%=lblP1%> &nbsp;&nbsp;</font>
							                </td>
							                <td align="center">
								                <font class="boldtextblack">
									                <input type="radio" name="Question1" value="1" <%if Question1 = 1 then Response.Write "checked"%>><%=lblYesGoto%> 5 &nbsp;&nbsp;<input type="radio" name="Question1" value="2" <%if Question1 = 2 then Response.Write "checked"%>><%=lblNo%>
								                </font>
							                </td>
						                </tr>
						                <tr valign="top">
					                        <td>
						                        <font class="boldtextblack">2) <%=lblP2%>&nbsp;&nbsp;</font>
					                        </td>
					                        <td>
						                        <select name="Question2">
							                        <option value=""></option>
							                        <%
								                        Response.Write "<option value=""1"""
								                        if Question2 = 1 then Response.Write " selected"
								                        Response.Write ">1</option>"
								                        Response.Write "<option value=""2"""
								                        if Question2 = 2 then Response.Write " selected"
								                        Response.Write ">2</option>"
								                        Response.Write "<option value=""3"""
								                        if Question2 = 3 then Response.Write " selected"
								                        Response.Write ">3</option>"
								                        Response.Write "<option value=""4"""
								                        if Question2 = 4 then Response.Write " selected"
								                        Response.Write ">" & lbl4OrMore & "</option>"
							                        %>
						                        </select>									
					                        </td>
				                        </tr>	
				                        <tr>
							                <td>
								                <font class="boldtextblack">3) <%=lblP3%> &nbsp;&nbsp;</font>
							                </td>
							                <td align="center">
								                <font class="boldtextblack">
									                <input type="radio" name="Question3" value="1" <%if Question3 = 1 then Response.Write "checked"%>><%=lblYesGoto%> 4 &nbsp;&nbsp;<input type="radio" name="Question3" value="2" <%if Question3 = 2 then Response.Write "checked"%>><%=lblNo%>
								                </font>
							                </td>
						                </tr>			
						                <tr valign="top">
					                        <td>
						                        <font class="boldtextblack">4) <%=lblP4%>&nbsp;&nbsp;</font>
					                        </td>
					                        <td>
						                        <select name="Question4">
							                        <option value=""></option>
							                        <%
								                        Response.Write "<option value=""1"""
								                        if Question4 = 1 then Response.Write " selected"
								                        Response.Write ">1</option>"
								                        Response.Write "<option value=""2"""
								                        if Question4 = 2 then Response.Write " selected"
								                        Response.Write ">2</option>"
								                        Response.Write "<option value=""3"""
								                        if Question4 = 3 then Response.Write " selected"
								                        Response.Write ">3</option>"
								                        Response.Write "<option value=""4"""
								                        if Question4 = 4 then Response.Write " selected"
								                        Response.Write ">" & lbl4OrMore & "</option>"
							                        %>
						                        </select>									
					                        </td>
				                        </tr>	
				                        <tr>
							                <td>
								                <font class="boldtextblack">5) <%=lblP5%> &nbsp;&nbsp;</font>
							                </td>
							                <td align="center">
								                <font class="boldtextblack">
									                <input type="radio" name="Question5" value="1" <%if Question5 = 1 then Response.Write "checked"%>><%=lblYesGoto%> 6 &nbsp;&nbsp;<input type="radio" name="Question5" value="2" <%if Question5 = 2 then Response.Write "checked"%>><%=lblNo%>
								                </font>
							                </td>
						                </tr>			
				                        <tr valign="top">
					                        <td>
						                        <font class="boldtextblack">6) <%=lblP6%>&nbsp;&nbsp;</font>
					                        </td>
					                        <td>					  					                            
						                        <select name="Question6">							                        
						                            <option value=""></option>
							                        <%
							                            Response.Write "<option value=""1"""
								                        if Question6 = 1 then 
								                            Response.Write " selected"
								                        end if 
								                        Response.Write ">" & lblVery & "</option>"
								                        Response.Write "<option value=""2"""
								                        if Question6 = 2 then 
								                            Response.Write " selected"
								                        end if 
								                        Response.Write ">" & lblSomewhat & "</option>"
								                        Response.Write "<option value=""3"""
								                        if Question6 = 3 then 
								                            Response.Write " selected"
								                        end if 
								                        Response.Write ">" & lblNotatall & "</option>"								                        
							                        %>
						                        </select>									
					                        </td>
				                        </tr>					                        						                
						            </table>
					            </td>
				            </tr>
				        </table>    
				        							
				    </td>
				</tr>
			</table>
			
			<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">
				<tr valign="top">
					<td align="left">
						<br />
						<font class="subheaderBlue"><%=lblP7%> :&nbsp;&nbsp;</font>
					</td>
				</tr>
			</table>
			
			<br />
			<table border="1" cellpadding="0" cellspacing="0" width="710" align="center">
			    <tr>
	                <td>
		                <table border="0" cellpadding="0" cellspacing="0" width="700" align="center">			    
				            <tr>
				                <td align="left">
		                            <font class="boldtextblack">
			                            <input type="checkbox" name="Question7YH" value="1" <%if Question7YH = 1 then Response.Write "checked"%>><%=lblQuestion7YH%> 			                            
		                            </font>
	                            </td>
	                            <td align="left">
		                            <font class="boldtextblack">
			                            <input type="checkbox" name="Question7NNH" value="1" <%if Question7NNH = 1 then Response.Write "checked"%>><%=lblQuestion7NNH%> &nbsp;&nbsp;			                            
		                            </font>
	                            </td>
	                        </tr>	            
	                        <tr>
				                <td align="left">
		                            <font class="boldtextblack">			                            
			                            <input type="checkbox" name="Question7YNH" value="1" <%if Question7YNH = 1 then Response.Write "checked"%>><%=lblQuestion7YNH%> &nbsp;&nbsp;			                            
		                            </font>
	                            </td>
	                            <td align="left">
		                            <font class="boldtextblack">			                            
			                            <input type="checkbox" name="Question7NNone" value="1" <%if Question7NNone = 1 then Response.Write "checked"%>><%=lblQuestion7NNone%> &nbsp;&nbsp;
		                            </font>
	                            </td>
	                        </tr>	            
	                        <tr>
				                <td align="left">
		                            <font class="boldtextblack">
			                            <input type="checkbox" name="Question7Other" value="1" <%if Question7Other = 1 then Response.Write "checked"%>><%=lblQuestion7Other%> &nbsp;&nbsp;
		                            </font>
	                            </td>
	                            <td align="left">
		                            <font class="boldtextblack">			                            
			                            <input type="checkbox" name="Question7NTime" value="1" <%if Question7NTime = 1 then Response.Write "checked"%>><%=lblQuestion7NTime%> &nbsp;&nbsp;
		                            </font>
	                            </td>
	                        </tr>	            
	                        <tr>
				                <td align="left">
		                            <font class="boldtextblack">
			                            <input type="text" size="30" name="Question7OtherText" value="<%=Question7OtherText%>">                
		                            </font>
	                            </td>
	                            <td align="left">
		                            <font class="boldtextblack">
			                            <input type="checkbox" name="Question7NFamiliar" value="1" <%if Question7NFamiliar= 1 then Response.Write "checked"%>><%=lblQuestion7NFamiliar%> &nbsp;&nbsp;
		                            </font>
	                            </td>
	                        </tr>	            
	                    </table>
	                 </td>       
	            </tr>
	        </table>	
			
			<!-- End Participation Data -->
			
			<!-- demographics -->
			<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">
				<tr valign="top">
					<td align="left">
						<br />
						<font class="subheaderBlue"><%=lblParticipationDemo%> :&nbsp;&nbsp;</font>
					</td>
				</tr>
			</table>
			
			<br />  
			<table border="1" cellpadding="0" cellspacing="0" width="550" align="center">
			<tr>
				<td>
					<table border="0" cellpadding="0" cellspacing="0" width="540">
					    <tr>
			                <td>
						        <font class="boldtextblack">&nbsp; <%=lblStudentsInClass%> :&nbsp;&nbsp;</font>
					        </td>
					        <td>
						        <select name="StudentsInClass">
							        <option value=""></option>
							        <%
							            for i = 0 to 99
                                            Response.Write "<option value = """ & right("00" & i,2) & """"
                                            ' write the teacher
                                            if right("00" & StudentsInClass,2) = right("00" &  i,2) then 
                                                Response.Write " selected"
                                            end if
                                            Response.Write ">" & right("00" &  i,2) & "</option>"
								        next 
							        %>
						        </select>
        				     </td>
                        </tr>
					    <tr>
			                <td>
						        <font class="boldtextblack">&nbsp; <%=lblGender%> :&nbsp;&nbsp;</font>
					        </td>
					        <td>
						        <select name="sex">
							        <option value=""></option>
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
                        <tr>
                            <td>
                                    <font class="boldtextblack">&nbsp; <%=lblAge%> :&nbsp;&nbsp;</font>
                            </td>
					        <td>
                                    <select name="age">
	                                    <option value=""></option>
	                                    <%
	                                    Response.Write "<option value=""2"""
	                                    if intAge = 2 then Response.Write " selected"
	                                    Response.Write ">20-29</option>"
	                                    Response.Write "<option value=""3"""
	                                    if intAge = 3 then Response.Write " selected"
	                                    Response.Write ">30-39</option>"
	                                    Response.Write "<option value=""4"""
	                                    if intAge = 4 then Response.Write " selected"
	                                    Response.Write ">40-49</option>"
	                                    Response.Write "<option value=""5"""
	                                    if intAge = 5 then Response.Write " selected"
	                                    Response.Write ">50-59</option>"
	                                    Response.Write "<option value=""6"""
	                                    if intAge = 6 then Response.Write " selected"
	                                    Response.Write ">60 +</option>"
                        				
	                                    %>
                                    </select>            						
                                </td>
                            </tr>
                        </table>
			        </td>
		        </tr>
			</table>
				
			<!-- table of other educational pursuits-->
			<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">
				<tr valign="top">
					<td align="left">
						<br />
						<font class="subheaderBlue"><%=lbl5%> :&nbsp;&nbsp;</font>
					</td>
				</tr>
			</table>
			
			<br />  
			<table border="1" cellpadding="0" cellspacing="0" width="550" align="center">
			<tr>
				<td>
					<table border="0" cellpadding="0" cellspacing="0" width="540" align="center">
						<tr>
							<td>
								<font class="boldtextblack">
									&nbsp;&nbsp;a) <%=lbl5a%> &nbsp;&nbsp;
								</font>
							</td>
							<td align="center">
								<select name="yr1" onChange="javascript:document.forms.Screens.intQ5a.value = Number(this.value * 12) + Number(document.forms.Screens.mth1.value);">
									<%
									for introw = 0 to 40 
										Response.Write "<option value=""" & introw & """"
										if introw = intyr1 then 
											Response.Write " selected"
										end if 
										Response.Write ">" & introw & "</option>"
									next 
									%>
								</select>
								<font class="boldtextblack"><%=lblYrs%></font>
								<select name="mth1" onChange="javascript:document.forms.Screens.intQ5a.value = Number(document.forms.Screens.yr1.value * 12) + Number(this.value);">
									<%
									for introw = 0 to 11 
										Response.Write "<option value=""" & introw & """"
										if introw = intmth1 then 
											Response.Write " selected"
										end if 
										Response.Write ">" & introw & "</option>"
									next 
									%>
								</select>
								<font class="boldtextblack"><%=lblMths%></font>
								<input type="hidden" name="intQ5a" size="5" value="<%=intQ5a%>">
							</td>
						</tr>
						<tr>
							<td>
								<font class="boldtextblack">
									&nbsp;&nbsp;b) <%=lbl5b%>&nbsp;&nbsp;
								</font>
							</td>
							<td align="center">
								<select name="yr2" onChange="javascript:document.forms.Screens.intQ5b.value = Number(this.value * 12) + Number(document.forms.Screens.mth2.value);">
									<%
									for introw = 0 to 40 
										Response.Write "<option value=""" & introw & """"
										if introw = intyr2 then 
											Response.Write " selected"
										end if 
										Response.Write ">" & introw & "</option>"
									next 
									%>
								</select>
								<font class="boldtextblack"><%=lblYrs%></font>
								<select name="mth2" onChange="javascript:document.forms.Screens.intQ5b.value = Number(document.forms.Screens.yr2.value * 12) + Number(this.value);">
									<%
									for introw = 0 to 11 
										Response.Write "<option value=""" & introw & """"
										if introw = intmth2 then 
											Response.Write " selected"
										end if 
										Response.Write ">" & introw & "</option>"
									next 
									%>
								</select>
								<font class="boldtextblack"><%=lblMths%></font>
								<input type="hidden" name="intQ5b" size="5" value="<%=intQ5b%>">
							</td>
						</tr>
						<tr>
							<td>
								<font class="boldtextblack">
									&nbsp;&nbsp;c) <%=lbl5c%>&nbsp;&nbsp;
								</font>
							</td>
							<td align="center">
								<select name="yr3" onChange="javascript:document.forms.Screens.intQ5c.value = Number(this.value * 12) + Number(document.forms.Screens.mth3.value);">
									<%
									for introw = 0 to 40 
										Response.Write "<option value=""" & introw & """"
										if introw = intyr3 then 
											Response.Write " selected"
										end if 
										Response.Write ">" & introw & "</option>"
									next 
									%>
								</select>
								<font class="boldtextblack"><%=lblYrs%></font>
								<select name="mth3" onChange="javascript:document.forms.Screens.intQ5c.value = Number(document.forms.Screens.yr3.value * 12) + Number(this.value);">
									<%
									for introw = 0 to 11 
										Response.Write "<option value=""" & introw & """"
										if introw = intmth3 then 
											Response.Write " selected"
										end if 
										Response.Write ">" & introw & "</option>"
									next 
									%>
								</select>
								<font class="boldtextblack"><%=lblMths%></font>
								<input type="hidden" name="intQ5c" size="5" value="<%=intQ5c%>">
							</td>
						</tr>
						<%
						' display the length of time at each class
						if intClasses > 0 then 
							intCount = 4
							Response.Write "<tr><td colspan=""2""><font class=""boldtextblack"">&nbsp;&nbsp;d) " & lbl5d & "</td></tr>"
							do while not rstData.eof
								Response.Write "<tr><td align=""center""><font class=""boldtextblack"">" & right("000" & rstData("intClassID"),9) & "</font></td>"
								Response.Write "<td align=""center""><select name=""yr" & intcount & """ onChange=""javascript:document.forms.Screens.intQ5d" & intCount  & ".value = '" & right("000" & rstData("intClassID"),9) & "' + (Number(this.value * 12) + Number(document.forms.Screens.mth" & intCount & ".value));"">"
								for introw = 0 to 40 
									Response.Write "<option value=""" & introw & """"
									if introw = rstData("years") then 
										Response.Write " selected"
									end if 
									Response.Write ">" & introw & "</option>"
								next 
								Response.Write "</select>"
									
								Response.Write "<font class=""boldtextblack"">&nbsp;" & lblYrs & "&nbsp;</font>"
								Response.Write "<select name=""mth" & intCount & """ onChange=""javascript:document.forms.Screens.intQ5d" & intCount & ".value = '" & right("000" & rstData("intClassID"),9) & "' + (Number(document.forms.Screens.yr" & intCount & ".value * 12) + Number(this.value));"">"
								for introw = 0 to 11 
									Response.Write "<option value=""" & introw & """"
									if introw = rstData("months") then 
										Response.Write " selected"
									end if 
									Response.Write ">" & introw & "</option>"
								next 
								Response.Write "</select>"
								Response.Write "<font class=""boldtextblack"">&nbsp;" & lblMths & "&nbsp;</font>"
								Response.Write "<input type=""hidden"" name=""intQ5d" & intCount & """ size=""5"" value=""" & right("000" & rstData("intClassID"),9) & (rstData("years") * 12 + rstData("months")) & """>"
								
								Response.Write "</td></tr>"
								intCount = intCount + 1 
								rstData.MoveNext 
							loop 
							' passes the number of classes for updating
							Response.Write "<input type=""hidden"" name=""intClasses"" size=""5"" value=""" & intClasses & """>"
							rstData.movefirst	
						end if 
						%>
					</table>
				</td>
			</tr>
			</table>
			
			<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">
				<tr valign="top">
					<td align="left">
						<br />
						<font class="subheaderBlue"><%=lbl6%> :&nbsp;&nbsp;</font>
					</td>
				</tr>
			</table>
			
			<br />
			
			<!-- table of other educational pursuits-->

			<table border="1" cellpadding="0" cellspacing="0" width="550" align="center">
			<tr>
				<td>
					<table border="0" cellpadding="0" cellspacing="0" width="540" align="center">
						<tr>
							<td>
								<font class="boldtextblack">
									&nbsp;&nbsp;a) <%=lblA%> &nbsp;&nbsp;
								</font>
							</td>
							<td align="center">
								<font class="boldtextblack">
									<input type="radio" name="intQ6a" value="1" <%if intq6a = 1 then Response.Write "checked"%>><%=lblYes%> &nbsp;&nbsp;<input type="radio" name="intQ6a" value="2" <%if intq6a = 2 then Response.Write "checked"%>><%=lblNo%>
								</font>
							</td>
						</tr>
						<tr>
							<td>
								<font class="boldtextblack">
									&nbsp;&nbsp;b) <%=lblB%> &nbsp;&nbsp;
								</font>
							</td>
							<td align="center">
								<font class="boldtextblack">
									<input type="radio" name="intQ6b" value="1" <%if intq6b = 1 then Response.Write "checked"%>><%=lblYes%> &nbsp;&nbsp;<input type="radio" name="intQ6b" value="2" <%if intq6b = 2 then Response.Write "checked"%>><%=lblNo%>
								</font>
							</td>
						</tr>
						<tr>
							<td>
								<font class="boldtextblack">
									&nbsp;&nbsp;c) <%=lblC%> &nbsp;&nbsp;
								</font>
							</td>
							<td align="center">
								<font class="boldtextblack">
									<input type="radio" name="intQ6c" value="1" <%if intq6c = 1 then Response.Write "checked"%>><%=lblYes%> &nbsp;&nbsp;<input type="radio" name="intQ6c" value="2" <%if intq6c = 2 then Response.Write "checked"%>><%=lblNo%>
								</font>
							</td>
						</tr>
						<tr>
							<td>
								<font class="boldtextblack">
									&nbsp;&nbsp;d) <%=lblD%> &nbsp;&nbsp;
								</font>
							</td>
							<td align="center">
								<font class="boldtextblack">
									<input type="radio" name="intQ6d" value="1" <%if intq6d = 1 then Response.Write "checked"%>><%=lblYes%> &nbsp;&nbsp;<input type="radio" name="intQ6d" value="2" <%if intq6d = 2 then Response.Write "checked"%>><%=lblNo%>
								</font>
							</td>
						</tr>
						<tr>
							<td>
								<font class="boldtextblack">
									&nbsp;&nbsp;e) <%=lblE%> &nbsp;&nbsp;
								</font>
							</td>
							<td align="center">
								<font class="boldtextblack">
									<input type="radio" name="intQ6e" value="1" <%if intq6e = 1 then Response.Write "checked"%>><%=lblYes%> &nbsp;&nbsp;<input type="radio"  name="intQ6e" value="2" <%if intq6e = 2 then Response.Write "checked"%>><%=lblNo%>
								</font>
							</td>
						</tr>
						<tr>
							<td>
								<font class="boldtextblack">
									&nbsp;&nbsp;f) <%=lblF%> &nbsp;&nbsp;
								</font>
							</td>
							<td align="center">
								<font class="boldtextblack">
									<input type="radio" name="intQ6f" value="1" <%if intq6f = 1 then Response.Write "checked"%>><%=lblYes%> &nbsp;&nbsp;<input type="radio"  name="intQ6f" value="2" <%if intq6f = 2 then Response.Write "checked"%>><%=lblNo%>
								</font>
							</td>
						</tr>
						<tr>
							<td>
								<font class="boldtextblack">
									&nbsp;&nbsp;g) <%=lblG%> &nbsp;&nbsp;
								</font>
							</td>
							<td align="center">
								<font class="boldtextblack">
									<input type="radio" name="intQ6g" value="1" <%if intq6g = 1 then Response.Write "checked"%>><%=lblYes%> &nbsp;&nbsp;<input type="radio"  name="intQ6g" value="2" <%if intq6g = 2 then Response.Write "checked"%>><%=lblNo%>
								</font>
							</td>
						</tr>
						<tr>
							<td>
								<font class="boldtextblack">
									&nbsp;&nbsp;h) <%=lblH%> &nbsp;&nbsp;
								</font>
							</td>
							<td align="center">
								<font class="boldtextblack">
									<input type="radio" name="intQ6h" value="1" <%if intq6h = 1 then Response.Write "checked"%>><%=lblYes%> &nbsp;&nbsp;<input type="radio"  name="intQ6h" value="2" <%if intq6h = 2 then Response.Write "checked"%>><%=lblNo%>
								</font>
							</td>
						</tr>
						<tr>
							<td>
								<font class="boldtextblack">
									&nbsp;&nbsp;i) <%=lblI%> &nbsp;&nbsp;
								</font>
							</td>
							<td align="center">
								<font class="boldtextblack">
									<input type="radio" name="intQ6i" value="1" <%if intq6i = 1 then Response.Write "checked"%>><%=lblYes%> &nbsp;&nbsp;<input type="radio"  name="intQ6i" value="2" <%if intq6i = 2 then Response.Write "checked"%>><%=lblNo%>
								</font>
							</td>
						</tr>
						<tr>
							<td>
								<font class="boldtextblack">
									&nbsp;&nbsp;j) <%=lblJ%> &nbsp;&nbsp;
								</font>
							</td>
							<td align="center">
								<font class="boldtextblack">
									<input type="radio" name="intQ6j" value="1" <%if intq6j = 1 then Response.Write "checked"%>><%=lblYes%> &nbsp;&nbsp;<input type="radio"  name="intQ6j" value="2" <%if intq6j = 2 then Response.Write "checked"%>><%=lblNo%>
								</font>
							</td>
						</tr>
						<tr>
							<td>
								<font class="boldtextblack">
									&nbsp;&nbsp;k) <%=lblK%> &nbsp;&nbsp;
								</font>
							</td>
							<td align="center">
								<font class="boldtextblack">
									<input type="radio" name="intQ6k" value="1" <%if intq6k = 1 then Response.Write "checked"%>><%=lblYes%> &nbsp;&nbsp;<input type="radio"  name="intQ6k" value="2" <%if intq6k = 2 then Response.Write "checked"%>><%=lblNo%>
								</font>
							</td>
						</tr>
					</table>
				</td>
			</tr>
			</table>
			<hr>
			<br />
			<%
			if Request.Form("hiddenAction") <> "" then  
				Response.Write "<p align=""center"">" & strReminder & "</p>"
			end if 
			
			%>
			<table border="1" cellpadding="0" cellspacing="0" width="750" align="center">
				<tr>
					<td align="center"><font class="boldtextblack"><%=lblID%></font></td>
					<td align="center"><font class="boldtextblack"><%=lblLanguage%></font></td>
					<td align="center"><font class="boldtextblack"><%=lblSize%></font></td>
					<td align="center"><font class="boldtextblack"><%=lblComplete%></font></td>
				</tr>
				<%
								
				if intClasses = 0 then 
					Response.Write "<tr><td colspan=""7"">&nbsp;<font class=""regtextmaroon"">This teacher has no classes.</font></td></tr>"
				else
					do while not rstData.EOF 
						Response.Write "<tr><td><a href=""edi_teacher_class.asp?teacher=" & left(right("000" & rstData("intClassID"),9),8) & "&class=" & right(rstData("intClassID"),1) & """ class=""reglinkBlue"">" & right("000" & rstData("intClassID"),9) & "</a></td>"
						Response.Write "<td><font class=""regtextblack"">" & rstData("strLanguage") & "</font></td>"
						Response.Write "<td align=""center""><font class=""regtextblack"">" & rstData("intStudents") & "</font></td>"
						Response.Write "<td align=""center""><font class=""regtextblack"">" & rstData("Completed") & "</font></td>"
						rstData.MoveNext 
					loop
				end if 
				%>
			</table>
			<br />
			</td>
		</tr>
		</table>
	</form>
	<%'end if%>	
	<!-- #include virtual="/shared/page_footer.inc" -->
</body>
</html>
<%
	' close and kill recordset and connection
	call close_adodb(rstData)
	'call close_adodb(conn)
	call close_adodb(conn)
end if

' set form defaults
sub add_mode()
	' load the first school
	strName = ""
'	strAddress = ""
	strCity = ""
	intProvince = 1
	strPostal = ""
	strPhone = "" 
	strFax = ""
	strEmail = "" 
	strComments = ""
	intQ5a = 0
	intQ5b = 0
	intQ5c = 0
	Question1 = 0
	Question2	= 0
	Question3	= 0
	Question4	= 0
	Question5	= 0
	Question6	= 0
	Question7YH	= 0
	Question7YNH	= 0
	Question7NNH	= 0
	Question7NNone	= 0
	Question7NTime	= 0
	Question7NFamiliar	= 0
	Question7Other	= 0
	Question7OtherText		= 0
	StudentsInClass = 0
end sub

sub load_values(intTeacher)	
	if intTeacher = 0 then 
		introw = 0 
	else
		for introw = 0 to ubound(aData,2)
			if clng(intTeacher) = aData(0,introw) then 
				exit for
			end if
		next

		if intRow > ubound(aData,2) then 
			intRow = 0 
		end if 	
	end if 
	
	strName = aData(2,intRow)
	strEmail = aData(3,intRow)
	strPassword = aData(4,introw)
	strPhone = aData(5,introw)
	strFax = aData(6,introw) 
	intSex = aData(7,introw) 
	intAge = aData(8,introw)
	intQ5a = aData(9,introw)
	intyr1 = int(intQ5a / 12)
	intMth1 = intQ5a mod 12
	
	intQ5b = aData(10,introw)
	intyr2 = int(intQ5b / 12)
	intMth2 = intQ5b mod 12
	
	intQ5c = aData(11,introw)
	intyr3 = int(intQ5c / 12)
	intMth3 = intQ5c mod 12
	
	intQ6a = aData(13,introw) 
	intQ6b = aData(14,introw) 
	intQ6c = aData(15,introw) 
	intQ6d = aData(16,introw) 
	intQ6e = aData(17,introw) 
	intQ6f = aData(18,introw) 
	intQ6g = aData(19,introw) 
	intQ6h = aData(20,introw) 
	intQ6i = aData(21,introw) 
	intQ6j = aData(22,introw) 
	intQ6k = aData(23,introw)
end sub

sub load_ParticipationValues(intTeacher)	
	if intTeacher = 0 then 
		introw = 0 
	else
		for introw = 0 to ubound(aData,2)
			if clng(intTeacher) = aData(1,introw) then 
				exit for
			end if
		next

		if intRow > ubound(aData,2) then 
			intRow = 0 
		end if 	
	end if 
	
'	strName = aData(0,intRow)	
	Question1 = aData(2,intRow)
	Question2	= aData(3,intRow)
	Question3	= aData(4,intRow)
	Question4	= aData(5,intRow)
	Question5	= aData(6,intRow)
	Question6	= aData(7,intRow)
	Question7YH	= aData(8,intRow)
	Question7YNH	= aData(9,intRow)
	Question7NNH	= aData(10,intRow)
	Question7NNone	= aData(11,intRow)
	Question7NTime	= aData(12,intRow)
	Question7NFamiliar	= aData(13,intRow)
	Question7Other	= aData(14,intRow)
	Question7OtherText		= aData(15,intRow)
    StudentsInClass = aData(16,intRow)
end sub
%>