<!-- #include virtual="/shared/security.asp" -->
<%
' public variables
dim aData
dim strName,Question1,Question2,		Question3	,	Question4	,	Question5	,	Question6	,	Question7YH	,	Question7YNH	,	Question7NNH	,	Question7NNone	,	Question7NTime	,	Question7NFamiliar,		Question7Other	,	Question7OtherText	
on error resume next

' if the user has not logged in they will not be able to see the page
if blnSecurity then
	'call open_adodb(conn, "DATA")
	call open_adodb(conn, "MACEDI")
	
	set rstData = server.CreateObject("adodb.recordset")
	
	if Request.Form("Action") = "Update" then
		intTeacher = Request.Form("code") 
			
		' build the SQL statement		
		 strSQL = "UPDATE  teacherParticipation SET Question1 = " & checkNull(Request.Form("Question1")) & ", Question2 = " & checkNull(Request.Form("Question2")) & ", Question3 = " & checkNull(Request.Form("Question3")) & ", Question4 = " & checkNull(Request.Form("Question4")) & ", Question5 = " & checkNull(Request.Form("Question5")) & ", Question6 = " & checkNull(Request.Form("Question6")) & ", Question7YH = " & checkNull(Request.Form("Question7YH")) & ", Question7YNH = " & checkNull(Request.Form("Question7YNH")) & ", Question7NNH = " & checkNull(Request.Form("Question7NNH")) & ", Question7NNone = " & checkNull(Request.Form("Question7NNone")) & ", Question7NTime = " & checkNull(Request.Form("Question7NTime")) & ", Question7NFamiliar = " & checkNull(Request.Form("Question7NFamiliar")) & ", Question7Other = " & checkNull(Request.Form("Question7Other")) & ", Question7OtherText = " & checkNull(Request.Form("Question7OtherText")) & _
                " WHERE intTeacherID = " & intTeacher
		
		'Response.Write strSql
		' update the record
		conn.execute strSql
		
		' build the error string
		if conn.errors.count > 0 then 
			strError = "<font class=""regtextred"">Error Number : " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/></font>"
		end if
	end if 
	
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
					Response.Redirect "edi_teacher_participation.asp"
				end if 
			end if 
		end if
																		
		' load the values
	   call load_values(intTeacher)
	else
		intTeachers = 0
		' only allow add to teachers...
		'call add_mode
	end if 

	' close the recordset
	rstData.Close 
	
	
	
%>
<html>
<head>
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
	<form name="Screens" method="POST" action="edi_teacher_participation.asp"> 
	    <!-- breadcrumb-->
		<a class="reglinkMaroon" href="edi_teacher.asp"><%=strHome%></a>&nbsp;<font class="regtextblack">></font>&nbsp;<font class="boldtextblack"><%=lblPartTitle%></font>
		<table border="1" bordercolor="006600" cellpadding="0" cellspacing="0" width="760">
		<tr><td>
			<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">
			<tr>
				<td align="right" width="430"><font class="headerBlue"><%=lblPartTitle%></font></td>
				<td align="right">
					<input type="hidden" name="Action" value="">
				<%
					if Request.form("Action") <> "Add" AND intTeachers > 0 then  
					%>
						<input type="button" value="<%=lblPartUpdate%>" name="Update"  onClick="javascript:update_TeacherParticipationCheck(this.name);">
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
					<td align="right">
						<font class="boldtextblack"><%=lblPartCode%> :&nbsp;&nbsp;</font>
					</td>
					<td>
					<%
						' teacher						
						Response.Write "<select name=""code"" onChange=""javascript:window.location='edi_teacher_participation.asp?teacher=' + this.value;"">"
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
						<font class="boldtextblack"><%=lblPartName%> :&nbsp;&nbsp;</font>
					</td>
					<td>
						<input type="text" size="80" name="name" readonly="true" value="<%=strName%>">
					</td>
				</tr>
			</table>
			
			<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">
				<tr valign="top">
					<td align="left">
						<br />
						<font class="subheaderBlue"><%=lblPartTitleSubHeader%> :&nbsp;&nbsp;</font>
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
								                <font class="boldtextblack">1) <%=lblPart1%> &nbsp;&nbsp;</font>
							                </td>
							                <td align="center">
								                <font class="boldtextblack">
									                <input type="radio" name="Question1" value="1" <%if Question1 = 1 then Response.Write "checked"%>><%=lblPartYes%> 5 &nbsp;&nbsp;<input type="radio" name="Question1" value="2" <%if Question1 = 2 then Response.Write "checked"%>><%=lblPartNo%>
								                </font>
							                </td>
						                </tr>
						                <tr valign="top">
					                        <td>
						                        <font class="boldtextblack">2) <%=lblPart2%>&nbsp;&nbsp;</font>
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
								                        Response.Write ">" & lblPart4OrMore & "</option>"
							                        %>
						                        </select>									
					                        </td>
				                        </tr>	
				                        <tr>
							                <td>
								                <font class="boldtextblack">3) <%=lblPart3%> &nbsp;&nbsp;</font>
							                </td>
							                <td align="center">
								                <font class="boldtextblack">
									                <input type="radio" name="Question3" value="1" <%if Question3 = 1 then Response.Write "checked"%>><%=lblPartYes%> 4 &nbsp;&nbsp;<input type="radio" name="Question3" value="2" <%if Question3 = 2 then Response.Write "checked"%>><%=lblPartNo%>
								                </font>
							                </td>
						                </tr>			
						                <tr valign="top">
					                        <td>
						                        <font class="boldtextblack">4) <%=lblPart4%>&nbsp;&nbsp;</font>
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
								                        Response.Write ">" & lblPart4OrMore & "</option>"
							                        %>
						                        </select>									
					                        </td>
				                        </tr>	
				                        <tr>
							                <td>
								                <font class="boldtextblack">5) <%=lblPart5%> &nbsp;&nbsp;</font>
							                </td>
							                <td align="center">
								                <font class="boldtextblack">
									                <input type="radio" name="Question5" value="1" <%if Question5 = 1 then Response.Write "checked"%>><%=lblPartYes%> 6 &nbsp;&nbsp;<input type="radio" name="Question5" value="2" <%if Question5 = 2 then Response.Write "checked"%>><%=lblPartNo%>
								                </font>
							                </td>
						                </tr>			
				                        <tr valign="top">
					                        <td>
						                        <font class="boldtextblack">6) <%=lblPart6%>&nbsp;&nbsp;</font>
					                        </td>
					                        <td>					                            
						                        <select name="Question6">							                        
						                            <option value=""></option>
							                        <%
							                            Response.Write "<option value=""1"""
								                        if Question6 = "l" then Response.Write " selected"
								                        Response.Write ">" & lblPartVery & "</option>"
								                        Response.Write "<option value=""2"""
								                        if Question6 = "2" then Response.Write " selected"
								                        Response.Write ">" & lblPartSomewhat & "</option>"
								                        Response.Write "<option value=""3"""
								                        if Question6 = "3" then Response.Write " selected"
								                        Response.Write ">" & lblPartNotatall & "</option>"								                        
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
						<font class="subheaderBlue"><%=lblPart7%> :&nbsp;&nbsp;</font>
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
			                            <input type="checkbox" name="Question7YH" value="1" <%if Question7YH = 1 then Response.Write "checked"%>><%=lblPartQuestion7YH%> 			                            
		                            </font>
	                            </td>
	                            <td align="left">
		                            <font class="boldtextblack">
			                            <input type="checkbox" name="Question7NNH" value="1" <%if Question7NNH = 1 then Response.Write "checked"%>><%=lblPartQuestion7NNH%> &nbsp;&nbsp;			                            
		                            </font>
	                            </td>
	                        </tr>	            
	                        <tr>
				                <td align="left">
		                            <font class="boldtextblack">			                            
			                            <input type="checkbox" name="Question7YNH" value="1" <%if Question7YNH = 1 then Response.Write "checked"%>><%=lblPartQuestion7YNH%> &nbsp;&nbsp;			                            
		                            </font>
	                            </td>
	                            <td align="left">
		                            <font class="boldtextblack">			                            
			                            <input type="checkbox" name="Question7NNone" value="1" <%if Question7NNone = 1 then Response.Write "checked"%>><%=lblPartQuestion7NNone%> &nbsp;&nbsp;
		                            </font>
	                            </td>
	                        </tr>	            
	                        <tr>
				                <td align="left">
		                            <font class="boldtextblack">
			                            <input type="checkbox" name="Question7Other" value="1" <%if Question7Other = 1 then Response.Write "checked"%>><%=lblPartQuestion7Other%> &nbsp;&nbsp;
		                            </font>
	                            </td>
	                            <td align="left">
		                            <font class="boldtextblack">			                            
			                            <input type="checkbox" name="Question7NTime" value="1" <%if Question7NTime = 1 then Response.Write "checked"%>><%=lblPartQuestion7NTime%> &nbsp;&nbsp;
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
			                            <input type="checkbox" name="Question7NFamiliar" value="1" <%if Question7NFamiliar= 1 then Response.Write "checked"%>><%=lblPartQuestion7NFamiliar%> &nbsp;&nbsp;
		                            </font>
	                            </td>
	                        </tr>	            
	                    </table>
	                 </td>       
	            </tr>
	        </table>	
			<hr />
			<br />
			<%
			if Request.Form("hiddenAction") <> "" then  
				Response.Write "<p align=""center"">" & strReminder & "</p>"
			end if 
			
			%>						
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

sub load_values(intTeacher)	
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
	
	strName = aData(0,intRow)	
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
end sub
%>