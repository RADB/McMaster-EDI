<!-- #include virtual="/shared/security.asp" -->
<%
' public variables
dim aData
dim intOptions
dim strEDIID
dim aHeader(6)
' variables from load data
dim intQ1a,intQ1b,intQ1c,intQ1d,intQ1e,intQ1f,intQ1g,intQ1h,strQ1hOther,intQ2a,intQ2b,intQ2c,intQ2d,intQ2e,strQ2eOther,intQ3a,intQ3b,intQ3c,intQ3d,intQ3e,intQ3f,intQ3g,intQ3h,strQ3hOther,modifiedUser,modifiedDate
on error resume next

' if the user has not logged in they will not be able to see the page
if blnSecurity then 
	'call open_adodb(conn, "DATA")
	call open_adodb(conn, "MACEDI")
	
	set rstData = server.CreateObject("adodb.recordset")
	set rstConfig = server.CreateObject("adodb.recordset")	
	
	if Request.Form("Action") = "Update" then		
		strEDIID = Request.Form("EDIID")
		
		strSQL = "UPDATE NWT_Identity SET " 
		for each item in Request.Form 
			if NOT (item = "frmAction" OR item = "frmSection" OR item = "frmSite" OR item ="frmSchool" OR item ="frmTeacher" OR item ="frmClass" OR item = "frmChild" OR item = "frmNextChild" OR item = "btnSave" OR item = "hdnLock" OR item = "hdnCheckBoxes" OR item = "Student" OR item = "classes" OR item = "email" OR item ="rpt" OR item ="XML" or item="CurrentSection" or item="EDIID" or item="Action") then 
				blnUpdate = true
				if left(item ,3) = "str" then
					strSql = strSql & item & " = " & checknull(Request.Form(item)) & ","
				else
					strSql = strSql & item & " = " & Request.Form(item) & ","
				end if 
			end if 
		next
		
		' remove the last comma
		strSql = left(strSql,len(strSql)-1) & ", modifiedDate = '" & now() & "', modifieduser = '" & session("id") & "' WHERE strEDIID = '" & strEDIID & "'"
		
		'Response.Write strSql
		' update the record
		conn.execute strSql
		
		' build the error string
		if conn.errors.count > 0 then 
			strError = "<font class=""regtextred"">Error Number : " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/></font>"
		end if
	else
		strEDIID = request.form("frmEDIYear") & Request.Form("frmSite") & Request.Form("frmSchool") & Request.Form("frmTeacher") & Request.Form("frmClass") & Request.Form("frmChild")
	end if 
	
    ' get the student identity data
	strSql = "SELECT intQ1a,intQ1b,intQ1c,intQ1d,intQ1e,intQ1f,intQ1g,intQ1h,strQ1hOther,intQ2a,intQ2b,intQ2c,intQ2d,intQ2e,strQ2eOther,intQ3a,intQ3b,intQ3c,intQ3d,intQ3e,intQ3f,intQ3g,intQ3h,strQ3hOther,modifiedUser,modifiedDate FROM NWT_Identity WHERE strEDIID = '" & strEDIID & "'"
	rstData.Open strSql, conn  
							
	if rstData.EOF then
		intChild = 0
		conn.execute "INSERT INTO NWT_Identity (strEDIID, modifieduser) VALUES ('" & strEDIID & "','" & session("id") &"')"	
		rstData.close
		rstData.Open strSql, conn
	end if 

	' get all the identity questions
	strSql = "SELECT I.English, I.French, I.isSectionHeader, I.[Section], I.[Question], I.[Option], I.[isHeader],  Col.[Language], Col.intOptions, Col.Col1, Col.Col2, Col.Col3, Col.Col4, Col.Col5, Col.Col6 FROM Page_Identity I  LEFT JOIN [Column Headers] Col ON I.HeaderType = Col.HID WHERE Col.[Language]='" & strLanguage & "' OR Col.[Language] Is Null ORDER BY I.[Section], I.[Sequence]"
								
	'open the Section A questions 
	rstConfig.Open strSql, conn
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
	<form name="Identity" method="POST" action="edi_teacher_identity.asp"> 
	    <!-- breadcrumb-->
		<a class="reglinkMaroon" href="edi_teacher.asp"><%=strHome%></a>&nbsp;<font class="regtextblack">></font>&nbsp;<a class="reglinkMaroon" href="edi_teacher_class.asp"><%=lblClassCrumb%></a>&nbsp;<font class="regtextblack">></font>&nbsp;<font class="boldtextblack"><%=lblIdentity%></font>
		<table border="1" bordercolor="006600" cellpadding="0" cellspacing="0" width="760">
		<tr><td>
			<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">
			<tr>
				<td align="right" width="450"><font class="headerBlue"><%=lblIdentity & " (" & strEDIID & ")"%></font></td>
				<td align="right">
					<input type="hidden" name="Action" value="">
					<input type="hidden" name="EDIID" value="<%=strEDIID%>">
					<input type="button" value="<%=strSaveIdentity%>" name="Update"  onClick="javascript:goSaveIdentity();">
					<input type="hidden" name="strLanguageCompleted" value="<%=session("language")%>" />	
					<input type="button" value="<%=strExit%>" name="Exit" onClick="javascript:window.location='edi_teacher_class.asp';">
					&nbsp;
				</td>
			</tr>
			<tr><td colspan="2"><%=strError%></td></tr>
			<!-- sections here-->
			</table>
			
			<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">
				<tr valign="top">
					<td align="left">
						<br />
						<font class="subheaderBlue"><%=lblIdentitySubHeader%>&nbsp;&nbsp;</font>
					</td>
				</tr>
			</table>
			
			<br />

			<%							
			' bln for inner table existance
			blnTable = false
			intOptions = 0 
			intRow = 0

			Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""750"" align=""left"">"																					
			do while not rstConfig.EOF 				
				if rstConfig("isSectionHeader") then 
					if blnTable then 
						intRow = 0
						Response.Write "</table>"
						Response.Write "<tr><td><br /></td></tr>"
					end if				
					Response.Write "<tr><td>"						
						Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""750"">"
							Response.Write "<tr><td>"						
							response.write "<font class=""subheaderBlue"">" & rstConfig("Question") & ")&nbsp;&nbsp;" & rstConfig(session("language")) & "</font>"									
							Response.Write "</td></tr>"					
						Response.Write "</table>"
					Response.Write "</td></tr>"						
				else
					' write the header row if it is a header			
					if rstConfig("isHeader") then 
						' check to see if the last table needs to be closed
													
						' new section
						Response.Write "<tr><td>"
						
						' get the number of options 
						intOptions = rstConfig("intOptions")
						if intOptions > 0 then 										
							' inner table for each section
							Response.Write "<table border=""1"""
						else
							Response.Write "<table border=""0"""
						end if 
						
						response.write " cellpadding=""0"" cellspacing=""0"" width=""750"">"																					
						Response.Write "<tr>"		
						Response.Write "<td align=""left"" colspan=""2""><font class=""subHeaderBlue"">" & rstConfig(strLanguage) & "</font></td>" 
						
						if intOptions > 0 then 																		
							for intCol = 1 to intoptions
								Response.Write "<td align=""center"" valign=""middle"" width=""70""><font class=""boldTextBlack"">" & rstConfig("Col" & intCol) & "</font></td>"
								' store the header values
								aHeader(intCol) = rstConfig("Col" & intCol)								
							next 
						end if 
									
						Response.Write "</tr>"
													
						' set the inner table to true
						blnTable = true 
					else
						' if it is just a question and no header then write a new table
						if intOptions = 0 then 
							Response.Write "<tr><td>"	
							Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""750"" align=""center"">"																					
								Response.Write "<tr><td align=""left"" valign=""top""><font class=""boldTextBlack"">&nbsp;&nbsp;" 
								' only write the questions
								if rstConfig("option") = 0 then 
									if rstConfig("question") < 10 then 
										Response.Write "&nbsp;&nbsp;"
									end if 
									Response.Write rstConfig("question") & " ) &nbsp;&nbsp;</font></td><td align=""left""><font class=""boldTextBlack"">" &  rstConfig(strLanguage) & "</font>"
								end if
																			
								Response.Write "</td>"
								Response.Write "</tr>"
							Response.Write "</table>"
							Response.Write "</td></tr>"
							Response.Write "<tr><td><br /></td></tr>"
						else
							intRow = introw + 1
							if intRow mod 2 = 1 then 
								strColour = "whitesmoke"
							else
								strColour = "white"
							end if 
							Response.Write "<tr bgcolor=""" & strColour & """><td align=""left"" valign=""top"">&nbsp;<font class=""boldTextBlack"">" 
							' only write the questions
							if rstConfig("option") > 0 then 
								Response.Write chr(rstConfig("option")) & " ) &nbsp;&nbsp;</font></td><td><font class=""boldTextBlack"">" &  rstConfig(strLanguage) & "</font>"
								strQuestion = "intQ" & rstConfig("Question") & chr(rstConfig("option"))	

								if (rstConfig("Question") = 1 and rstConfig("option") = 104) or (rstConfig("Question") = 2 and rstConfig("option") = 101) or (rstConfig("Question") = 3 and rstConfig("option") = 104) then 
									strQuestion2 = "strQ" & rstConfig("Question") & chr(rstConfig("option")) & "Other"
									response.write "<input type=""text"" maxlength=""50"" name=""" & strQuestion2 & """ size=""45"" value=""" & rstdata(strQuestion2) & """ />"
								end if
							end if																			
																	
							Response.Write "</td>"
												
							' write the radio option buttons
							for intCol = 1 to intoptions
								Response.Write "<td align=""center"" valign=""middle"" width=""70""><input type=""radio"" title=""" & aHeader(intCol) &""" name=""" & strQuestion &  """ value=""" & intCol & """"
								if rstData(strQuestion) = intCol then 
									response.write " checked=""checked"""
								end if 
								response.write " /></td>"
							next 
							Response.Write "</tr>"
						end if 										
					end if 
				end if
				rstConfig.movenext			
			loop
			%>
			<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">
			<tr>
				<td align="right" width="450"></td>
				<td align="right">				
					<br />
					<input type="button" value="<%=strSaveIdentity%>" name="Update"  onClick="javascript:goSaveIdentity();">
					<input type="button" value="<%=strExit%>" name="Exit" onClick="javascript:window.location='edi_teacher_class.asp';">
					&nbsp;
				</td>
			</tr>
			<!-- sections here-->
			</table>
			<%
		' close the last inner table if it is open
		if blnTable then 
			Response.Write "</table>"
		end if 
									
		' end the row
		Response.Write "</td></tr>"
		Response.Write "</table>"
		%>
			
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
	call close_adodb(rstConfig)
	call close_adodb(rstData)
	'call close_adodb(conn)
	call close_adodb(conn)
end if
%>