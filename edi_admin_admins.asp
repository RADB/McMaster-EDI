<!-- #include virtual="/shared/admin_security.asp" -->
<%
' if the user has not logged in they will not be able to see the page
if blnSecurity then
on error resume next
	' open edi connection
	'call open_adodb(conn, "EDI")
    call open_adodb(conn, "MACEDI")
	set rstEmail = server.CreateObject("adodb.recordset")
	
	' update the profile
	if Request.Form("strName").Count > 0 then 
		if Request.Form("Action") = "Update" then 
			strSql = "UPDATE admins SET strName = " & checknull(Request.form("strName")) & ", intAccess = " & Request.form("intSecurity") & " WHERE strEmail = '" & Request.Form("strEmail") & "'"
			'Response.Write strSql
			conn.execute strSql 
		
			if conn.errors.count > 0 then
				strError = "<font class=""regtextred"">Error Number : " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/></font>"
			else
				strError = "<font class=""regtextgreen"">Changes made successfully.<br /><br /></font>"
			end if 
		elseif Request.Form("Action") = "Delete" then 
			strSql = "DELETE FROM admins WHERE strEmail = '" & Request.Form("strEmail") & "'"
			conn.execute strSql 
			
			if conn.errors.count > 0 then
				strError = "<font class=""regtextred"">Error Number : " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/></font>"
			else
				strError = "<font class=""regtextgreen"">Changes made successfully.<br /><br /></font>"
			end if 
		elseif Request.Form("Action") = "Save" then 
			strSql = "INSERT INTO admins (strName,strEmail, strPassword, intAccess) VALUES(" & checkNull(Request.Form("strName")) & "," & checkNull(Request.Form("strEmail")) & "," & checkNull(Request.Form("strPassword")) & "," & checkNull(Request.Form("intAccess")) & ")"
			conn.execute strSql
			
			if conn.errors.count > 0 then
				strError = "<font class=""regtextred"">Error Number : " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/></font>"
			else
				strError = "<font class=""regtextgreen"">Changes made successfully.<br /><br /></font>"
			end if 
		end if 
	end if 
	
	' find the record with the email address in it	
	strSql = "SELECT strName, strEmail, strPassword, intAccess FROM users WHERE strType =0"	
	
	' open the recordset
	rstEmail.Open strSql, conn
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
</head>
<body>
	<!-- #include virtual="/shared/page_header.inc" -->
	<form name="Screens" method="POST" action="edi_admin_admins.asp"> 
		<input type="hidden" name="Action" value="">
		<a class="reglinkMaroon" href="edi_admin.asp">Home</a>&nbsp;<font class="regtextblack">></font>&nbsp;<font class="boldtextblack">Admin Management</font>
		<table border="1" bordercolor="006600" cellpadding="0" cellspacing="0" width="760">
		<tr>
			<td>
				<table border="0" cellpadding="0" cellspacing="0" width="760">
					<tr>
						<td align="right" width="430">
							<font class="headerBlue">Admin Management</font>
						</td>
						<td align="right">
							<%
							if Request.form("Action") <> "Add" OR Request.form("Action").Count = 0 then  
							%>
								<input type="button" value="Add" name="SubmitAction" title="ADD ADMIN" onClick="javascript:confirm_Add(this.value);">
								<input type="button" value="Delete" name="SubmitAction" title="DELETE ADMIN" onClick="javascript:confirm_Delete(this.value);">
								<input type="button" value="Update" name="SubmitAction" title="UPDATE ADMIN" onClick="javascript:admin_Check(this.value);">
							<%
							else
							%>
									<input type="button" value="Cancel" name="Cancel" title="Cancel Add Mode" onClick="javascript:window.location='edi_admin_admins.asp';">
									<input type="button" value="Save" name="SubmitAction" title="SAVE ADMIN" onClick="javascript:admin_Check(this.value);">
							<%
							end if 
							%>
							<input type="button" value="Exit" name="Exit" title="EXIT Screen" onClick="javascript:window.location='edi_admin.asp';">
							&nbsp;
						</td>
					</tr>
					<tr><td><br/></td></tr>
				</table>
				<table border="0" cellpadding="0" cellspacing="0" width="760" align="center">
					<tr><td colspan="2"><%=strError%></td></tr>
					<%
					if Request.form("Action") <>  "Add" then 					
					%>
						<tr>
							<td align="right">
								<font class="boldTextBlack">Admin: &nbsp;</font>
							</td>
							<td align="left">
								<select name="intName" onChange="javascript:changeAdmin(this.value);">
									<%
									intCount = 0
									do while not rstEmail.EOF
										Response.Write "<option value=""" & intCount & """"
										if (Request.Form("intName").Count = 0 and intCount = 0) OR (Request.Form("Action") = "Delete" and intCount = 0) then
											strName = rstEmail("strName")
											strEmail = rstEmail("strEmail")	
											strPassword = rstEmail("strPassword")
											intSecurity = rstEmail("intAccess")
											
											Response.Write " selected"
										elseif (cint(Request.Form("intName")) = intCount and Request.Form("Action") <> "Delete") OR (Request.Form("Action") = "Save" and Request.Form("strName") = rstEmail("strName")) then 	
											strName = rstEmail("strName")
											strEmail = rstEmail("strEmail")	
											strPassword = rstEmail("strPassword")
											intSecurity = rstEmail("intAccess")
											
											Response.Write " selected"
										end if 
										Response.Write ">" & rstEmail("strName") & "</option>"
										intCount = intCount + 1
										rstEmail.MoveNext 
									loop
									%>
								</select>
							</td>
						</tr>
					<%
					end if
				   %>
				   <tr>
						<td align="right">
							<font class="boldTextBlack">Name: &nbsp;</font>
						</td>
						<td align="left">
							<input type="text" name="strName" size="25" value="<%=strName%>">
						</td>
					</tr>
					<tr>
						<td align="right">
							<font class="boldTextBlack">Email: &nbsp;</font>
						</td>
						<td align="left">
							<input type="text" name="strEmail" size="25" value="<%=strEmail%>"
							<%
							if Request.form("Action") <>  "Add" then 					
								Response.write " readonly"
							end if 
							%>
							>
						</td>
					</tr>
					<tr>
						<td align="right">
							<font class="boldTextBlack">Password: &nbsp;</font>
						</td>
						<td  align="left">
							<input name="strPassword" size="25" value="<%=strPassword%>"
							<%
							if Request.form("Action") <>  "Add" then 					
								Response.write " readonly type=""password"""
							else
								Response.write " type=""text"""
							end if 
							%>
							>	
						</td>
					</tr>
					<tr>
						<td align="right">
							<font class="boldTextBlack">Security Access: &nbsp;</font>
						</td>
						<td  align="left">
							<select name="intSecurity">
								<%
									Response.Write "<option value=""0"""
									if intSecurity = 0 then 
										Response.Write " selected"
									end if 
									Response.Write ">Deny</option>"
									Response.Write "<option value=""1"""
									if intSecurity = 1 then 
										Response.Write " selected"
									end if 
									Response.Write ">Grant</option>"
								%>
							</select>
						</td>
					</tr>
					<tr><td><br/></td></tr>
					<%
					if Request.form("Action") <>  "Add" then 					
					%>
					<tr>
						<td align="center" colspan="2">
							<font class="subheaderblue">Please note that you can only change the name and security access.  You cannot alter the email or password.</font>
						</td>
					</tr>
					<tr><td><br/></td></tr>
					<%
					end if 
					%>
				</table>
			</td>			
		</tr>
		</table>	
		<br/> 
	
	</form>
	
	<!-- #include virtual="/shared/page_footer.inc" -->
</body>
</html>
<%
	' close and kill the connection and recordset
	call close_adodb(rstEmail)
	call close_adodb(conn)
end if
%>
