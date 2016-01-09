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
		strSql = "UPDATE admins SET strName = " & checknull(Request.form("strName")) & ", strEmail = " & checknull(Request.Form("strEmail")) & ", strPassword = " & checknull(Request.Form("strPassword")) & " WHERE strEmail = '" & session("id") & "'"
		conn.execute strSql 
		
		if conn.errors.count > 0 then
			strError = "<font class=""regtextred"">Error Number : " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/></font>"
		else
			strError = "<font class=""regtextgreen"">Changes made successfully.<br /><br /></font>"
		end if 
	end if 
	
	' find the record with the email address in it	
	strSql = "SELECT strName, strEmail, strPassword FROM users WHERE strEmail ='" & session("id") & "'"	
	
	' open the recordset
	rstEmail.Open strSql, conn
	if not rstEmail.EOF then 
		strName = rstEmail("strName")
		strEmail = rstEmail("strEmail")
		strPassword = rstEmail("strPassword")
	end if 
%>
<html>
<head>
<!--<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />-->
	<!-- added UTF8 Encoding to get rid of funny characters -->
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" /> 

	<!-- Start CSS files-->
		<link rel="stylesheet" type="text/css" href="Styles/edi.css">
	<!-- End CSS files -->
</head>
<body>
	<!-- #include virtual="/shared/page_header.inc" -->
	<form name="Passwords" method="POST" action="edi_admin_profile.asp"> 
		<a class="reglinkMaroon" href="edi_admin.asp">Home</a>&nbsp;<font class="regtextblack">></font>&nbsp;<font class="boldtextblack">Personal Profile</font>
		<table border="1" bordercolor="006600" cellpadding="0" cellspacing="0" width="760">
		<tr>
			<td>
				<table border="0" cellpadding="0" cellspacing="0" width="760">
					<tr>
						<td align="right" width="430">
							<font class="headerBlue">Personal Profile</font>
						</td>
						<td align="right">
							<input type="submit" name="Save" value="Save">
							<input type="button" value="Exit" name="Exit" title="EXIT Screen" onClick="javascript:window.location='edi_admin.asp';">
							&nbsp;
						</td>
					</tr>
					<tr><td><br/></td></tr>
				</table>
				<table border="0" cellpadding="0" cellspacing="0" width="760" align="center">
					<tr><td colspan="2"><%=strError%></td></tr>
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
							<input type="text" name="strEmail" size="25" value="<%=strEmail%>">
						</td>
					</tr>
					<tr>
						<td align="right">
							<font class="boldTextBlack">Password: &nbsp;</font>
						</td>
						<td  align="left">
							<input type="text" name="strPassword" size="25" value="<%=strPassword%>">
						</td>
					</tr>
					<tr><td><br/></td></tr>
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
