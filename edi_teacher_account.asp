<!-- #include virtual="/shared/security.asp" -->
<%
' if the user has not logged in they will not be able to see the page
if blnSecurity then
on error resume next
	' open edi connection
	'call open_adodb(conn, "DATA")
    call open_adodb(conn, "MACEDI")
	set rstEmail = server.CreateObject("adodb.recordset")
	
	' update the profile
	if Request.Form("strName").Count > 0 then 
		strSql = "UPDATE teachers SET strName = " & checknull(Request.form("strName")) & ", strEmail = " & checknull(Request.Form("strEmail")) & ", strPassword = " & checknull(Request.Form("strPassword")) & " WHERE strEmail = '" & session("id") & "'"
		conn.execute strSql 
		
		if conn.errors.count > 0 then
			strError = "<font class=""regtextred"">Error Number : " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/></font>"
		else
			strError = "<font class=""regtextgreen"">"
			
			if session("language") = "French" Then 
				strError = strError & "Les changements ont fait avec succès." 
			else
				strError = strError & "Changes made successfully.<br /><br /></font>"
			end if 
		end if 
	end if 
	
	' find the record with the email address in it	
	strSql = "SELECT strName, strEmail, strPassword FROM teachers WHERE strEmail ='" & session("id") & "'"	
	
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
    <!-- added UTF8 Encoding to get rid of funny characters -->
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" /> 
	<!-- Start CSS files-->
		<link rel="stylesheet" type="text/css" href="Styles/edi.css">
	<!-- End CSS files -->
</head>
<body>
	<!-- #include virtual="/shared/page_header.inc" -->
	<form name="Passwords" method="POST" action="edi_teacher_account.asp">
		<a class="reglinkMaroon" href="edi_teacher.asp"><%=strHome%></a>&nbsp;<font class="regtextblack">></font>&nbsp;<font class="boldtextblack"><%=lblPass%></font>
		<table border="1" bordercolor="006600" cellpadding="0" cellspacing="0" width="760">
		<tr>
			<td>
				<table border="0" cellpadding="0" cellspacing="0" width="760">
					<tr>
						<td align="right" width="430">
							<font class="headerBlue"><%=lblPass%></font>
						</td>
						<td align="right">
							<input type="submit" name="<%=strSave%>" value="<%=strSave%>">
							<input type="button" value="<%=strExit%>" name="<%=strExit%>" title="EXIT Screen" onClick="javascript:window.location='edi_teacher.asp';">
							&nbsp;
						</td>
					</tr>
					<tr><td><br/></td></tr>
				</table>
				<table border="0" cellpadding="0" cellspacing="0" width="760" align="center">
					<tr><td colspan="2"><%=strError%></td></tr>
					<tr><td colspan="2" align="center"><font class="subheaderBlue"><%=strMsg%></font><br /><br /></td></tr>
					<tr>
						<td align="right">
							<font class="boldTextBlack"><%=lblName%>: &nbsp;</font>
						</td>
						<td align="left">
							<input type="text" name="strName" size="25" value="<%=strName%>">
						</td>
					</tr>
					<tr>
						<td align="right">
							<font class="boldTextBlack"><%=lblEmail%>: &nbsp;</font>
						</td>
						<td align="left">
							<input type="text" name="strEmail" size="25" value="<%=strEmail%>">
						</td>
					</tr>
					<tr>
						<td align="right">
							<font class="boldTextBlack"><%=lblPassword%>: &nbsp;</font>
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
