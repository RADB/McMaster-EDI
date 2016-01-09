<!-- #include virtual="/shared/admin_security.asp" -->
<%
' if the user has not logged in they will not be able to see the page

if blnSecurity then
	' open edi connection
	'call open_adodb(conn, "EDI")
	'call open_adodb(conn, "DATA")
    call open_adodb(conn, "MACEDI")
	set rstEmail = server.CreateObject("adodb.recordset")
	set rstData = server.CreateObject("adodb.recordset")
	dim strMessage
	dim strMessage2
	dim emailCount 
	Server.ScriptTimeout=1200	
	conn.ConnectionTimeout=1200
	strMessage = ""

	' if greater than 0 then the user has requested to send a reminder to the user
	if Request.form("email").count > 0 AND len(Request.form("email")) > 0 then 
		on error resume next 
		' find the record with the email address in it	
		strSql = "SELECT strName, strEmail, strPassword FROM [users] WHERE strEmail ='" & Request.Form("email") & "'"	
	
		' open the recordset
		rstEmail.Open strSql, conn
		
		' build the email	
		htmltext ="<html><head><title>Password</title></head><body><center><img src=""http://www.e-edi.ca/images/e-edi.gif"" alt=""e-EDI"" name=""e-edi.gif""><br><a href=""http://www.e-edi.ca"">www.e-edi.ca</a><br /><br /><font color=""black"">Your username at e-EDI is: <b>" & rstEmail("strEmail") & "</b><br /><br /><font color=""black"">Your password at e-EDI is: <b>" & rstEmail("strPassword") & "</b></font></center></body></html>"
	
		' send the email
		' 20050421
		'set objmail = server.CreateObject("CDONTS.NewMail")
		 Set objMail = Server.CreateObject("CDO.Message")
		 with objMail
			.From = "webmaster@e-edi.ca"
			.To = Request.form("email")
			.Subject = "e-EDI Password"
			
			.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
			.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSMTPServer
			.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
			.Configuration.Fields.Update
			
			.HTMLBody = htmlText
			.Send			
			
			if err.number<>0 then 
				strError = strError & "<font class=""boldTextRed"">Error sending to " & .To & ": " & err.description & "</font><br />"
				err.Clear()
			else
				' set the message		
				strMessage = "<font class=""boldTextBlack"">Your reminder has been sent to " & .To & ".</font><br /><br />"
			end if  
		end with 
		set objmail = nothing		
		' close the recordset
		rstEmail.Close
		
		' set the message		
		'strMessage = "<font class=""boldTextBlack"">Your reminder has been sent to " & Request.Form("email") & ".</font><br /><br />"
		on error goto 0
	elseif Request.Form("site").Count >0 AND len(Request.Form("site")) >0 then 
		on error resume next 
		' find the record with the email address in it	
		strSql = "SELECT strName, strEmail, strPassword FROM [teachersatsite] WHERE intSiteID =" & Request.Form("site") 	
	
		' open the recordset
		rstEmail.Open strSql, conn
				
		' set the error string to null
		strError = ""
		Set objMail = Server.CreateObject("CDO.Message")
    	with objMail
		    .Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
			.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSMTPServer
			.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
			.Configuration.Fields.Update
		end with 		
		emailCount = 0
		strMessage =  "<font class=""boldTextBlack"">Your reminder has been sent to the following people at site " & Request.Form("site") & " starting " & now() & ":</font><br />"
		'response.write strMessage
		do while not rstEmail.EOF 
			' build the email	
			htmltext ="<html><head><title>Password</title></head><body><center><img src=""http://www.e-edi.ca/images/e-edi.gif"" alt=""e-EDI"" name=""e-edi.gif""><br><a href=""http://www.e-edi.ca"">www.e-edi.ca</a><br /><br /><font color=""black"">Your username at e-EDI is: <b>" & rstEmail("strEmail") & "</b><br /><br /><font color=""black"">Your password at e-EDI is: <b>" & rstEmail("strPassword") & "</b></font></center></body></html>"
			'Response.Write htmltext	
			' send the email
			' 20050421
			'set objmail = server.CreateObject("CDONTS.NewMail")
			 'Set objMail = Server.CreateObject("CDO.Message")
			 
			 ' recreate the object every 400 emails
			 if (emailcount mod 400) = 0 then 
				set objmail = nothing 
				Set objMail = Server.CreateObject("CDO.Message")
				with objMail
					.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
					.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSMTPServer
					.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
					.Configuration.Fields.Update
				end with
			 end if 
			 
			 with objMail
				.From = "webmaster@e-edi.ca"
				.To = rstEmail("strEmail")
				.Subject = "e-EDI Password"
				
				'.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
				'.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSMTPServer
				'.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
				'.Configuration.Fields.Update
				
				.HTMLBody = htmlText
				'if emailCount >= 801 AND emailCount <= 1000 then 				
				.Send 							
				if err.number<>0 then 
					strError = strError & "<font class=""boldTextRed"">Error sending to " & .To & ": " & err.description & "</font><br />"					
					err.Clear()
				else
					strEmails =  strEmails & "<font class=""boldTextBlack"">" & .To & "</font><br />"					
				end if 
				'end if 
			end with 
			emailCount = emailCount + 1 
			' set the message					
		
		'	set objmail = nothing		
			
			rstEmail.MoveNext
		loop		
		' close the recordset
		rstEmail.Close
		'strMessage = strMessage & "<br />"
		set objmail = nothing		
		
		on error goto 0
	end if
	
	' type (0) Admin (1) Teacher
	strSql = "SELECT DISTINCT strName, strEmail, strPassword, strType FROM [users] WHERE strEmail IS NOT  NULL ORDER BY strType, strName"
	' open the recordset
	rstEmail.Open strSql, conn	
	
	' select the site just inserted
	rstData.Open "SELECT DISTINCT intSiteID FROM [TeachersAtSite] ORDER BY intSiteID", conn
	
	if not rstData.EOF then 
		' store info in array
		aData = rstData.GetRows 
								
		' get the total number of sites
		intSites = ubound(aData,2) + 1							
	else
		intSites = 0 
	end if		 
		
	' close the recordset
	rstData.Close 
%>
<html>
<head>
<!--<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />-->
	<!-- added UTF8 Encoding to get rid of funny characters -->
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" /> 

	<!-- Start CSS files-->
		<link rel="stylesheet" type="text/css" href="Styles/edi.css">
	<!-- End CSS files -->
	
	<script language="JavaScript">
	<!--
		function goForm(strEmail)
		{
			// check to be sure the email has a value
			if (strEmail.length  == 0)
				alert('This user does not have an email associated with them. Please add their email address and try again.');
			else
			{
				document.forms.Passwords.email.value = strEmail;
				document.forms.Passwords.submit();
			}
		}
		function goSiteForm(site)
		{
				document.forms.Passwords.site.value = site;
				document.forms.Passwords.submit();
		}
	//-->
	</script>
</head>
<body>
	<!-- #include virtual="/shared/page_header.inc" -->
	<form name="Passwords" method="POST" action="edi_admin_passwords.asp"> 
		<a class="reglinkMaroon" href="edi_admin.asp">Home</a>&nbsp;<font class="regtextblack">></font>&nbsp;<font class="boldtextblack">Password Management</font>
		<table border="1" bordercolor="006600" cellpadding="0" cellspacing="0" width="760">
		<tr>
			<td>
				<table border="0" cellpadding="0" cellspacing="0" width="760">
					<tr>
						<td align="right" width="490">
							<font class="headerBlue">Password Management</font>
						</td>
						<td align="right">
							<input type="button" value="Exit" name="Exit" title="EXIT Screen" onClick="javascript:window.location='edi_admin.asp';">
							&nbsp;
						</td>
					<tr><td colspan="2"><br/></td></tr>
				</table>
				<table border="0" cellpadding="0" cellspacing="0" width="760">
					<tr>
						<td align="center">
							<input type="hidden" name ="email" value="">
							<input type="hidden" name ="site" value="">
							<%							
							'response.write Server.ScriptTimeout
							if len(strMessage)>0 then
								Response.Write strMessage 
								response.write strEmails
								response.write "Completed at " & now()
								response.write "<br />"
							end if 
							
							if len(strError) > 0 then 
								Response.Write strError
							end if
							
							' put all addresses in an array
							aMembers = rstEmail.GetRows 
							
							Response.Write "<font class=""boldTextBlack"">Whom do you wish to remind about their username and password?</font><br /><br />"
							%>
						</td>
					</tr>
				</table>	
				<table border="0" cellpadding="0" cellspacing="0" width="600">
					<tr>
						<td width="300" align="right">
							<%	
							' build the option box
							Response.Write "<font class=""boldTextBlack"">Administrators: </font></td>"
							Response.Write "<td width=""175"" align=""left""><select name=""Admin_email"">"
							for row = 0 to uBound(aMembers,2)
								if aMembers(3,row) = 0 then 
									Response.Write "<option value=""" & aMembers(1,row) & """>" & trim(aMembers(0,row)) & "</option>"
								else 
									exit for 
								end if 
							next
							Response.Write "</select></td>"
							response.write "<td width=""125"" align=""left""><input type=""button"" name=""Send"" value=""Send"" onClick=""javascript:goForm(document.forms.Passwords.Admin_email.value);"">"
							'.options(document.forms.Passwords.Admin_email.selectedIndex)
							if row <= ubound(aMembers,2) then 
								Response.Write "</td></tr>"
								Response.Write "<tr><td><br/></td></tr>"
								Response.Write "<tr><td width=""300"" align=""right""><font class=""boldTextBlack"">Teachers: </font></td>"
								Response.Write "<td width=""175"" align=""left""><select name=""teacher_email"">"
								for row = row to uBound(aMembers,2)
									Response.Write "<option value=""" & aMembers(1,row) & """>" & trim(aMembers(0,row)) & "</option>"
								next
								Response.Write "</select></td>"
								response.write "<td  width=""125"" align=""left""><input type=""button"" name=""Send"" value=""Send"" onClick=""javascript:goForm(document.forms.Passwords.teacher_email.value);"">"
								'options(document.forms.Passwords.teacher_email.selectedIndex)
							end if 							
							
							if intSites > 0 then 
								Response.Write "</td></tr>"
								Response.Write "<tr><td><br/></td></tr>"
								Response.Write "<tr><td width=""300"" align=""right""><font class=""boldTextBlack"">Sites: </font></td>"	
								Response.Write "<td width=""175"" align=""left""><select name=""sites"">"
							
								for intRow = 0 to ubound(aData,2)
									Response.Write "<option value = """ & right("000" & aData(0,intRow),3) & """"
									Response.Write ">" & right("000" & aData(0,intRow),3) & "</option>"
								next
								Response.Write "</select></td>"
								response.write "<td  width=""125"" align=""left""><input type=""button"" name=""Send"" value=""Send"" onClick=""javascript:goSiteForm(document.forms.Passwords.sites.value);"">"
								'.options(document.forms.Passwords.sites.selectedIndex)
							end if 
							%>
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
	call close_adodb(rstData)
	call close_adodb(rstEmail)
	call close_adodb(conn)
	'call close_adodb(conn)
end if
%>
