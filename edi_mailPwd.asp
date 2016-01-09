<!-- #include virtual="/shared/dbase.asp" -->
<% 
stremail= Request.QueryString("email")
call open_adodb(conn,"MACEDI")
%>

<html>
	<head>
		<title>e-EDI.ca Password Help Page</title>
		<link rel="stylesheet" type="text/css" href="../Styles/edi.css">
	</head>
			
	<body>
		<!-- #include virtual="/shared/page_header.inc" -->
		<table  width="760" border="0" cellpadding="0" cellspacing="0">			
			<tr><td colspan="3"><br /></td></tr>
			
			<%		
			set	rstPassword = server.createobject("ADODB.recordset") 
			strQuery = "SELECT strPassword FROM users WHERE strEmail='"& strEmail &"'"
			rstPassword.Open strQuery, conn

			if not (rstPassword.eof and rstPassword.bof) then 
			on error resume next 
				htmltext ="<html><head><title>Password</title></head><body><center><img src=""http://www.e-edi.ca/images/e-edi.gif"" alt=""e-EDI"" name=""e-edi.gif""><br><a href=""http://www.e-edi.ca"">www.e-edi.ca</a><br /><br /><font color=""black"">Your password at e-EDI is: <b>" & rstPassword("strPassword") & "</b></font></center></body></html>"

				'set objmail = server.CreateObject("CDONTS.NewMail")
				'	objmail.From = "webmaster@e-edi.ca"
				'	objmail.To = strEmail
				'	objmail.Subject = "e-EDI Password"
				'	objmail.BodyFormat = 0
				'	objmail.MailFormat = 0
				'	objmail.Body = htmlText
				'	objmail.Send 
				'set objmail = nothing
				strerror = ""
				Set objMail = Server.CreateObject("CDO.Message")
				 with objMail
					.From = "webmaster@e-edi.ca"
					.To = strEmail
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
				
				
				if strError = "" then 
					response.write "<tr><td colspan=""3"" align=""center""><font class=""regText"">Your password has been sent to&nbsp;</font><font class=""boldTextBlack"">" & strEmail & "!</font>&nbsp;&nbsp;<input type=""button"" value=""Login"" onclick=""javascript:window.location='default.asp?email=" & strEmail & "';"" name=""Login""></td></tr>"			
				else
					response.write "<tr><td colspan=""3"" align=""center"">" & strError & "</td></tr>"
				end if 
				
				on error goto 0
			else 
				Response.write "<tr><td colspan=""3"" align=""center""><font class=""boldTextBlack"">The specified email """ & strEmail & """ does not exist in the system. <br><Br>Please contact the <a href=""mailto:webmaster@e-edi.ca"" class=""reglink"">Webmaster</a>.</font><br><br></td></tr>"
			end if 
			
			call close_adodb(rstPassword)
			call close_adodb(conn)	
			%>	
		</table>
		<!-- #include virtual="/shared/page_footer.inc" -->
	</body>
</html>