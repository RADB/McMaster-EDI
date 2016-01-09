<!-- #include virtual="/shared/dbase.asp" -->
<!-- #include virtual="/shared/browser_check.asp" -->

<%
if NOT session("admin") then
%>
<html>
	<head>
		<title>www.e-EDI.ca Security</title>
		<script language="javascript" type="text/javascript" src="js/login.js"></script>
		<link rel="stylesheet" type="text/css" href="Styles/edi.css">
	</head><BODY >
	<%
	Response.Write "<body "
	if intVersion = 0 then 
		Response.Write "onload=""javascript:checkFocus(login.check.value);"""
	end if
	Response.Write ">"

	strTopBar = "<table width=""760"" border=""1"" cellpadding=""1"" cellspacing=""0"" bordercolor=""006600"" bgcolor=""Gainsboro"">"
	strTopBar = strTopBar & "<tr><td><table width=""750"" border=""0"" cellpadding=""3"" cellspacing=""0""><tr><td align=""left""><font class=""boldTextBlack"">&nbsp;&nbsp;"
	strTopBar = strTopBar & "Session expired/Not logged on"
	strTopBar = strTopBar & "</td><td align=""right""><font class=""boldTextBlack"">" & formatdatetime(now,vblongdate) & " - " & time() & "&nbsp;&nbsp;</font></td></tr></table>"	
	strTopBar = strTopBar &	"</td></tr></table>"
	'strTopBar = strTopBar &	"<br />"
	%>
	
	<!-- #include virtual="/shared/page_header.inc" -->
	<form name="login" method="post" action="default.asp">
	<input type="hidden" name="check" value="0">
	<br />
	<table width="760" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td align="center">				
				<font class="boldTextBlack">Your session has either expired or you do not have permission to be here!</font>
			</td>
		</tr>
	</table>
	<br />
	<br />
	<table width="760" border="1" cellpadding="0" cellspacing="0">
		<tr>
			<td align="middle">			
				<table width="750" border="0" cellpadding="0" cellspacing="0">
				<tr>	
					<td rowspan="5" valign="middle">
						<img src="images/hhsc.jpg" alt="Hamilton Health Sciences" title="Hamilton Health Sciences" name="hhsc	">	
					</td>
					<td align="middle" colspan="3" valign="Middle">
						<br />
						<font class="headerBlack">Account Sign On</font>
						<br />						
						<br />
					</td>
					<td rowspan="5" valign="middle">
						<img src="images/fhslogo.jpg" width="150" alt="McMaster University Faculty of Health Sciences" title="McMaster University Faculty of Health Sciences" name="fhslogo">
					</td>
				</tr>
				<tr valign="top">
					<td width="100" align="right" nowrap> 
						<font class="boldtextblack">Email :&nbsp;&nbsp;</font>
					</td>
					<td width="175" align="left">
						<input type="text" name="email" value="<%=strEmail%>" size="25">						
					</td>
					<td width="275">
						<!-- default language set to English -->
						<%
						strLanguage = Request.Cookies("e-EDI")("Language")
						' sets default to english
						if srLanguage = "" then 
							strLanguage = "English"
						end if
						if strLanguage = "English" then 
							Response.Write "<INPUT type=""radio"" id=""language"" name=""language"" value=""English"" checked>"
							Response.Write "<font class=""boldtextblack"">English&nbsp;&nbsp;</font>"
							Response.Write "<INPUT type=""radio"" id=""language"" name=""language"" value=""French"">"
							Response.Write "<font class=""boldtextblack"">French&nbsp;&nbsp;</font>"
							strConfirmLanguage = "Are you sure you want to lock this student?\n\nOnce locked you will no longer be able to edit their EDI."
							strConfirmLanguage2 = "Your questionnaire is now being sent to McMaster."
						elseif strLanguage = "French" then 
							Response.Write "<INPUT type=""radio"" id=""language"" name=""language"" value=""English"">"
							Response.Write "<font class=""boldtextblack"">English&nbsp;&nbsp;</font>"
							Response.Write "<INPUT type=""radio"" id=""language"" name=""language"" value=""French"" checked>"
							Response.Write "<font class=""boldtextblack"">French&nbsp;&nbsp;</font>"
							strConfirmLanguage = "&Ecirc;tes-vous s&ucirc;r de vouloir v&eacute;rouiller le questionnaire de cet &eacute;l&egrave;ve?\n\nUne fois v&eacute;rouill&eacute;, vous ne pourrez plus modifier le questionnaire."
							strConfirmLanguage2 = "Your questionnaire is now being sent to McMaster."
						end if 
						%>
					</td>
				</tr>
				<tr>
				    <td width="100" align="right">
						<font class="boldtextblack" nowrap>Password :&nbsp;&nbsp;</font>
					</td>
					<td width="175" align="left">
						<input type="password" name="password" value="" size="25">						
					</td>
					<td width="275">
						&nbsp;
						<input type="submit" name="Login" value="Login">
					</td>
				</tr>
				<tr>
					<td></td>
					<td colspan="2">
						<INPUT type="checkbox" id="savecookie" name="savecookie" checked>
						<font class="regtextblack" nowrap>Save my settings</font>
					</td>
				</tr>
				</table>
				<br />
			</td>
		</tr>
		</table>
		<!-- #include virtual="/shared/page_footer.inc" -->
	</form>
	</body>
</html>
<%
	blnSecurity = false
else
	blnSecurity = true
	strTopBar = "<table width=""760"" border=""1"" cellpadding=""1"" cellspacing=""0"" bordercolor=""006600"" bgcolor=""Gainsboro"">"
	strTopBar = strTopBar & "<tr><td><table width=""750"" border=""0"" cellpadding=""3"" cellspacing=""0""><tr><td align=""left""><font class=""boldTextBlack"">&nbsp;&nbsp;"
		call open_adodb(conn,"MACEDI")
		set rstUser = server.CreateObject("adodb.recordset")
			strQuery = "SELECT strName FROM users WHERE strEmail='" & session("id") & "'"
			rstUser.open strQuery,conn
			if not(rstUser.EOF) then
				strTopBar = strTopBar & trim(rstUser("strname"))
			else 
				strTopBar = strTopBar & "No user found"
			end if 
		call close_adodb(rstUser)
		call close_adodb(conn)
	strTopBar = strTopBar & "</td><td align=""right""><font class=""boldTextBlack"">" 
	
	if session("language") = "French" then 
		strTopBar = strTopBar & french_day(datepart("w",date(),vbSunday)) & ", le " & day(date()) & " " & French_month(month(date())) & " " & year(date())
		strConfirmLanguage = "&Ecirc;tes-vous s&ucirc;r de vouloir v&eacute;rouiller le questionnaire de cet &eacute;l&egrave;ve?\n\nUne fois v&eacute;rouill&eacute;, vous ne pourrez plus modifier le questionnaire."
		strConfirmLanguage2 = "Your questionnaire is now being sent to McMaster."
	else
		strTopBar = strTopBar & formatdatetime(now,vblongdate) 
		strConfirmLanguage = "Are you sure you want to lock this student?\n\nOnce locked you will no longer be able to edit their EDI."
		strConfirmLanguage2 = "Your questionnaire is now being sent to McMaster."
	end if 
	strTopBar = strTopBar & " - " & time() & "&nbsp;&nbsp;</font></td></tr></table>"	
	strTopBar = strTopBar &	"</td></tr></table>"
	'strTopBar = strTopBar &	"<br />"
end if 	

function French_Month(intMonth)
	select case intMonth
		case 1 
			French_Month = "janvier"
		case 2
			French_Month = "février"
		case 3
			French_Month = "mars"
		case 4
			French_Month = "avril"
		case 5 
			French_Month = "mai"
		case 6 
			French_Month = "juin"
		case 7 
			French_Month = "juillet"
		case 8 
			French_Month = "août"
		case 9 
			French_Month = "septembre"
		case 10 
			French_Month = "octobre"
		case 11
			French_Month = "novembre"
		case 12 
			French_Month = "décembre"
	end Select
end function

function French_Day(intDay)
	select case intDay
		case 1 
			French_Day = "dimanche"
		case 2
			French_Day = "lundi"
		case 3 
			French_Day = "mardi"
		case 4 
			French_Day = "mercredi"
		case 5
			French_Day = "jeudi"
		case 6 
			French_Day = "vendredi"
		case 7 
			French_Day = "samedi"
	end select 
end Function
%>