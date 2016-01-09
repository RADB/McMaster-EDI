<!-- #include virtual="/shared/dbase.asp" -->
<!-- #include virtual="/shared/browser_check.asp" -->
<%
on error resume next
'-----------------------
' check values will only check if they have already been here
'-----------------------
' 0 - first time to page
' 1 - incorrect password
' 2 - username not found
' 3 - correct
'-----------------------
' will not be available if they just got to the page
if ( Request.form("check").count > 0 ) then 
	strEmail = replace(Request.form("email"),"'","''")
	strPass = Request.form("password")
	strLanguage = Request.Cookies("e-EDI")("Language")
	if strLanguage = "" then 
		strLanguage = "English"
	end if
				
	call open_adodb(conn,"MACEDI")
    
	set rstUser = server.CreateObject ("adodb.recordset")
	strQuery = "SELECT strType, strEmail, strPassword, intAccess,intProvince FROM [users] WHERE strEmail='" & strEmail & "'" 
	rstUser.Open strQuery, conn

	if rstUser.bof and rstUser.eof then 
		' if username is not found 
		call close_adodb(rstUser)
		call close_adodb(conn)
		intCheck = 2		
	else				
		strEmail = replace(rstUser("strEmail"),"'","''")
		do while not rstUser.eof
			if  strPass = rstUser("strPassword") then
				intCheck = 3
				' correct email and password
				session("user") = true
				session("id") = strEmail
				'session("browser") = intVersion
				session("language") = Request.Form("Language")
				session("province") =  rstUser("intProvince")
							
				' if admin usertype logs in then set admin to true
				if rstUser("strType") = 0  then
					session("admin") = true
					session("user") = false
					session("access") = rstUser("intAccess")
				else
					session("admin") = false
				end if
						
				if Request.Form("saveCookie") = "on" then 
					Response.Cookies("e-EDI")("Language") = Request.Form("Language")
					Response.Cookies("e-EDI")("Email") = Request.Form("Email")
					Response.Cookies("e-EDI")("SaveSettings") = Request.Form("saveCookie")
					Response.Cookies("e-EDI").path = "/"
					Response.Cookies("e-EDI").expires = dateadd("m",1,now())
				else
					' expires the cookie immediately
					Response.Cookies("e-EDI").expires = dateadd("d",-1,now())
				end if
														
				' goto the main edi page
				if session("admin") then 
					Response.Redirect ("edi_admin.asp")                    
                    'server.transfer ("edi_admin.asp")
				else
					Response.Redirect ("edi_teacher.asp")
                    
				end if
				
				exit do
			else
				' incorrect password
				intCheck = 1
			end if 
			
			rstUser.movenext
		loop
		
		' close and kill the recordset and connection object
		call close_adodb(rstUser)
	    call close_adodb(conn)
    end if	
else
	' log the user out 
	if Request.QueryString("status") = "logout" then 
		' kill the session
		session.Abandon 
	end if 
	
	intCheck = 0
	' gets language from cookie
	strLanguage = Request.Cookies("e-EDI")("Language")
	' sets default to english
	if strLanguage = "" then 
		strLanguage = "English"
	end if
	' gets email from cookie
	strEmail = Request.Cookies("e-EDI")("Email")
	strSave = Request.Cookies("e-EDI")("SaveSettings") 
	if strSave = "" then 
		strSave = "on"
	end if 
end if 
%>	
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"> 	
<html>
	<!-- #include virtual="/shared/head.asp" -->	
	<body
	<% 
		if intVersion = 0 then 
			if strEmail <> "" then 
				Response.Write " onload=""javascript:checkFocus(2);"""
			else
				Response.Write " onload=""javascript:checkFocus(1);"""
			end if 
		end if  
	%>
	>
	<form name="login" method="post" action="default.asp" onsubmit="javascript:return checkForm();">
		<input type="hidden" name="check" value="0">
		<table width="760" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td align="center">
			<!--<a href="http://www.fhs.mcmaster.ca/cscr">-->
				<a href="http://www.offordcentre.com/readiness/index.html"><img src="images/main_banner.jpg" border="0" alt="Offord Centre for Child Studies" name="main_banner.jpg"></a>
			</td>
			<!--<td>
				<font class="headerBlue">The Canadian Centre for Studies of Children at Risk </font>
				<font class="headerBlue">Offord Centre for Child Studies</font>
				<br />
				<font class="headerBlue">McMaster University </font>
				<br />
				<font class="headerBlue">Hamilton Health Sciences </font>
				<br />
				<font class="headerBlue">Hamilton, ON, Canada </font>
				<br />
				<a class="headerLinkBlue" href="http://www.fhs.mcmaster.ca/cscr">www.fhs.mcmaster.ca/cscr</a>   
				<br />
				<font class="headerBlue">Tel: (905) 521-2100, ext. 74377 </font>
			</td>-->
		</tr>
		</table>

		<br />

		<table width="760" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td align="center">
				<font class="headerMaroon">EARLY DEVELOPMENT INSTRUMENT</font>
				<br />
				<font class="subHeaderMaroon">A Population-Based Measure for Communities</font>
				<br />
				<br />
				<font class="headerMaroon">INSTRUMENT DE MESURE DU D&Eacute;VELOPPEMENT DE LA PETITE ENFANCE</font>
				<br />
				<font class="subHeaderMaroon">Une mesure ax&eacute;e sur la population &agrave; l'intention des collectivit&eacute;s</font>
			</td>
		</tr>
		</table>
		<br />
		<!-- login page -->
		<table width="760" border="1" cellpadding="0" cellspacing="0">
		<tr>
			<td align="center">			
				<table width="750" border="0" cellpadding="0" cellspacing="0">
				<tr>	
					<td rowspan="5" valign="middle">
						<img src="images/hhsc.jpg" alt="Hamilton Health Sciences" title="Hamilton Health Sciences" name="hhsc	">
					</td>
					<td align="center" colspan="3" valign="Middle">
						<br />
						<font class="headerBlack">Account Sign On</font>
						<br />						
						<br />
					</td>
					<td rowspan="5" valign="middle">
						<img src="images/fhslogo.jpg" width="150" alt="McMaster University Faculty of Health Sciences" title="McMaster University Faculty of Health Sciences" name="fhslogo">
					</td>
				</tr>
				<%
				'-----------------------
				' check values 
				'-----------------------
				' 0 - first time to page - will not enter here
				' 1 - incorrect password
				' 2 - username not found
				' 3 - correct - already redirected
				'-----------------------
				if ( Request.form("check").count > 0 ) then 
					Response.Write "<script language=""javascript"">"
					Response.Write "document.login.check.value =" & intCheck & ";"
					Response.Write "</script>"
					
					Response.Write "<tr><td colspan=""3"" align=""center"">"
					select case intCheck
						case 1
							Response.Write "<font class=""boldtextred"">Incorrect password... </font>&nbsp;<a class=""bigLinkRed"" onMouseOver=""javascript:window.status='Email my forgotten password to " & strEmail & ".'; return true;"" onMouseOut=""javascript:window.status=''; return true;"" href=""javascript:window.location='edi_mailPwd.asp?email=" & strEmail & "';"">Click Here</a><font class=""boldtextred"">&nbsp;if you forget your password</font><br /><br />"
						case 2
							Response.Write "<font class=""boldtextred"">Email not found... Please double check your email address.</font><br /><br />"						
					end select
					Response.Write "</td></tr>"
				end if 
				%>
				<tr valign="top">
					<td width="100" align="right" nowrap> 
						<font class="boldtextblack">Email :&nbsp;&nbsp;</font>
					</td>
					<td width="175" align="left">
						<input type="text" title="Email address" name="email" value="<%=strEmail%>" size="25">						
					</td>
					<td width="275">
						<!-- default language set to English -->
						<%
						if strLanguage = "English" then 
							Response.Write "<INPUT type=""radio"" id=""language"" name=""language"" value=""English"" checked>"
							Response.Write "<font class=""boldtextblack"">English&nbsp;&nbsp;</font>"
							Response.Write "<INPUT type=""radio"" id=""language"" name=""language"" value=""French"">"
							Response.Write "<font class=""boldtextblack"">French&nbsp;&nbsp;</font>"
						elseif strLanguage = "French" then 
							Response.Write "<INPUT type=""radio"" id=""language"" name=""language"" value=""English"">"
							Response.Write "<font class=""boldtextblack"">English&nbsp;&nbsp;</font>"
							Response.Write "<INPUT type=""radio"" id=""language"" name=""language"" value=""French"" checked>"
							Response.Write "<font class=""boldtextblack"">French&nbsp;&nbsp;</font>"
						end if 
						%>
					</td>
				</tr>
				<tr>
				    <td width="100" align="right">
						<font class="boldtextblack">Password :&nbsp;&nbsp;</font>
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
						<font class="regtextblack">Save my settings (Password will not be saved)</font>
					</td>
				</tr>
				</table>
				<br />
			</td>
		</tr>
		</table>
		<br />
        <table  width="760" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td align="center">
              <font class="regTextMaroon">&copy; McMaster University, Hamilton, Ontario, Canada.</font>
              <br />
              <font class="regTextMaroon">The Early Development Instrument (EDI), authored by Dr. Magdalena Janus et al, is the copyright of McMaster University (Copyright &copy; 2000, McMaster University).</font>
              <br /><br />
              <font class="regTextMaroon">The EDI has been provided under license from McMaster University and must not be copied, distributed or used in any way without the prior written consent of McMaster University. Contact the Offord Center for Child Studies for licensing details, email: <a href="mailto:walshci@mcmaster.ca" class="reglinkMaroon">walshci@mcmaster.ca</a></font>
            </td>
          </tr>
        </table>
		<!-- #include virtual="/shared/page_footer.inc" -->
	</form>
	</body>
</html>

