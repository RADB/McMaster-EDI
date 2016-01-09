<!-- #include virtual="/shared/admin_security.asp" -->
<%

if  session("admin") then
	dim blnExists 
	dim blnLicensed 
' if the user has not logged in they will not be able to see the page
	if Request.QueryString("zip").Count = 0 then 
		
		call open_adodb(conn, "MACEDI")
        
        strFileName = "f:\websites\e-edica\www\backups\ExportToAccessComplete.txt"
        strLocation =   replace(replace(strFileName,"f:\websites\e-edica\www\",""),"\","/")      
	    strSql = "Exec dbo.ExportDataToAccess"

        conn.execute strsql
        if conn.errors.count > 0 then 		
			Response.Write conn.errors(0).number & conn.errors(0).description
		end if 
        call close_adodb(conn)
        
'		
		Response.Write "<br>"
						
		
		' set the header
		strHeader = "Data Backup to Access"
		' set the subheader
		strSub = "<font class=""subHeaderBlue"">Your file will be ready momentarily...</font><br /><font class=""subHeaderBlue"">Thank you for your patience.</font>"
		
		'set exists = false
		blnExists = false
	else
		
		blnExists = IsFileExists(Request.QueryString("zip"))
		' check to be sure the file exists
		If  blnExists = True  Then
			' set the header
			strHeader = "Data Backup Complete"
			' set the subheader
			strSub = "<font class=""subHeaderBlue"">Your file is now ready...</font><br /><br /><a href=""" & replace(Request.QueryString("zip"),"ExportToAccessComplete.txt","EDI_2011/mdb") & """ class=""bigLinkRed""><img src=""images/winzip.gif"" title=""Download file here"" border=""0"">&nbsp;" & replace(Request.QueryString("zip"),"ExportToAccessComplete.txt","EDI_2011/mdb") & "</a>"
		Else
			' does not exist - run for another 5 seconds
			strLocation = Request.QueryString("zip")
			' set the header
			strHeader = "Data Backup"
			' set the subheader
			strSub = "<font class=""subHeaderBlue"">Your file will be ready momentarily...</font><br /><font class=""subHeaderBlue"">Thank you for your patience.</font>"
		End If
		'Response.Write "<SCRIPT LANGUAGE=""JavaScript"">"
		'	Response.write "window.open('" & Request.QueryString("zip") & "');"
		'	Response.write "window.close();"
		'Response.Write "</SCRIPT>"
	end if
	%>
<html>
<head>
	<!-- Start CSS files-->
		<link rel="stylesheet" type="text/css" href="Styles/edi.css">
	<!-- End CSS files -->

	<title>e-EDI: Admin Backup to Access</title>
	
	<!-- Start Meta Tags -->
		<meta name="author" content="Andrew Renner">
		<meta name="description" content="electronic EDI, an early development instrument">
		<meta name="keywords" content="Andrew,Renner,McMaster,University,Hamilton,Ontario,Canada,children,early, development, instrument,experience,education,edi,e-edi">
	<!-- End Meta Tags -->
</head>
<body>
	<table  width="480" border="0" cellpadding="0" cellspacing="0">		
		<!--<tr><td align="right" colspan="3"><a href="javascript:window.close();" class="bigLinkBlue">Close Window</a>&nbsp;&nbsp;</td></tr>-->
		<tr>
			<td width="160"></td>
			<td valign="middle" align="center" width="160">
				<img src="images/e-edi.gif" alt="e-EDI.ca" name="e-edi.gif">
				<br/><br/>
			</td>
			<td align="right" valign="top">
				<%
				if blnExists then 
				%>
					<a href="javascript:window.close();" class="bigLinkBlue">Close Window</a>
				<%
				end if 
				%>
				&nbsp;&nbsp;
			</td>
		</tr>
	</table>
	<table  width="480" border="0" cellpadding="0" cellspacing="0">		
		<tr>
			<td align="center">
				<%
				if not blnExists then 
				%>
					<img src="images/hourglass.gif" name="Hourglass" title="Please be patient... your file will be ready soon.">
				<%
				end if 
				%>
			</td>
			<td align="center" valign="top">
				<font class="headerBlue"><%=strHeader%></font>
				<br />
				<%=strSub%>
			</td>
			<td>
				<%
				if not blnExists then 
				%>
				<img src="images/hourglass.gif" name="Hourglass" title="Please be patient... your file will be ready soon.">
				<%
				end if 
				%>
			</td>
		</tr>
	</table>
	<table border="0" cellpadding="0" cellspacing="0" width="480" align="center">
	<tr>
		<td align="left">
			
			<%
				if strError <> "" then 
					Response.Write strError & "<br /><br />"
				end if 
			%>
		</td>
	</tr>
	</table>
	<%'javascript:goWindow('edi_admin_zip.asp','Map','490','240','top=0,left=125')
	if strLocation <> "" then 
		Response.Write "<SCRIPT LANGUAGE=""JavaScript"">"			
			Response.write "NewUrl = 'edi_admin_CopyToAccess.asp?zip=" & strLocation & "','Backup','0','0','top=0,left=0';"
			Response.write "setTimeout('top.location.href = NewUrl',3000);"
		Response.Write "</SCRIPT>"
	end if 
	%>
	<br />
	<table  width="480" border="0" cellpadding="0" cellspacing="0" align="center">			
		<tr>
			<td align="center">
				<!--<hr color="006600" size="1">-->
				<font class="boldtextGreen">© The Canadian Centre for Studies of Children at Risk</font>
				<br />
				<font class="boldtextGreen">McMaster University & Hamilton Health Sciences, Hamilton ON, Canada</font>
				<br />
				<font class="boldtextGreen">Tel.(905)521-2100, ext.74377</font>
			</td>
		</tr>
	</table>
</body>
</html>
<%
else
	Response.Write "<SCRIPT LANGUAGE=""JavaScript"">"
			Response.write "window.opener.document.location = 'edi_admin.asp';"
			Response.write "window.close();"
	Response.Write "</SCRIPT>"
end if

' **********************************
' Function to check file Existance
' **********************************
Function IsFileExists(byVal FileName)
	 
	If FileName = ""  Then
		IsFileExists = False
		Exit Function
	End If
	 
	Dim objFSO
	    
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	    
	filename = "f:\websites\e-edica\www\" & filename
   ' response.write filename
	
	If objFSO.FileExists( FileName ) = True  Then
		IsFileExists = True
	Else
		IsFileExists = False
	End If
	  
	Set objFSO = Nothing   
End Function
%>