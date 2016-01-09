<!-- #include virtual="/shared/admin_security.asp" -->
<%

if  session("admin") then
	dim blnExists 
	dim blnLicensed 
' if the user has not logged in they will not be able to see the page
	if Request.QueryString("zip").Count = 0 then 
		dim eZip
		dim strFilename, strFiletozip
		call open_adodb(conn, "MACEDI")
        'set rstBackup = server.CreateObject("adodb.recordset")
	
        dim filename 
        dim datetime 
        datetime = now()
        filename = year(datetime) & "_"& right("0" & month(datetime),2) & "_"& right("0" & day(datetime),2)& "_"&right("0" & hour(datetime),2)&right("0" & minute(datetime),2)&right("0" & second(datetime),2)
        strFileName = "c:\websites\e-edica\www\backups\EDI_Backup_" & filename & ".bak"
        strLocation =   replace(replace(strFileName,"c:\websites\e-edica\www\",""),"\","/")      
	    strSql = "Exec dbo.BackupDataBase '" & filename & "'"

        conn.execute strsql
        call close_adodb(conn)
        
'		set eZip = server.CreateObject("XceedSoftware.XceedZip")
'		
'		With ezip
'			blnLicensed = .License("SFX50-A416Z-K8DPW-NKCA")
'			'.zip.Compression = xclHigh
'			strFiletozip = "d:\websites\e-edica\data\edi_data.mdb"  
'			strFilename = "d:\websites\e-edica\www\zips\backup_" & year(date) & monthname(month(date),true)& right("0" & Day(date),2) & "_" & right("0" &hour(now),2) & right("0" &minute(now),2) & right("0" & second(now),2) & ".zip"
'			' set the file location
'			strLocation = replace(replace(strFileName,"d:\websites\e-edica\www\",""),"\","/")		'	Response.Write strFiletozip		'	Response.Write strFilename		'	Response.Write strLocation
'			.UseTempFile = False			' No need for a temp file for this application
'			.PreservePaths = false			' Do not store the paths where the source files are from
'			.ProcessSubfolders = false
			' set the file to zip
'			.FilesToProcess = strfiletozip
			
			' set the filename
'			.ZipFilename = strFilename
   
'			' zip the file 
'			call .zip
'		End With
		
		'
		
		' kill the object
		set eZip = nothing		
		
		Response.Write "<br>"
						
		' split the returned string into an array
		'aProgress = split(strFilename,"|")
						
		' if success then redirect to the PDF
		'if aProgress(0) = 4 then 
		'	strLocation = "zips/" & aProgress(2)		
		'else 
			' zip file creation failed
		'	strError = "<font class=""regtextred"">" &  aProgress(1) & " - " & aProgress(2) & "</font>"	 
		'end if 
		
		' set the header
		strHeader = "Data Backup"
		' set the subheader
		strSub = "<font class=""subHeaderBlue"">Your file will be ready momentarily...</font><br /><font class=""subHeaderBlue"">Thank you for your patience.</font>"
		
		'set exists = false
		blnExists = false
	else
		'https://www.e-edi.ca/edi_admin_zip.asp?zip=backups/EDI_Backup_2014_11_04_164332.bak
		'response.write Request.QueryString("zip")
		blnExists = IsFileExists(Request.QueryString("zip"))
		' check to be sure the file exists
		If  blnExists = True  Then
			' set the header
			strHeader = "Data Backup Complete"
			' set the subheader
			strSub = "<font class=""subHeaderBlue"">Your file is now ready...</font><br /><br /><a href=""" & Request.QueryString("zip") & """ class=""bigLinkRed""><img src=""images/winzip.gif"" title=""Download file here"" border=""0"">&nbsp;" & replace(Request.QueryString("zip"),"zips/","")  & "</a>"
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

	<title>e-EDI: Admin Backup</title>
	
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
			'Response.write "NewUrl = 'edi_admin_zip.asp?zip=" & strLocation & "','Backup','0','0','top=0,left=0';"
			Response.write "NewUrl = 'edi_admin_zip.asp?zip=" & strLocation & "','Backup','0','0','top=0,left=0';"
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
	    
	filename = "c:\websites\e-edica\www\" & filename
   ' response.write filename
	
	If objFSO.FileExists( FileName ) = True  Then
		IsFileExists = True
	Else
		IsFileExists = False
	End If
	  
	Set objFSO = Nothing   
End Function
%>