
<%
'<!-- #include virtual="/shared/admin_security.asp" -->
if  session("user") then
	dim blnExists 
	if session("Language") = "English" then
		strGenerate = "Generating Report"
		strMoment = "Your file will be ready momentarily..."
		strThank = "Thank you for your patience."
		strComplete = "Report Complete"
		strReady = "Your file is now ready..."
		strClose = "Close Window"
		strPatient = "Please be patient... your file will be ready soon."
		strErrorMsg = "The report cannot be generated as no data is available that fits the criteria."
		strDownload = "Download file here"
		strTitle = "e-EDI: Teacher Report Generator"
	else
		strGenerate = "Rapport engendrant"
		strMoment = "Votre fichier sera prêt momentanément..."
		strThank = "Merci pour votre patience."
		strComplete = "Le rapport Complet"
		strReady = "Votre fichier est maintenant prêt.. "
		strClose = "Fenêtre proche"
		strPatient = "S'il vous plaît être patient... votre fichier sera prêt bientôt. "
		strErrorMsg = "Le rapport ne peut pas être engendré comme aucunes données sont disponibles qu'ajuste les critères. "
		strDownload = "Le fichier de chargement ici"
		strTitle = "e-EDI: Le Générateur de Rapport d'enseignant"
	end if 
	
' if the user has not logged in they will not be able to see the page
	if Request.form("rpt").Count <> 0 then 		
		dim ePDF
		dim strFilename
	
		set ePDF = server.CreateObject("BS_PDF_EXPORTER.BS_PDF_EXPORT")
		
		SELECT CASE Request.form("rpt")
		Case "Generate"
			' get the XML filename
			strXML = Request.Form("XML")
						
			With ePDF
			   ' turn logging on
			   .Logging = True
			   
			   ' log to file
			   '.LogPath = App.Path
						

			   ' set the connection string
			   .Connection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=d:\websites\e-edica\data\edi_data.mdb"
				
				select case lcase(strXML)
					case "class_summary.rpx"
						if Request.Form("classes") = 0 then 
							' set the PDF file name
							.Sqlstring = "SELECT * FROM class_summary WHERE strEmail = '" & Request.Form("email") & "'"
						else
							' specific class
							if len(Request.Form("classes")) = 9 then 
								.Sqlstring = "SELECT * FROM class_summary WHERE intClassID = " & Request.Form("classes")
							end if
						end if
					case "edi_summary.rpx"
						if Request.Form("student") = 0 then 
							' set the PDF file name
							.Sqlstring = "SELECT * FROM student_summary WHERE strEmail = '" & Request.Form("email") & "' ORDER BY strEDIID"
						else
							' all students at site
							if len(Request.Form("student")) = 3 then 
								.Sqlstring = "SELECT * FROM student_summary WHERE intSiteID = " & Request.Form("student")	 & " ORDER BY strEDIID"						
							' all students at school
							elseif len(Request.Form("student")) = 6 then 
								.Sqlstring = "SELECT * FROM student_summary WHERE intSchoolID = " & Request.Form("student") & " ORDER BY strEDIID"
							' all teachers students
							elseif len(Request.Form("student")) = 8 then 
								.Sqlstring = "SELECT * FROM student_summary WHERE intTeacherID = " & Request.Form("student") & " ORDER BY strEDIID"
							' all classes students
							elseif len(Request.Form("student")) = 9 then 
								.Sqlstring = "SELECT * FROM student_summary WHERE intClassID = " & Request.Form("student") & " ORDER BY strEDIID"
							' student
							elseif len(Request.Form("student")) = 11 then 
								.Sqlstring = "SELECT * FROM student_summary WHERE strEDIID = '" & Request.Form("student") & "' ORDER BY strEDIID" 
							end if
						end if
				end select
				
			   ' set the XML File path
			   .XMLFilePath = "d:\websites\e-edica\XMLfiles\" & strXML
					
				' set the PDF file path
			   strPath = "d:\websites\e-edica\www\pdfs\"
					   
			   ' set the PDF file path
			   .PDFFilePath = strPath
			   
			   ' set the PDF file name
			   .PDFFileName = replace(strXML,".rpx","_")
			   
			   ' export pdf
			   strFilename = .Export_To_Pdf
			End With
				
			' split the returned string into an array
			aProgress = split(strFilename,"|")
			
			' if success then redirect to the PDF
			if aProgress(0) = 4 then 
				strLocation = replace(aProgress(2),strPath, "pdfs/")	
			elseif aProgress(0) = 2 and aProgress(2) = "Recordset is Empty" then 
				strError = "<font class=""regtextred"">" & strErrorMsg & "</font>"
			else
				' put in error var and print below 
				' one of two things happened 
				'	1) paramater missing - write the parameter that is missing
				'	2) error has occurred - write the error
				strError = "<font class=""regtextred"">" &  aProgress(0) & " - " & aProgress(1) & " - " & replace(aProgress(2),"ActiveReports","") & "</font>"
			end if 
		END SELECT

		' kill the object
		set ePDF = nothing
			
		' set the header
		strHeader = strGenerate
		' set the subheader
		strSub = "<font class=""subHeaderBlue"">" & strMoment & "</font><br /><font class=""subHeaderBlue"">" & strThank & "</font>"
			
		'set exists = false
		blnExists = false
	
	elseif Request.QueryString("pdf").Count <>0 then 
		blnExists = IsFileExists(Request.QueryString("pdf"))
		' check to be sure the file exists
		If  blnExists = True  Then
			' set the header
			strHeader = strComplete
			' set the subheader
			strSub = "<font class=""subHeaderBlue"">" & strReady & "</font><br /><br /><a href=""" & Request.QueryString("pdf") & """ class=""bigLinkRed""><img src=""images/pdf.gif"" title=""" & strDownload & """ border=""0"">&nbsp;" & replace(Request.QueryString("pdf"),"pdfs/","")  & "</a>"
		Else
			' does not exist - run for another 5 seconds
			strLocation = Request.QueryString("pdf")
			' set the header
			strHeader = strGenerate
			' set the subheader
			strSub = "<font class=""subHeaderBlue"">" & strMoment & "</font><br /><font class=""subHeaderBlue"">" & strThank & "</font>"
		End If
	else
		' close - no rights to be here
		Response.Write "<SCRIPT LANGUAGE=""JavaScript"">"
		'	Response.write "window.open('" & Request.QueryString("zip") & "');"
		'	Response.write "window.close();"
		Response.Write "</SCRIPT>"
	end if
	%>
<html>
<head>
	<!-- Start CSS files-->
		<link rel="stylesheet" type="text/css" href="Styles/edi.css">
	<!-- End CSS files -->

	<title><%=strTitle%></title>
	
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
					<a href="javascript:window.close();" class="bigLinkBlue"><%=strClose%></a>
				<%
				end if 
				%>
				&nbsp;&nbsp;
			</td>
		</tr>
	</table>
	<table  width="490" border="1" cellpadding="0" cellspacing="0" bordercolor="006600">	
		<tr>
			<td align="center">		
				<table  width="480" border="0" cellpadding="0" cellspacing="0">		
					<tr>
						<td align="center">
							<%
							if not blnExists then 
							%>
								<img src="images/hourglass.gif" name="Hourglass" title="<%=strPatient%>">
							<%
							end if 
							%>
						</td>
						<td align="center" valign="top">
							<font class="headerBlue"><%=strHeader%></font>
							<br />
							<%=strSub%>
							<br />
							<br />
						</td>
						<td>
							<%
							if not blnExists then 
							%>
							<img src="images/hourglass.gif" name="Hourglass" title="<%=strPatient%>">
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
				<%'javascript:goWindow('edi_teacher_zip.asp','Map','490','240','top=0,left=125')
				if strLocation <> "" then 
					Response.Write "<SCRIPT LANGUAGE=""JavaScript"">"
						'Response.write "NewUrl = 'edi_teacher_zip.asp?zip=" & strLocation & "','Backup','0','0','top=0,left=0';"
						Response.write "NewUrl = 'edi_teacher_reports.asp?pdf=" & strLocation & "','Generate Report','0','0','top=0,left=0';"
						Response.write "setTimeout('top.location.href = NewUrl',3000);"
					Response.Write "</SCRIPT>"
				end if 
				%>
			</td>
		</tr>
	</table>
	<br />
	
	<table  width="480" border="0" cellpadding="0" cellspacing="0">			
		<tr>
			<td align="center">
				<!--<hr color="006600" size="1">-->				
				<font class="boldtextGreen">© Offord Centre for Child Studies</font>
			    <br />
			    <font class="boldtextGreen">McMaster University &amp; Hamilton Health Sciences, Hamilton ON, Canada</font>
			    <br />
			    <font class="boldtextGreen">Tel.(905)521-2100, ext.74352 or ext.77370</font>
		    </td>
		</tr>
	</table>
</body>
</html>
<%
else
	' recalls the main page - if expired will inform user
	Response.Write "<SCRIPT LANGUAGE=""JavaScript"">"
			Response.write "window.opener.document.location = 'edi_teacher.asp';"
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
	    
	filename = "d:\websites\e-edica\www\" & filename
	
	If objFSO.FileExists( FileName ) = True  Then
		IsFileExists = True
	Else
		IsFileExists = False
	End If
	  
	Set objFSO = Nothing   
End Function
%>