<!-- #include virtual="/shared/admin_security.asp" -->
<%
' if the user has not logged in they will not be able to see the page
if blnSecurity then
%>
<html>
<head>
	<!-- Start CSS files-->
		<link rel="stylesheet" type="text/css" href="../Styles/edi.css">
	<!-- End CSS files -->
	
	<%
	Set UploadProgress = Server.CreateObject("Persits.UploadProgress")
	PID = "PID=" & UploadProgress.CreateProgressID()
	barref = "edi_upload_framebar.asp?to=10&" & PID
	%>
	
	<script language="JavaScript">
	function ShowProgress()
	{
	  strAppVersion = navigator.appVersion;
	  // only accept xls files
	  if (document.MyForm.FILE1.value != "" && document.MyForm.FILE1.value.indexOf('.xls') > 0)
	  {
	    if (strAppVersion.indexOf('MSIE') != -1 && strAppVersion.substr(strAppVersion.indexOf('MSIE')+5,1) > 4)
	    {
	      winstyle = "dialogWidth=375px; dialogHeight:130px; center:yes";
	      window.showModelessDialog('<% = barref %>&b=IE',null,winstyle);
	    }
	    else
	    {
	      window.open('<% = barref %>&b=NN','','width=370,height=115', true);
	    }
	  }	  
	  else
	  {
		alert('The following error occurred: \n \n You must supply a valid path to a Microsoft EXCEL file (*.xls)');
		return false;
	  }
	}
	</script> 
	<script language="javascript" type="text/javascript" src="js/window.js"></script>
</head>
<body>
	<!-- #include virtual="/shared/page_header.inc" -->
	<form name="MyForm" method="POST" enctype="multipart/form-data" action="edi_admin_import_progress.asp?<% = PID %>"	OnSubmit="return ShowProgress();"> 
		<a class="reglinkMaroon" href="edi_admin.asp">Home</a>&nbsp;<font class="regtextblack">></font>&nbsp;<font class="boldtextblack">Data Import File Selection</font>
		<table border="1" bordercolor="006600" cellpadding="0" cellspacing="0" width="760">
		<tr><td>			
			<table border="0" cellpadding="0" cellspacing="0" width="760">
				<tr>
					<td align="right" width="490">
						<font class="headerBlue">Data Import File Selection</font>
					</td>
					<td align="right">
						<input type="button" value="Tutorial" name="Tutorial" title="Show Import Tutorial" onClick="javascript:goWindow('EDI_ADMIN_IMPORT_tutorial.htm', 'edi_import_tutorial', 824, 648,'toolbar=no,scrollbars=no,location=no,statusbar=no,menubar=no,resizable=no');"><input type="button" value="Exit" name="Exit" title="EXIT Screen" onClick="javascript:window.location='edi_admin.asp';">
						&nbsp;
					</td>
				<tr><td colspan="2"><br/></td></tr>
			</table>
			<table border="0" cellpadding="0" cellspacing="0" width="760">
			<tr>
				<td width="600" align="center">
					<br />
					<font class="regtextblack"> Select your properly formatted Microsoft Excel file</font>
				</td>
				<td align="right" width="160"></td>
			</tr>
			<tr>
				<td align="right" width="600">
					<input type="FILE" size="60" name="FILE1">
				</td>
				<td width="160"></td>
			</tr>
			<tr>
				<td align="center">
					<img src="images/greenbb.gif">
					<a class="reglinkMaroon" href="documents/edi%20template.xls">Download template here</a>
					<img src="images/greenbb.gif">
					<br/>
				</td>
			</tr>
			<tr>
				<td width="600" align="right">
					<font class="regtextblack">Once you have selected your Microsoft Excel file click --></font>
					<input type=SUBMIT value="Upload!" id=SUBMIT1 name=SUBMIT1>
				</td>
				<td align="right" width="160"></td>
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
	Set UploadProgress = nothing
end if
%>
