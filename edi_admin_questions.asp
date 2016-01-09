<!-- #include virtual="/shared/admin_security.asp" -->
<%
' if the user has not logged in they will not be able to see the page
if blnSecurity then
%>
<html>
<head>
    <!-- added UTF8 Encoding to get rid of funny characters -->
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" /> 

	<!-- Start CSS files-->
		<link rel="stylesheet" type="text/css" href="../Styles/edi.css">
	<!-- End CSS files -->
	
	<script Language="JavaScript">
	<!--
		function goForm(strSection)
		{	
			// set the section value
			document.forms.Questions.section.value = strSection;
			document.forms.Questions.submit();
		}
	//-->
	</script>
</head>
<body>
	<!-- #include virtual="/shared/page_header.inc" -->
	<%
	' open edi connection
	'call open_adodb(conn, "EDI")
	call open_adodb(conn, "MACEDI")
	%>
	<form name="Questions" method="POST" action="edi_admin_questions_sections.asp"> 
		<a class="reglinkMaroon" href="edi_admin.asp">Home</a>&nbsp;<font class="regtextblack">></font>&nbsp;<font class="boldtextblack">Questions Data Bank</font>
		<table border="1" bordercolor="006600" cellpadding="0" cellspacing="0" width="760">
		<tr><td>
			<table border="0" cellpadding="0" cellspacing="0" width="760">
				<tr>
					<td align="right" width="480">
						<font class="headerBlue">Questions Data Bank</font>
					</td>
					<td align="right">
						<input type="button" value="Exit" name="Exit" title="EXIT Screen" onClick="javascript:window.location='edi_admin.asp';">
						&nbsp;
					</td>
				</tr>
			</table>
			
			<table border="0" cellpadding="0" cellspacing="0" width="500" align="center">
			<tr><td><br/></td></tr>
			<tr>
				<td>
					<font class="boldTextBlack">Which sections questions do you wish to edit?</font>
				</td>
			</tr>
			<tr>
				<td>
					<font class="regTextBlack">1) </font><a href="javascript:goForm('_Demographics');" class="regLinkMaroon">Child Demographics</a>
				</td>
			</tr>
			<tr>
				<td>
					<font class="regTextBlack">2) </font><a href="javascript:goForm('A');" class="regLinkMaroon">Section A</a>
				</td>
			</tr>
			<tr>
				<td>
					<font class="regTextBlack">3) </font><a href="javascript:goForm('B');" class="regLinkMaroon">Section B</a>
				</td>
			</tr>
			<tr>
				<td>
					<font class="regTextBlack">4) </font><a href="javascript:goForm('C');" class="regLinkMaroon">Section C</a>
				</td>
			</tr>
			<tr>
				<td>
					<font class="regTextBlack">5) </font><a href="javascript:goForm('D');" class="regLinkMaroon">Section D</a>
				</td>
			</tr>
			<tr>
				<td>
					<font class="regTextBlack">6) </font><a href="javascript:goForm('E');" class="regLinkMaroon">Section E</a>
				</td>
			</tr>
			</table>
			<br/> 
			</td>
		</tr>
		</table>
		<br/> 
		<input type="hidden" name="section" value="">	
	</form>
	
	<!-- #include virtual="/shared/page_footer.inc" -->
</body>
</html>
<%
	call close_adodb(conn)
end if
%>
