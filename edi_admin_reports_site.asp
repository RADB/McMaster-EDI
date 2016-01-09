<!-- #include virtual="/shared/admin_security.asp" -->
<%
' if the user has not logged in they will not be able to see the page
if blnSecurity then
%>
<html>
	<!-- #include virtual="/shared/head.asp" -->	
<body>
	<!-- #include virtual="/shared/page_header.inc" -->
	<%
	' open edi connection
	'call open_adodb(conn, "DATA")
    call open_adodb(conn, "MACEDI")
	
	%>
	<form name="site_report" method="POST" action="edi_admin_reports.asp" target="Reports"> 
		<a class="reglinkMaroon" href="edi_admin.asp">Home</a>&nbsp;<font class="regtextblack">></font>&nbsp;<font class="boldtextblack">Site Summary Report</font>
		<table border="1" bordercolor="006600" cellpadding="0" cellspacing="0" width="760">
		<tr>
			<td>
				<table border="0" cellpadding="0" cellspacing="0" width="760">
					<tr>
						<td align="right" width="440">
							<font class="headerBlue">Site Summary</font>
						</td>
						<td align="right">
							<input type="button" value="Exit" name="Exit" title="EXIT Screen" onClick="javascript:window.location='edi_admin.asp';">
							&nbsp;
						</td>
					<tr><td colspan="2"><%="<br/>" & strError%></td></tr>
				</table>
				<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">
					<tr>
						<td align="center">
							<font class="boldtextblack">Choose the site you wish to report on:</font>
							<select name="sites">
								<option value="0" selected>All Sites</option>
								<%
								'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
								' get the sites for the drop down box
								'//////////////////////////////////////////////////////////////////////	
								set rstSites = server.CreateObject("adodb.recordset")
	
								' open all languages
								rstSites.Open "SELECT DISTINCT intSiteID, sitename FROM sites_summary ORDER BY intSiteID", conn
								
								do while not rstSites.EOF  
									Response.Write "<option value=""" & rstSites("intSiteID") & """>" & right("000" & rstSites("intSiteID"),3) & " - " & rstSites("sitename") & "</option>"
									rstSites.MoveNext
								loop
								%>
							</select>
							<input type="hidden" value="site_summary.rpx" name="XML">
							<input type="submit" value="Generate" name="rpt" title="Generate Report" onClick="javascript:goWindow('','Reports','520','280','top=0,left=125,resizable=yes,scrollbars=no');">
						</td>
					</tr>
				</table>
				<br/> 
			</td>
		</tr>
		</table>
		<br/> 
	</form>
	
	<!-- #include virtual="/shared/page_footer.inc" -->
</body>
</html>
<%
	call close_adodb(rstSites)
	call close_adodb(conn)
end if
%>
