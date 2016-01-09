<!-- #include virtual="/shared/security.asp" -->
<%
' if the user has not logged in they will not be able to see the page
if blnSecurity then
	'call open_adodb(conn,"TABLES")
    call open_adodb(conn, "MACEDI")
	set rstTables = server.CreateObject ("adodb.recordset")
	
	strSql = "SELECT 
%>

<html>
<!-- #include virtual="/shared/head.asp" -->	
<body>
	<!-- #include virtual="/shared/page_header.inc" -->
	<br />
	
	<table width="760" cellpadding="2" cellspacing="2" border="1">
		<tr>
			<td align="left" width="253">
				<!--<input type="text" id="site" name="site">
				<br>-->
				<font class="boldTextBlack">Site:</font>
			</td>
			<td align="center" width="253">
				<!--<input type="text" id="school" name="school">
				<br>-->
				<font class="boldTextBlack">School:</font>
			</td>
			<td align="center" width="253">
				<!--<input type="text" id="teacher" name="teacher">
				<br>-->
				<font class="boldTextBlack">Teacher:</font>
			</td>
		</tr>
		<tr><td colspan="3"><br /></td></tr>
		<tr>
			<td align="left" width="253">
				&nbsp;&nbsp;
				<font class="boldTextBlack">Class Size</font>
				&nbsp;
				<input type="text" id="size" name="size">
			</td>
			<td align="left" colspan="2" width="506">
				<font class="boldTextBlack">Records Completed</font>
				&nbsp;
				<input type="text" id="complete" name="complete">
			</td>
	</table>
	<!-- #include virtual="/shared/page_footer.inc" -->
</body>
</html>
<%
end if
%>