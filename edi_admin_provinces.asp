<!-- #include virtual="/shared/admin_security.asp" -->
<%
' if the user has not logged in they will not be able to see the page
if blnSecurity then
%>
<html>
<head>	
	<title>Provinces Look Up</title>
	<!-- Start CSS files-->
		<link rel="stylesheet" type="text/css" href="../Styles/edi.css">
	<!-- End CSS files -->
	
	<script Language="JavaScript">
	<!--
		function goForm(intItem)
		{
			if ((document.forms.Provinces.english.value.length == 0 || document.forms.Provinces.french.value.length == 0) && intItem < 0)
				alert('You must enter both the english and french values for the Province!!');
			else
			{
				document.forms.Provinces.item.value = intItem;
				document.forms.Provinces.submit();
			}
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

	if Request.form("item").Count > 0 then 
		intItem = Request.form("item")		
		if intItem < 0 then 
			strSql = "INSERT INTO [LU Provinces](english,french) VALUES(" & checknull(Request.Form("english")) & "," & checknull(Request.Form("french")) & ")"
		else
			strSql = "UPDATE [LU Provinces] SET english = " & checknull(Request.Form("english" & intItem)) & ",french = " & checknull(Request.Form("french" & intItem)) & " WHERE pid = " & Request.Form("pid" & intItem)
		end if 
		'Response.Write strSql
		' insert or update the record		
		conn.execute(strSql)
	end if 
	
	set rstProvinces = server.CreateObject("adodb.recordset")
	
	strSql = "SELECT pid, english, french FROM [LU Provinces] ORDER BY english"
	
	'open the recordset
	rstProvinces.Open strSql, conn
	%>
	<form name="Provinces" method="POST" action="edi_admin_Provinces.asp"> 
		<a class="reglinkMaroon" href="edi_admin.asp">Home</a>&nbsp;<font class="regtextblack">></font>&nbsp;<font class="boldtextblack">Provinces Look Up</font>
		<table border="1" bordercolor="006600" cellpadding="0" cellspacing="0" width="760">
		<tr><td>
			<table border="0" cellpadding="0" cellspacing="0" width="760">
				<tr>
					<td align="right" width="460">
						<font class="headerBlue">Provinces Look Up</font>
					</td>
					<td align="right">
						<input type="button" value="Exit" name="Exit" title="EXIT Screen" onClick="javascript:window.location='edi_admin.asp';">
						&nbsp;
					</td>
				<tr><td colspan="2"><br/></td></tr>
			</table>
			<table border="0" cellpadding="0" cellspacing="0" width="500" align="center">
			<tr>
				<td align="center"></td>
				<td><input type="text" size="15" name="english" value=""></td>
				<td><input type="text" size="20" name="french" value=""></td>
				<td align="left"><a href="javascript:goForm('-1');" class="regLinkMaroon">Add</a></td>
			</tr>
			<tr>
				<th class="subheaderMaroon" align="left" width="100"></th>
				<th class="subheaderMaroon" align="left" width="125">English</th>
				<th class="subheaderMaroon" align="left" width="150">French</th>
				<th class="subheaderMaroon" align="left" width="125"></th>
			</tr>
			<%
			' dumps all Provinces into an array
			aProvinces = rstProvinces.GetRows
			'*******************************
			' aProvinces(0,row) = Province ID (pid)
			' aProvinces(1,row) = English
			' aProvinces(2,row) = French
			'*******************************
			for intProvince = 0 to ubound(aProvinces,2)
				' holds the Province id
				Response.Write "<tr><td align=""center""></td>"
				' holds english name for the Province
				Response.Write "<td><input type=""text"" size=""15"" name=""english" & intProvince & """ value=""" & aProvinces(1,intProvince) & """></td>"
				' holds french name for the Province
				Response.Write "<td><input type=""text"" size=""20"" name=""french" & intProvince & """ value=""" & aProvinces(2,intProvince) & """></td>"
				' hidden variable holds original Province id for updating purposes
				Response.Write "<td align=""left""><a href=""javascript:goForm(" & intProvince & ");"" class=""regLinkMaroon"">Update</a><input type=""hidden"" name=""pid" & intProvince & """ value=""" & aProvinces(0,intProvince) & """></td></tr>"
			next
			%>
			</table>
			<br/> 
			</td>
		</tr>
		</table>
		<br/> 
		<input type="hidden" name="item" value="">
	</form>
	
	<!-- #include virtual="/shared/page_footer.inc" -->
</body>
</html>
<%
	call close_adodb(rstProvinces)
	call close_adodb(conn)
end if

function checknull(strTemp)
	if isnull(strTemp) or len(strTemp) = 0 then 
		checknull = "null"
	else
		checknull = "'" & replace(strTemp,"'","''") & "'"
	end if 
end function
%>
