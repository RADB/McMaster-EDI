<!-- #include virtual="/shared/admin_security.asp" -->
<%
on error resume next
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
	
		<script language="JavaScript">
	<!--
		function goForm(intItem)
		{
			/*if (document.forms.Classes.code.value.length == 0 && intItem < 0) 
				alert('You must specify a class code value!!');
			else if (isNaN(document.forms.Classes.code.value) && intItem < 0)
				alert('The class code value can only contain numeric values!!');
			else
			{*/
				document.forms.Classes.item.value = intItem;
				document.forms.Classes.submit();
			//}
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
			strSql = "INSERT INTO [LU Classes](intClassID,English,French) VALUES(" & Request.Form("code") & "," & checknull(Request.Form("english")) & "," & checknull(Request.Form("french")) & ")"
		else
			strSql = "UPDATE [LU Classes] SET intClassid = " & Request.Form("code" & intItem) & ",English = " & checknull(Request.Form("english" & intItem)) & ",French = " & checknull(Request.Form("french" & intItem)) & " WHERE intClassID = " & Request.Form("cid" & intItem)
		end if 
		'Response.Write strSql
		' insert or update the record		
		conn.execute(strSql)
		
		if conn.errors.count > 0 AND err.number <> 0 then 
			Response.Write "<font class=""boldtextRed"">Your request could not be completed.  </font>"
			Response.Write "<br /><br />"
			Response.Write "<font class=""boldtextBlack"">The language code ""</font><font class=""boldtextRed"">" 
			if intItem < 0 then 
				Response.Write right("000" & Request.Form("code"),3) 
			else
				Response.Write right("000" & Request.Form("code" & intItem),3) 
			end if 
			Response.Write "</font><font class=""boldtextBlack"">"" is already in use.</font>"
			'Response.Write err.number & " - " & err.Description
			' 3001 - Arguments are of the wrong type, are out of acceptable range, or are in conflict with one another.
			' -2147217900 - [Microsoft][ODBC Microsoft Access Driver] The changes you requested to the table were not successful because they would create duplicate values in the index, primary key, or relationship. Change the data in the field or fields that contain duplicate data, remove the index, or redefine the index to permit duplicate entries and try again. 
			' -2147217900 - [Microsoft][ODBC Microsoft Access Driver] Syntax error in INSERT INTO statement. 
		end if 
	end if 
	
	set rstClasses = server.CreateObject("adodb.recordset")
	
	strSql = "SELECT intClassID, English, French FROM [LU Classes] ORDER BY intClassID"
	
	'open the recordset
	rstClasses.Open strSql, conn
	%>
	<form name="Classes" method="POST" action="edi_admin_classes.asp"> 
		<a class="reglinkMaroon" href="edi_admin.asp">Home</a>&nbsp;<font class="regtextblack">></font>&nbsp;<font class="boldtextblack">Class Times Look Up</font>
		<table border="1" bordercolor="006600" cellpadding="0" cellspacing="0" width="760">
		<tr><td>
			<table border="0" cellpadding="0" cellspacing="0" width="760">
				<tr>
					<td align="right" width="450">
						<font class="headerBlue">Class Times Look Up</font>
					</td>
					<td align="right">
						<input type="button" value="Exit" name="Exit" title="EXIT Screen" onClick="javascript:window.location='edi_admin.asp';">
						&nbsp;
					</td>
				<tr><td colspan="2"><br/></td></tr>
			</table>
			<table border="0" cellpadding="0" cellspacing="0" width="760" align="center">
			<!--<tr>
				<td align="center"><input type="text" size="5" name="code" value="" maxlength="3"></td>
				<td><input type="text" size="45" name="english" value=""></td>
				<td align="left"><a href="javascript:goForm('-1');" class="regLinkMaroon">Add</a></td>
			</tr>-->
			<tr>
				<th class="subheaderMaroon" align="center" width="125">Value</th>
				<th class="subheaderMaroon" align="left" width="275">English</th>
				<th class="subheaderMaroon" align="left" width="275">French</th>
			</tr>
			<%
			' dumps all languages into an array
			aClasses = rstClasses.GetRows
			'*******************************
			' aClasses(0,row) = value
			' aClasses(1,row) = description
			'*******************************
			for intClass = 0 to ubound(aClasses,2)
				' holds the language id
				Response.Write "<tr><td align=""center""><input type=""text"" maxlength=""1"" size=""5"" name=""code" & intClass & """ value=""" & aClasses(0,intClass) & """ readonly></td>"
				' holds description for the class time
				Response.Write "<td><input type=""text"" size=""35"" name=""english" & intClass & """ value=""" & aClasses(1,intClass) & """></td>"
				Response.Write "<td><input type=""text"" size=""35"" name=""french" & intClass & """ value=""" & aClasses(2,intClass) & """></td>"
				' hidden variable holds original class id for updating purposes
				Response.Write "<td align=""left""><a href=""javascript:goForm(" & intClass & ");"" class=""regLinkMaroon"">Update</a><input type=""hidden"" name=""cid" & intClass & """ value=""" & aClasses(0,intClass) & """></td></tr>"
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
	call close_adodb(rstClasses)
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
