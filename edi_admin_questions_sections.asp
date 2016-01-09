<!-- #include virtual="/shared/admin_security.asp" -->
<%
' if the user has not logged in they will not be able to see the page
if blnSecurity then
	' open edi connection
	'call open_adodb(conn, "TABLES")
    call open_adodb(conn, "MACEDI")
	set rstSection = server.CreateObject("adodb.recordset")

	strTable = "page_Section" & Request.form("section")
	
	' get the language
	if Request.Form("language").Count > 0 then 
		strLanguage = Request.Form("language")
	else
		strLanguage = session("language")
	end if 
	
	if Request.form("uid").Count > 0 then 
		intItem = Request.form("uid")		
		if intItem < 0 then 
			' insert if 
			'strSql = "INSERT INTO [" & strTable & "] (" & strLanguage & ") VALUES(" & checknull(Request.Form("english")) & "," & checknull(Request.Form("french")) & ")"
		else
			strSql = "UPDATE [" & strTable & "] SET " & strLanguage & " = " & checknull(Request.Form("q" & intItem)) & " WHERE iid = " & intItem
			
			' insert or update the record	
			conn.execute(strSql)
		end if 
		'Response.Write strSql
		
	end if 

%>
<html>
<head>
	<!-- Start CSS files-->
		<link rel="stylesheet" type="text/css" href="Styles/edi.css">
	<!-- End CSS files -->
	
	<script Language="JavaScript">
	<!--
		function goForm(intItem)
		{
			if (intItem >= 0)
			{
				var ilen = (eval('document.forms.Questions.q' + intItem + '.value.length'))
		
				if (ilen == 0)
				{
					alert('You cannot enter nothing for the value!!');
				}
				else
				{
					document.forms.Questions.uid.value = intItem;
					document.forms.Questions.submit();
				}
			}
			else
			{
				document.forms.Questions.uid.value = intItem;
				document.forms.Questions.submit();
			}
		}
	//-->
	</script>

</head>
<body>
	<!-- #include virtual="/shared/page_header.inc" -->
	<%	
	' issectionheader - only one - will be first
	' headertypes 
	' isheader  - top of header type
	' question - in order of numbered questions
	strSql = "SELECT * FROM [" & strTable & "] ORDER BY [isSectionHeader],[section],[isheader],[headertype],[question], [option]"
	rstSection.Open strSql, conn
	
	if rstSection.EOF then 
		strHeader = ""
	else
		strHeader = rstSection(strLanguage)
	end if
	%>
	<form name="Questions" method="POST" action="edi_admin_questions_sections.asp"> 
		<a class="reglinkMaroon" href="edi_admin.asp">Home</a>&nbsp;<font class="regtextblack">></font>&nbsp;<a class="reglinkMaroon" href="edi_admin_questions.asp">Questions Data Bank</a>&nbsp;<font class="regtextblack">></font>&nbsp;<font class="boldtextblack">Section <%=replace(Request.Form("section"),"_" ,"")%></font>
		<table border="1" bordercolor="006600" cellpadding="0" cellspacing="0" width="760">
		<tr>
			<td>
				<table border="0" cellpadding="0" cellspacing="0" width="760">
					<tr>
						<td align="right" width="490">
							<font class="headerBlue"><%=strHeader%></font>
						</td>
						<td align="right">
							<input type="button" value="Exit" name="Exit" title="EXIT Screen" onClick="javascript:window.location='edi_admin.asp';">
							&nbsp;
						</td>
					</tr>
				</table>
				
				<table border="0" cellpadding="0" cellspacing="0" width="700" align="center">
				<tr><td colspan="2"><br/></td></tr>
				<tr>
					<td colspan="2" align="center">
					<!-- default language set to English -->
						<%
						if strLanguage = "English" then 
							Response.Write "<INPUT type=""radio"" id=""language"" name=""language"" value=""English"" checked>"
							Response.Write "<font class=""boldtextblack"">English&nbsp;&nbsp;</font>"
							Response.Write "<INPUT type=""radio"" id=""language"" name=""language"" value=""French"" onClick=""javascript:goForm(-1);"">"
							Response.Write "<font class=""boldtextblack"">French&nbsp;&nbsp;</font>"
						elseif strLanguage = "French" then 
							Response.Write "<INPUT type=""radio"" id=""language"" name=""language"" value=""English"" onClick=""javascript:goForm(-1);"">"
							Response.Write "<font class=""boldtextblack"">English&nbsp;&nbsp;</font>"
							Response.Write "<INPUT type=""radio"" id=""language"" name=""language"" value=""French"" checked>"
							Response.Write "<font class=""boldtextblack"">French&nbsp;&nbsp;</font>"
						end if 
						%>
					</td>
				</tr>
				<tr><td colspan="2"><br/></td></tr>
				
				<%
				if not rstSection.EOF then 
					Response.Write "<tr><td align=""right"" nowrap>"
					' hidden id value
					Response.Write "<input type=""hidden"" name=""uid"" value=""" & rstSection("iid") & """>"
					' hidden section
					Response.Write "<input type=""hidden"" name=""section"" value=""" & Request.Form("section") & """>"
					Response.Write "<font class=""boldtextblack"">Header:&nbsp;&nbsp;</font>"
					Response.Write "<input type=""text"" size=""90"" name=""q" & rstSection("iid") & """ value=""" & strHeader & """>"
					Response.Write "</td>"
					Response.Write "<td align=""right"" width=""60""><a href=""javascript:goForm(" & rstSection("iid") & ");"" class=""regLinkMaroon"">Update</a></td>"
					Response.Write "</tr>"
					rstSection.MoveNext 
					do while not rstSection.eof
						Response.Write "<tr><td align=""right"">"
						if rstSection("isheader") then 
							Response.Write "<font class=""boldtextblack"">Sub:&nbsp;&nbsp;</font><input type=""text"" size=""85"" name=""q" & rstSection("iid") & """ value=""" & rstSection(strLanguage) & """>" 	
						else
							Response.Write "<font class=""boldtextblack"">Q" & rstSection("question")
							' write the option letter if there is one
							if rstSection("option") > 0 then 
								Response.Write chr(rstSection("option"))
							end if 
							 Response.Write ") </font><input type=""text"" size=""80"" name=""q" & rstSection("iid") & """ value=""" & rstSection(strLanguage) & """>" 	
						end if 
						Response.Write "<td align=""right"" width=""60""><a href=""javascript:goForm(" & rstSection("iid") & ");"" class=""regLinkMaroon"">Update</a></td>"
						Response.Write "</td></tr>"
						rstSection.MoveNext 
					loop
				end if 
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
	call close_adodb(conn)
end if
%>
