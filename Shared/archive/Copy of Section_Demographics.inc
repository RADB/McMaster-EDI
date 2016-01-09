	<%
	' get the demographic data
	strSql = "SELECT * FROM Student_Demographics WHERE strEDIID = '" & strEDIID & "' ORDER BY strEDIID"
		
	rstData.Open strSql, conn_data
	if not rstData.eof then 
		aStudent = rstData.getrows
	else
		strError = "<font class=""regtextred"">No data on child - " & strEDIID & "</font>"
	end if
		
	' close the recordset
	rstData.close

	' header							
	Response.Write "<br />&nbsp;<input type=""button"" onClick=""javascript:goWindow('" & strLink & "','Guide','490','500','top=0,left=125,resizable=yes');"" name=""btnGuide"" value=""GUIDE"" title=""GUIDE"">&nbsp;&nbsp;<font class=""subheaderBlue"">"	
	if strLanguage="English" Then 
		response.write "Demographics"
	else
		response.write "Démographique"
	end if 
	response.write "</font>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & left(strEDIID,3) & "','" & mid(strEDIID,4,3) & "','" & mid(strEDIID,7,2) & "','" & mid(strEDIID,9,1) & "','" & right(strEdiID,2) & "','A','_Demographics');"" class=""bigLinkRed"">Section A</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & left(strEDIID,3) & "','" & mid(strEDIID,4,3) & "','" & mid(strEDIID,7,2) & "','" & mid(strEDIID,9,1) & "','" & right(strEDIID,2)& "','B','_Demographics')"" class=""bigLinkRed"">Section B</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & left(strEDIID,3) & "','" & mid(strEDIID,4,3) & "','" & mid(strEDIID,7,2) & "','" & mid(strEDIID,9,1) & "','" & right(strEDIID,2)& "','C','_Demographics')"" class=""bigLinkRed"">Section C</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & left(strEDIID,3) & "','" & mid(strEDIID,4,3) & "','" & mid(strEDIID,7,2) & "','" & mid(strEDIID,9,1) & "','" & right(strEDIID,2)& "','D','_Demographics')"" class=""bigLinkRed"">Section D</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & left(strEDIID,3) & "','" & mid(strEDIID,4,3) & "','" & mid(strEDIID,7,2) & "','" & mid(strEDIID,9,1) & "','" & right(strEDIID,2)& "','E','_Demographics')"" class=""bigLinkRed"">Section E</a>"
	if NOT blnLock then 
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type=""button"" onClick=""javascript:goSaveEDI('" & left(strEDIID,3) & "','" & mid(strEDIID,4,3) & "','" & mid(strEDIID,7,2) & "','" & mid(strEDIID,9,1) & "','" & right(strEDIID,2)& "','_Demographics','_Demographics')"" name=""btnSave"" value=""" & lblSaveEDI & """ title=""" & lblSaveEDI & """>"
	end if 
	response.write "<br />"
										
	' get all the demographic questions - changed May 24, 2004
	' strSql = "SELECT question, english, french,option FROM Page_Section_Demographics WHERE Question>0 ORDER BY Question, Option"
	strSql = "SELECT question, english, french,option FROM Page_Section_Demographics ORDER BY Question, Option"
	
	' removed May 24, 2004
	'Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""750"" align=""center"">"																					
	'Response.Write "<tr><td colspan=""2""><br /></td></tr>"
	'
	'if strLanguage = "English" then 
	'	Response.write "<tr><td colspan=""2"" align=""center""><font class=""headerBlue"">Child Demographics<font></td></tr>"
	'else
	'	Response.write "<tr><td colspan=""2"" align=""center""><font class=""headerBlue"">Enfant Démographique<font></td></tr>"
	'end if 
	
	'Response.Write "<tr><td colspan=""2""><br /></td></tr>"
	'Response.Write "</table>"
						
	'open the demographic questions 
	rstData.Open strSql, conn_Tables,1  

	' added May 24, 2004	
	if not rstData.eof then 
		Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""750"" align=""center"">"																					
		Response.Write "<tr><td colspan=""3""><br /></td></tr>"
		Response.write "<tr><td width=""13%"" align=""right"">&nbsp;" & strPrevious & "</td><td width=""74%"" align=""center"" valign=""middle""><font class=""headerBlue"">" &  rstData(strLanguage) & "<font></td><td width=""13%"" align=""left"">&nbsp;" & strNext & "</td></tr>"
		Response.Write "<tr><td colspan=""3""><br /></td></tr>"
		Response.Write "</table>"
		
		rstdata.movenext
	end if 
	
	' Menu
	Response.Write "<table border=""1"" cellpadding=""0"" cellspacing=""0"" width=""750"" align=""center"">"																					
										
	do while not rstData.EOF 
		intRow = rstData("question")
		if intRow mod 2 = 1 then 
			strColour = "whitesmoke"
		else
			strColour = "white"
		end if 
		Response.Write "<tr bgcolor=""" & strColour & """>"
		response.write "<td align=""left"" valign=""top"">&nbsp;<font class=""boldTextBlack"">" 
		' only write the questions
		if rstData("option") = 0 then 
			if rstData("question") < 10 then 
				Response.Write "&nbsp;&nbsp;"
			end if 
			Response.Write rstData("question") & " ) &nbsp;&nbsp;</font></td><td><font class=""boldTextBlack"">" &  rstData(strLanguage) & ": </font>"
		end if
											

											
		select case rstData("question")
			' Date of completion
			case 1
				Response.Write "<font class=""lrgRegTextBlack"">"
				if not isnull(aStudent(8,0)) then 
					if strLanguage = "English" then 
						Response.Write right("00" & day(aStudent(8,0)),2) & "-" & monthname(datepart("m",aStudent(8,0)),true) & "-" & year(aStudent(8,0)) 		
					else
						Response.Write right("00" & day(aStudent(8,0)),2) & "-" & left(French_Month(datepart("m",aStudent(8,0))),3) & "-" & year(aStudent(8,0)) 		
					end if 						
				else
					Response.write "Incomplete"
				end if 
				Response.Write "</font>"
			' DOB
			case 2
				intmonth = month(aStudent(5,0))
				intday = day(aStudent(5,0))
				intyear = year(aStudent(5,0))	
				%>
				<select name="DOBday">
					<option value="-1"></option>
				<%
				for introw = 1 to 31
					Response.Write "<option value = """ & intRow & """"
					if intDay = intRow then 
						Response.write " selected"
					end if 
					' write the day
					Response.Write ">" & intRow & "</option>"
				next
				%>
				</select>
				<select name="DOBmonth">
					<option value="-1"></option>
				<%
				for introw = 1 to 12
					Response.Write "<option value = """ & intRow & """"
					if intMonth = intRow then 
						Response.write " selected"
					end if 
					' write the month
					if strLanguage = "English" then 
						Response.Write ">" & monthname(intRow,false) & "</option>"
					else
						Response.Write ">" & french_Month(intRow) & "</option>"
					end if 
				next
				%>
				</select>
				<select name="DOByear">
					<option value="-1"></option>
				<%
				for introw = 1 to 5
					Response.Write "<option value = """ & intRow + 1994 & """"
					if intYear = intRow + 1994 then 
						Response.write " selected"
					end if 
					' write the day
					Response.Write ">" & intRow + 1994 & "</option>"
				next
				%>
				</select>
				<%
			' Language Section
			case 3
				Response.Write "<select name=""intLanguage1"">"
				Response.Write "<option value=""-1""></option>"										
				for intRow = 0 to ubound(aLanguages,2)															
					' show the language
					Response.write "<option value=""" & aLanguages(0,introw) & """"
					if aStudent(14,0) = aLanguages(0,intRow) then 
						Response.Write " selected"
					end if 
					Response.Write ">" & aLanguages(intLanguage,introw) & "</option>"
				next
				Response.Write "</select>"
													
				Response.Write "<select name=""intLanguage2"">"
				Response.Write "<option value=""-1""></option>"										
				for intRow = 0 to ubound(aLanguages,2)															
					' show the language
					Response.write "<option value=""" & aLanguages(0,introw) & """"
					if aStudent(15,0) = aLanguages(0,intRow) then 
						Response.Write " selected"
					end if 
					Response.Write ">" & aLanguages(intLanguage,introw) & "</option>"
				next
													
			' Class Assignment
			case 4
				rstData.MoveNext 
													
				Response.Write "<select name=""intClassAssignment"">"
				Response.Write "<option value=""-1""></option>"										
				do while rstData("question") = 4															
					' show the language
					Response.write "<option value=""" & rstData("option") & """"
					if aStudent(16,0) = rstData("option") then 
						Response.Write " selected"
					end if 
					Response.Write ">" & rstData(strLanguage) & "</option>"
					rstData.MoveNext 
				loop
				Response.Write "</select>"
													
				rstData.MovePrevious 
			' Class Type
			case 5
				rstData.MoveNext 
													
				Response.Write "<select name=""intClassType"">"
				Response.Write "<option value=""-1""></option>"										
				do while rstData("question") = 5															
					' show the language
					Response.write "<option value=""" & rstData("option") & """"
					if aStudent(17,0) = rstData("option") then 
						Response.Write " selected"
					end if 
					Response.Write ">" & rstData(strLanguage) & "</option>"
					rstData.MoveNext 
				loop
				Response.Write "</select>"
													
				rstData.MovePrevious 
			' Student Status
			case 6
				rstData.MoveNext 
													
				Response.Write "<select name=""intStatus"">"
				Response.Write "<option value=""-1""></option>"										
				do while rstData("question") = 6															
					' show the language
					Response.write "<option value=""" & rstData("option") & """"
					if aStudent(18,0) = rstData("option") then 
						Response.Write " selected"
					end if 
					Response.Write ">" & rstData(strLanguage) & "</option>"
					rstData.MoveNext 
				loop
				Response.Write "</select>"
													
				rstData.MovePrevious 
			' Sex											
			case 7
				rstData.MoveNext 
													
				Response.Write "<select name=""intSex"">"
				Response.Write "<option value=""-1""></option>"										
				do while rstData("question") = 7															
					' show the language
					Response.write "<option value=""" & rstData("option") & """"
					if aStudent(4,0) = rstData("option") then 
						Response.Write " selected"
					end if 
					Response.Write ">" & rstData(strLanguage) & "</option>"
					rstData.MoveNext 
				loop
				Response.Write "</select>"
													
				rstData.MovePrevious 
			' Postal Code
			case 8
				Response.Write "<input type=""text"" size=""10"" name=""strPostal"" value=""" & aStudent(6,0) & """ maxlength=""7"">" 
			' ESL
			case 9
				rstData.MoveNext 
													
				Response.Write "<select name=""intESL"">"
				Response.Write "<option value=""-1""></option>"										
				do while rstData("question") = 9															
					' show the language
					Response.write "<option value=""" & rstData("option") & """"
					if aStudent(19,0) = rstData("option") then 
						Response.Write " selected"
					end if 
					Response.Write ">" & rstData(strLanguage) & "</option>"
					rstData.MoveNext 
				loop
				Response.Write "</select>"
													
				rstData.MovePrevious 
			' Special Needs
			case 10
				rstData.MoveNext 
													
				Response.Write "<select name=""intSpecial"">"
				Response.Write "<option value=""-1""></option>"										
				do while rstData("question") = 10															
					' show the language
					Response.write "<option value=""" & rstData("option") & """"
					if aStudent(20,0) = rstData("option") then 
						Response.Write " selected"
					end if 
					Response.Write ">" & rstData(strLanguage) & "</option>"
					rstData.MoveNext 
				loop
				Response.Write "</select>"
													
				rstData.MovePrevious 
			' Aboriginal
			case 11
				rstData.MoveNext 
													
				Response.Write "<select name=""intAboriginal"">"
				Response.Write "<option value=""-1""></option>"										
				do while rstData("question") = 11															
					' show the language
					Response.write "<option value=""" & rstData("option") & """"
					if aStudent(21,0) = rstData("option") then 
						Response.Write " selected"
				end if 
					Response.Write ">" & rstData(strLanguage) & "</option>"
					rstData.MoveNext 
				loop
				Response.Write "</select>"
													
				rstData.MovePrevious 
			' French Immersion
			case 12
				rstData.MoveNext 
													
				Response.Write "<select name=""intFrench"">"
				Response.Write "<option value=""-1""></option>"										
				do while rstData("question") = 12															
					' show the language
					Response.write "<option value=""" & rstData("option") & """"
					if aStudent(22,0) = rstData("option") then 
						Response.Write " selected"
					end if 
					Response.Write ">" & rstData(strLanguage) & "</option>"
					rstData.MoveNext 
				loop
				Response.Write "</select>"
													
				rstData.MovePrevious 
			' Other Immersion
			case 13
				rstData.MoveNext 
													
				Response.Write "<select name=""intOther"">"
				Response.Write "<option value=""-1""></option>"										
				do while not rstData.eof
					if rstData("question") = 13 then 															
						' show the language
						Response.write "<option value=""" & rstData("option") & """"
						if aStudent(23,0) = rstData("option") then 
							Response.Write " selected"
						end if 
						Response.Write ">" & rstData(strLanguage) & "</option>"
					end if 
					rstData.MoveNext 
				loop
				Response.Write "</select>"
													
				rstData.MovePrevious 
		end select
		Response.Write "</font></td>"
		Response.Write "</tr>"
		rstData.movenext
	loop
	Response.Write "</table>"	
	
	' header							
	Response.Write "<br />&nbsp;<input type=""button"" onClick=""javascript:goWindow('" & strLink & "','Guide','490','500','top=0,left=125,resizable=yes');"" name=""btnGuide"" value=""GUIDE"" title=""GUIDE"">&nbsp;&nbsp;<font class=""subheaderBlue"">"	
	if strLanguage="English" Then 
		response.write "Demographics"
	else
		response.write "Démographique"
	end if 
	response.write "</font>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & left(strEDIID,3) & "','" & mid(strEDIID,4,3) & "','" & mid(strEDIID,7,2) & "','" & mid(strEDIID,9,1) & "','" & right(strEdiID,2) & "','A','_Demographics');"" class=""bigLinkRed"">Section A</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & left(strEDIID,3) & "','" & mid(strEDIID,4,3) & "','" & mid(strEDIID,7,2) & "','" & mid(strEDIID,9,1) & "','" & right(strEDIID,2)& "','B','_Demographics')"" class=""bigLinkRed"">Section B</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & left(strEDIID,3) & "','" & mid(strEDIID,4,3) & "','" & mid(strEDIID,7,2) & "','" & mid(strEDIID,9,1) & "','" & right(strEDIID,2)& "','C','_Demographics')"" class=""bigLinkRed"">Section C</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & left(strEDIID,3) & "','" & mid(strEDIID,4,3) & "','" & mid(strEDIID,7,2) & "','" & mid(strEDIID,9,1) & "','" & right(strEDIID,2)& "','D','_Demographics')"" class=""bigLinkRed"">Section D</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & left(strEDIID,3) & "','" & mid(strEDIID,4,3) & "','" & mid(strEDIID,7,2) & "','" & mid(strEDIID,9,1) & "','" & right(strEDIID,2)& "','E','_Demographics')"" class=""bigLinkRed"">Section E</a>"
	if NOT blnLock then 
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type=""button"" onClick=""javascript:goSaveEDI('" & left(strEDIID,3) & "','" & mid(strEDIID,4,3) & "','" & mid(strEDIID,7,2) & "','" & mid(strEDIID,9,1) & "','" & right(strEDIID,2)& "','_Demographics','_Demographics')"" name=""btnSave"" value=""" & lblSaveEDI & """ title=""" & lblSaveEDI & """>"
	end if 
	response.write "<br />"
	%>