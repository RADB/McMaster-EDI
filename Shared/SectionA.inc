	<input type="hidden" name="CurrentSection" value="A" />
	<%
	' get the demographic data
	strSql = "SELECT * FROM sectionA WHERE strEDIID = '" & strEDIID & "'"

	' student recordset
	set rstStudent = server.createobject("adodb.recordset")
		
	rststudent.Open strSql, conn
	if not rstStudent.eof then 	
									
		' get all the demographic questions
		strSql = "SELECT A.English, A.French, A.isSectionHeader, A.[Section], A.Question, A.[Option], A.isHeader,  C.[Language], C.intOptions, C.Col1, C.Col2, C.Col3, C.Col4, C.Col5, C.Col6 FROM Page_SectionA A  LEFT JOIN [Column Headers] C ON A.HeaderType = C.HID WHERE C.[Language]='" & strLanguage & "' OR C.[Language] Is Null ORDER BY A.isSectionHeader Desc, A.[Section], A.Question, A.[Option], A.isHeader"
  
  'open the Section A questions
  rstData.Open strSql, conn,1

  ' header
  Response.Write "<br />&nbsp;<input type=""button"" onclick=""javascript:goWindow('" & strLink & "','Guide','490','500','top=0,left=125,resizable=yes');"" name=""btnGuide"" value=""GUIDE"" title=""GUIDE"" />&nbsp;&nbsp;<a href=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','_Demographics','A');"" class=""bigLinkRed"">"
		if strLanguage="English" Then 
			response.write "Demographics"
		else
			response.write "Démographique"
		end if 
		response.write "</a>&nbsp;&nbsp; <font class=""subheaderBlue"">Section A</font>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','B','A')"" class=""bigLinkRed"">Section B</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','C','A')"" class=""bigLinkRed"">Section C</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','D','A')"" class=""bigLinkRed"">Section D</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','E','A')"" class=""bigLinkRed"">Section E</a>"
		
		if NOT blnLock then 			
			response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type=""button"" onclick=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','A','A')"" name=""btnSave"" value=""" & lblSaveEDI & """ title=""" & lblSaveEDI & """ />"					
		end if 
		response.write "<br />"					
									
		' Menu
		Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""750"" align=""center"">"																					
		Response.Write "<tr><td colspan=""3""><br /></td></tr>"
		Response.write "<tr><td width=""24%"" align=""right"">&nbsp;" & strPrevious & "</td><td width=""53%"" align=""center"" valign=""middle""><font class=""headerBlue"">" &  rstData(strLanguage) & "</font></td><td width=""23%"" align=""left"">&nbsp;" & strNext & "</td></tr>"
		Response.Write "<tr><td colspan=""3""><br /></td></tr>"
		Response.Write "</table>"
									
		rstData.MoveNext 
									
		' bln for inner table existance
		blnTable = false
		intOptions = 0 
		intRow = 0 
									
		' table will be <spacer Column><question><6 columns>
		Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""750"" align=""center"">"																					
		do while not rstData.EOF 
			' write the header row if it is a header
			if rstData("isHeader") then 
				' check to see if the last table needs to be closed
				if blnTable then 
					intRow = 0
					Response.Write "</table>"
					Response.Write "<tr><td><br /></td></tr>"
				end if 
											
				' new section
				Response.Write "<tr><td>"
											
				' inner table for each section
				Response.Write "<table border=""1"" cellpadding=""0"" cellspacing=""0"" width=""750"" align=""center"">"																					
				Response.Write "<tr>"		
				Response.Write "<td align=""left"" colspan=""2""><font class=""subHeaderBlue"">" & rstData(strLanguage) & "</font></td>" 
											
				' get the number of options 
				intOptions = rstData("intOptions")
				for intCol = 1 to intoptions
					Response.Write "<td align=""center"" valign=""middle"" width=""70""><font class=""boldTextBlack"">" & rstData("Col" & intCol) & "</font></td>"
					' store the header values for the title below 2005 09 15
					aHeader(intCol) = rstData("Col" & intCol)
				next 
				Response.Write "</tr>"
											
				' set the inner table to true
				blnTable = true 
			else
				' if it is just a question and no header then write a new table
				if intOptions = 0 then 
					Response.Write "<tr><td>"	
					Response.Write "<table border=""1"" cellpadding=""0"" cellspacing=""0"" width=""750"" align=""center"">"																					
						Response.Write "<tr bgcolor=""whitesmoke""><td align=""left"" valign=""top"" width=""35""><font class=""boldTextBlack"">&nbsp;&nbsp;" 
						' only write the questions
						if rstData("option") = 0 then 
							if rstData("question") < 10 then 
								Response.Write "&nbsp;&nbsp;"
							end if 
							strQuestion = "intQ" & rstData("Question")
							Response.Write rstData("question") & " &nbsp;&nbsp;</font></td><td align=""left""><font class=""boldTextBlack"">" &  rstData(strLanguage) & "</font>"
							response.write "<p align=""right""><select name=""days"">"
              response.write "<option value=""-1""""></option>"
							for intRow = 0 to 99 
								response.write "<option value=""" & intRow & """"
								if intRow = int(rstStudent(strQuestion)) then 
									response.write " selected=""selected"""
								end if 
								response.write ">" & intRow & "</option>"
							next 
							response.write "</select>."
							
							response.write "<select name=""days2"">"
              response.write "<option value=""-1""""></option>"
							for intRow = 0 to 5 step 5 
								response.write "<option value=""" & intRow & """"
                if rstStudent(strQuestion) <> "" then   
								  if intRow = int(right(formatnumber(rstStudent(strQuestion),1),1)) then 
									  response.write " selected=""selected"""
  								end if
                end if
								response.write ">" & intRow & "</option>"
							next 
							response.write "</select>"							
							response.write "</p>"
						end if
																	
						Response.Write "</td>"
						response.write getHelpLink(session("Province"), "SectionA", strQuestion)
						Response.Write "</tr>"
					Response.Write "</table>"
					Response.Write "</td></tr>"
					Response.Write "<tr><td><br /></td></tr>"
				else
					intRow = introw + 1
					if intRow mod 2 = 1 then 
						strColour = "whitesmoke"
					else
						strColour = "white"
					end if 
					Response.Write "<tr bgcolor=""" & strColour & """><td align=""left"" valign=""top"" width=""35"">&nbsp;<font class=""boldTextBlack"">" 
					' only write the questions
					if rstData("option") = 0 then 
						if rstData("question") < 10 then 
							Response.Write "&nbsp;&nbsp;"
						end if 
						Response.Write rstData("question") & " &nbsp;&nbsp;</font></td><td><font class=""boldTextBlack"">" &  rstData(strLanguage) & "</font>"
						strQuestion = "intQ" & rstData("Question")
					end if
																	
					'select case rstData("question")
														
					'end select
					Response.Write "</td>"
												
					' write the radio option buttons
					for intCol = 1 to intoptions
						' Response.Write "<td align=""center"" valign=""middle"" width=""70""><input type=""radio"" title=""" & aHeader(intCol) &""" name=""" & strQuestion & """ value=""" & intCol & """"
						' if rstStudent(strQuestion) = intCol then 
							' response.write " checked=""checked"""
						' end if 
						' response.write " /></td>"
						Response.Write "<td align=""center"" valign=""middle"" width=""70"">"
						response.write buildRadioButton("horizontal", strQuestion, aHeader(intcol),intcol, rstStudent(strQuestion),"","")
						response.write "</td>"
					next 
					response.write getHelpLink(session("Province"), "SectionA", strQuestion)
					Response.Write "</tr>"
				end if 										
			end if 
			rstData.movenext
		loop
		' close the last inner table if it is open
		if blnTable then 
			Response.Write "</table>"
		end if 
									
		' end the row
		Response.Write "</td></tr>"
		Response.Write "</table>"	
	
										
		' header
		Response.Write "<br />&nbsp;<input type=""button"" onclick=""javascript:goWindow('" & strLink & "','Guide','490','500','top=0,left=125,resizable=yes');"" name=""btnGuide"" value=""GUIDE"" title=""GUIDE"" />&nbsp;&nbsp;<a href=""javascript:goSaveEDI('" & left(strEDIID,3) & "','" & mid(strEDIID,4,3) & "','" & mid(strEDIID,7,2) & "','" & mid(strEDIID,9,1) & "','" & right(strEdiID,2) & "','_Demographics','A');"" class=""bigLinkRed"">"
		if strLanguage="English" Then 
			response.write "Demographics"
		else
			response.write "Démographique"
		end if 
		response.write "</a>&nbsp;&nbsp; <font class=""subheaderBlue"">Section A</font>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','B','A')"" class=""bigLinkRed"">Section B</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','C','A')"" class=""bigLinkRed"">Section C</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','D','A')"" class=""bigLinkRed"">Section D</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','E','A')"" class=""bigLinkRed"">Section E</a>"
		
		if NOT blnLock then 			
			response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type=""button"" onclick=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','A','A')"" name=""btnSave"" value=""" & lblSaveEDI & """ title=""" & lblSaveEDI & """ />"					
		end if 
		response.write "<br />"					
		
	else
		strError = "<font class=""regtextred"">No data on child - " & strEDIID & "</font>"
	end if
	
	' close and kill the additional questions recordset
	call close_adodb(rstStudent)
	%>