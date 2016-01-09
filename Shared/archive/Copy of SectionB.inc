	<%
	' get the demographic data
	strSql = "SELECT * FROM sectionB WHERE strEDIID = '" & strEDIID & "'"
		
	' student recordset
	set rstStudent = server.createobject("adodb.recordset")
		
	rststudent.Open strSql, conn_data
	if not rstStudent.eof then 	
									
		' get all the demographic questions
		strSql = "SELECT B.English, B.French, B.isSectionHeader, B.Section, B.Question, B.Option, B.isHeader,  C.Language, C.intOptions, C.Col1, C.Col2, C.Col3, C.Col4, C.Col5, C.Col6 FROM Page_SectionB B  LEFT JOIN [Column Headers] C ON B.HeaderType = C.HID WHERE C.Language='" & strLanguage & "' OR C.Language Is Null ORDER BY B.isSectionHeader, B.Section, B.Question, B.Option, B.isHeader"
									
		'open the Section A questions 
		rstData.Open strSql, conn_Tables,1  
									
		' header
		Response.Write "<br />&nbsp;<input type=""button"" onClick=""javascript:goWindow('" & strLink & "','Guide','490','500','top=0,left=125,resizable=yes');"" name=""btnGuide"" value=""GUIDE"" title=""GUIDE"">&nbsp;&nbsp;<a href=""javascript:goSaveEDI('" & left(strEDIID,3) & "','" & mid(strEDIID,4,3) & "','" & mid(strEDIID,7,2) & "','" & mid(strEDIID,9,1) & "','" & right(strEdiID,2) & "','_Demographics','B');"" class=""bigLinkRed"">"
		if strLanguage="English" Then 
			response.write "Demographics"
		else
			response.write "Démographique"
		end if 
		response.write "</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & left(strEDIID,3) & "','" & mid(strEDIID,4,3) & "','" & mid(strEDIID,7,2) & "','" & mid(strEDIID,9,1) & "','" & right(strEDIID,2)& "','A','B')"" class=""bigLinkRed"">Section A</a>&nbsp;&nbsp; <font class=""subheaderBlue"">Section B</font>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & left(strEDIID,3) & "','" & mid(strEDIID,4,3) & "','" & mid(strEDIID,7,2) & "','" & mid(strEDIID,9,1) & "','" & right(strEDIID,2)& "','C','B')"" class=""bigLinkRed"">Section C</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & left(strEDIID,3) & "','" & mid(strEDIID,4,3) & "','" & mid(strEDIID,7,2) & "','" & mid(strEDIID,9,1) & "','" & right(strEDIID,2)& "','D','B')"" class=""bigLinkRed"">Section D</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & left(strEDIID,3) & "','" & mid(strEDIID,4,3) & "','" & mid(strEDIID,7,2) & "','" & mid(strEDIID,9,1) & "','" & right(strEDIID,2)& "','E','B')"" class=""bigLinkRed"">Section E</a>"
		if not blnLock then
			response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type=""button"" onClick=""javascript:goSaveEDI('" & left(strEDIID,3) & "','" & mid(strEDIID,4,3) & "','" & mid(strEDIID,7,2) & "','" & mid(strEDIID,9,1) & "','" & right(strEDIID,2)& "','B','B')"" name=""btnSave"" value=""" & lblSaveEDI & """ title=""" & lblSaveEDI & """>"
		end if 
		response.write "<br />"									
		
		' Menu
		Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""750"" align=""center"">"																					
		Response.Write "<tr><td colspan=""3""><br /></td></tr>"
		Response.write "<tr><td width=""13%"" align=""right"">&nbsp;" & strPrevious & "</td><td width=""74%"" align=""center"" valign=""middle""><font class=""headerBlue"">" &  rstData(strLanguage) & "<font></td><td width=""13%"" align=""left"">&nbsp;" & strNext & "</td></tr>"
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
				next 
				Response.Write "</tr>"
											
				' set the inner table to true
				blnTable = true 
			else
				' if it is just a question and no header then write a new table
				if intOptions = 0 then 
					Response.Write "<tr><td>"	
					Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""750"" align=""center"">"																					
						Response.Write "<tr><td align=""left"" valign=""top""><font class=""boldTextBlack"">&nbsp;&nbsp;" 
						' only write the questions
						if rstData("option") = 0 then 
							if rstData("question") < 10 then 
								Response.Write "&nbsp;&nbsp;"
							end if 
							Response.Write rstData("question") & " ) &nbsp;&nbsp;</td><td valign=""middle""><font class=""boldTextBlack"">" &  rstData(strLanguage) & "</font>"
						end if
																	
						'select case rstData("question")
														
						'end select
						Response.Write "</td>"
						Response.Write "</tr>"
					Response.Write "</table>"
					Response.Write "</td></tr>"
					Response.Write "<tr><td><br /></td></tr>"
				else
					intRow = intRow + 1
					if intRow mod 2 = 1 then 
						strColour = "whitesmoke"
					else
						strColour = "white"
					end if 
					Response.Write "<tr bgcolor=""" & strColour & """><td align=""left"" valign=""top"">&nbsp;<font class=""boldTextBlack"">" 
					' only write the questions
					if rstData("option") = 0 then 
						if rstData("question") < 10 then 
							Response.Write "&nbsp;&nbsp;"
						end if 
						Response.Write rstData("question") & " ) &nbsp;&nbsp;</font></td><td align=""left"" valign=""middle""><font class=""boldTextBlack"">" &  rstData(strLanguage) & "</font>"
						strQuestion = "intQ" & rstData("Question")
					end if
																	
					select case rstData("question")
						case 40 
							if strLanguage = "English" Then
								response.write "<p align=""left""><font class=""boldTextBlack"">If yes, please specify:</font>"			
							else
								response.write "<p align=""left""><font class=""boldTextBlack"">Dans l’affirmative, précisez:</font>"			
							end if 							
							response.write "<input type=""text"" maxlength=""50"" name=""strOther"" size=""45"" value=""" & rstStudent("strOther") & """></p>"
					end select
					
					Response.Write "</td>"
												
					' write the radio option buttons
					for intCol = 1 to intoptions
						Response.Write "<td align=""center"" valign=""middle"" width=""70""><input type=""radio"" name=""" & strQuestion & """ value=""" & intCol & """"
						if rstStudent(strQuestion) = intCol then 
							response.write " checked"
						end if 
						response.write "></td>"
					next 
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
		Response.Write "<br />&nbsp;<input type=""button"" onClick=""javascript:goWindow('" & strLink & "','Guide','490','500','top=0,left=125,resizable=yes');"" name=""btnGuide"" value=""GUIDE"" title=""GUIDE"">&nbsp;&nbsp;<a href=""javascript:goSaveEDI('" & left(strEDIID,3) & "','" & mid(strEDIID,4,3) & "','" & mid(strEDIID,7,2) & "','" & mid(strEDIID,9,1) & "','" & right(strEdiID,2) & "','_Demographics','B');"" class=""bigLinkRed"">"
		if strLanguage="English" Then 
			response.write "Demographics"
		else
			response.write "Démographique"
		end if 
		response.write "</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & left(strEDIID,3) & "','" & mid(strEDIID,4,3) & "','" & mid(strEDIID,7,2) & "','" & mid(strEDIID,9,1) & "','" & right(strEDIID,2)& "','A','B')"" class=""bigLinkRed"">Section A</a>&nbsp;&nbsp; <font class=""subheaderBlue"">Section B</font>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & left(strEDIID,3) & "','" & mid(strEDIID,4,3) & "','" & mid(strEDIID,7,2) & "','" & mid(strEDIID,9,1) & "','" & right(strEDIID,2)& "','C','B')"" class=""bigLinkRed"">Section C</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & left(strEDIID,3) & "','" & mid(strEDIID,4,3) & "','" & mid(strEDIID,7,2) & "','" & mid(strEDIID,9,1) & "','" & right(strEDIID,2)& "','D','B')"" class=""bigLinkRed"">Section D</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & left(strEDIID,3) & "','" & mid(strEDIID,4,3) & "','" & mid(strEDIID,7,2) & "','" & mid(strEDIID,9,1) & "','" & right(strEDIID,2)& "','E','B')"" class=""bigLinkRed"">Section E</a>"
		if not blnLock then
			response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type=""button"" onClick=""javascript:goSaveEDI('" & left(strEDIID,3) & "','" & mid(strEDIID,4,3) & "','" & mid(strEDIID,7,2) & "','" & mid(strEDIID,9,1) & "','" & right(strEDIID,2)& "','B','B')"" name=""btnSave"" value=""" & lblSaveEDI & """ title=""" & lblSaveEDI & """>"
		end if 
		response.write "<br />"																		
	else
		strError = "<font class=""regtextred"">No data on child - " & strEDIID & "</font>"
	end if
	
	' close and kill the additional questions recordset
	call close_adodb(rstStudent)
	%>