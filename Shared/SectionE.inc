	<input type="hidden" name="CurrentSection" value="E" />
	<%
  select case session("province")
	    ' Alberta
	    case 3 
	        strWhere = " AND Alberta = 1 "
	    ' Manitoba
	    case 2 
	        strWhere= " AND Manitoba = 1 "
	    ' Saskatchewan
	    Case 5 
	        strWhere = " AND Saskatchewan = 1 "
	    ' Ontario
	    Case 1 					    
	        strWhere = " AND Ontario = 1 "
	    Case 6 
		strWhere = " AND NWT = 1 "
	    Case 7 
		strWhere = " AND Newfoundland = 1 "
	    Case 8 
		strWhere = " AND NovaScotia = 1 "
	    ' all others 
	    case else
	        strWhere = ""
	end select 
	
  ' get the demographic data
	strSql = "SELECT * FROM sectionE WHERE strEDIID = '" & strEDIID & "'"

  ' student recordset
  set rstStudent = server.createobject("adodb.recordset")

  rststudent.Open strSql, conn
  if not rstStudent.eof then
  '	aStudent = rstStudent.getrows
  else
  strError = "<font class=""regtextred"">No data on child - " & strEDIID & "</font>"
	end if 
    
	'if not rstStudent.eof then 								
		' get all the demographic questions
'		strSql = "SELECT E.[English], E.[French], E.[isSectionHeader], e.[section], E.[Question], E.[Option], E.isHeader,  C.[Language], C.intOptions, C.Col1, C.Col2, C.Col3, C.Col4, C.Col5, C.Col6 FROM Page_SectionE E  LEFT JOIN [Column Headers] C ON E.HeaderType = C.HID WHERE C.[Language]='" & strLanguage & "' OR C.[Language] Is Null ORDER BY E.[isSectionHeader], e.[section], E.[Question], E.[Option], E.[isHeader]"
    strSql = "SELECT E.[English], E.[French], E.[isSectionHeader], e.[section], E.[Question], E.[Option], E.isHeader,  C.[Language], C.intOptions, C.Col1, C.Col2, C.Col3, C.Col4, C.Col5, C.Col6 FROM Page_SectionE E  LEFT JOIN [Column Headers] C ON E.HeaderType = C.HID WHERE (C.[Language]='" & strLanguage & "' OR C.[Language] Is Null) " & strWhere & " ORDER BY E.Sequence"

  'open the Section D questions
  rstData.Open strSql, conn,1


  ' additional questions recordset
  set rstAdditional = server.createobject("adodb.recordset")
  ' get site specific questions
  strSql = "SELECT strQ6, strQ7, strQ8, strQ9, strQ10, strQ11, strQ6French, strQ7French, strQ8French, strQ9French, strQ10French, strQ11French FROM sites WHERE intSiteID = " & strSite

		rstAdditional.open strSql, conn
									
		' header
		Response.Write "<br />&nbsp;<input type=""button"" onclick=""javascript:goWindow('" & strLink & "','Guide','490','500','top=0,left=125,resizable=yes');"" name=""btnGuide"" value=""GUIDE"" title=""GUIDE"" />&nbsp;&nbsp;<a href=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','_Demographics','E');"" class=""bigLinkRed"">"
		if strLanguage="English" Then 
			response.write "Demographics"
		else
			response.write "Démographique"
		end if 
		response.write "</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','A','E')"" class=""bigLinkRed"">Section A</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','B','E')"" class=""bigLinkRed"">Section B</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','C','E')"" class=""bigLinkRed"">Section C</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','D','E')"" class=""bigLinkRed"">Section D</a>&nbsp;&nbsp; <font class=""subheaderBlue"">Section E</font>"
		if not blnLock then 
			response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type=""button"" onclick=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','E','E')"" name=""btnSave"" value=""" & lblSaveEDI & """ title=""" & lblSaveEDI & """ />"
		end if 
		response.write "<br />"
		
		' Menu
		Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""750"" align=""center"">"																					
		Response.Write "<tr><td colspan=""3""><br /></td></tr>"
		Response.write "<tr><td width=""23%"" align=""right"">&nbsp;" & strPrevious & "</td><td width=""54%"" align=""center"" valign=""middle""><font class=""headerBlue"">" &  rstData(strLanguage) & "</font></td><td width=""23%"" align=""left"">&nbsp;" & strNext & "</td></tr>"
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
					Response.Write "</td></tr>"
					Response.Write "<tr><td><br /></td></tr>"
				end if 
											
				' new section
				Response.Write "<tr><td align=""center"">"
											
				' get the number of options 
				intOptions = rstData("intOptions")
				if intOptions > 0 then 										
					' inner table for each section
					Response.Write "<table border=""1"""
				else
					Response.Write "<table border=""0"""
				end if 
				' header
				response.write " cellpadding=""0"" cellspacing=""0"" width=""750"" align=""center"">"																					
				Response.Write "<tr>"		
				Response.Write "<td align=""left"" colspan=""2""><font class=""subHeaderBlue"">" & rstData(strLanguage) & "</font></td>" 
				
				if intOptions > 0 then 																		
					for intCol = 1 to intoptions
						Response.Write "<td align=""center"" valign=""middle"" width=""70""><font class=""boldTextBlack"">" & rstData("Col" & intCol) & "</font></td>"
						' store the header values for the title below 2005 09 15
						aHeader(intCol) = rstData("Col" & intCol)
					next 
				end if 
				
				Response.Write "</tr>"
											
				' set the inner table to true
				blnTable = true 
			else
				' if it is just a question and no header then write a new table
				if intOptions = 0 then 
					Response.Write "<tr><td>"	
					Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""750"" align=""center"">"																					
						Response.Write "<tr><td align=""left"" valign=""top"" width=""50""><font class=""boldTextBlack"">&nbsp;&nbsp;" 
						' only write the questions
						if rstData("option") = 0 then 
							if rstData("question") < 10 then 
								Response.Write "&nbsp;&nbsp;"
							end if 
							Response.Write rstData("question") & " ) &nbsp;&nbsp;</td><td><font class=""boldTextBlack"">" &  rstData(strLanguage) & "</font>"
						end if
																	
					
					Response.Write "</td>"
					Response.Write "</tr>"
					Response.Write "</table>"
					Response.Write "</td></tr>"
					Response.Write "<tr><td><br /></td></tr>"
				else
					intRow = intRow + 1
					if intRow mod 2 = 1 then 
						'strColour = "whitesmoke"
						strColour = "#f5f5f5"
					else
						strColour = "white"
					end if 					
					
					Response.Write "<tr style=""background-color:" & strColour & ";""><td align=""left"" valign=""top"" style=""width:35px;"">&nbsp;<font class=""boldTextBlack"">" 
					' only write the questions
					if rstData("option") = 0 then 
						if rstData("question") < 10 then 
							if session("province") = 5 then 
								if (intQuestion+1) < 10 then 
								  Response.Write "&nbsp;&nbsp;"                 
								end if
								if rstData("question") > 5 then 
								  intQuestion2 = intQuestion  + 1
								  intQuestion = intQuestion + 1
								else
								  intQuestion2 = rstData("question") 
								end if 
							else              
								Response.Write "&nbsp;&nbsp;"
								intQuestion2 = rstData("question")                
							end if 
						else
							intQuestion2 = rstData("question")
						end if 
            
						Response.Write intQuestion2 & " &nbsp;&nbsp;</font></td><td align=""left"" ><font class=""boldTextBlack"">" &  rstData(strLanguage) & "</font>"
						strQuestion = "intQ" & rstData("Question")
					elseif rstData("option") >= 97 then             
						if rstData("question") < 10 then 
						  if session("province") = 5 then 
							if (intQuestion+1) < 10 then 
							  Response.Write "&nbsp;&nbsp;"                 
							end if
							if rstData("question") > 5 then 
							  intQuestion2 = intQuestion - rstData("question") + 7
							else
							  intQuestion2 = rstData("question") 
							end if 
						else              
							Response.Write "&nbsp;&nbsp;"
							intQuestion2 = rstData("question")
						end if 
					else
						intQuestion2 = rstData("question")
					end if 
            
					Response.Write intQuestion2 & chr(rstData("option")) & " &nbsp;&nbsp;</font></td><td align=""left"" style=""width:630px;""><font class=""boldTextBlack"">" &  rstData(strLanguage) & "</font>"
					if session("province") = 6 then
						strQuestion = "intQNWT" & rstData("Question") & chr(rstData("option"))
					else
						strQuestion = "intQ" & rstData("Question") & chr(rstData("option"))
					end if 
				end if
				
				if session("province") = 6 then 
					 if rstData("question") = 2 AND rstData("option") = 106 then
						response.write "<input type=""text"" maxlength=""50"" name=""strOtherServices"" size=""45"" value=""" & rstStudent("strOtherServices") & """ />"	
					 end if 
				else
					select case rstData("question")
						case 1,3
						
								response.write "<p align=""left""><font class=""boldTextBlack"">"& lblESpecifyPrint &"</font>"
							
							
							if rstData("question") = 1 then 
								response.write "<input type=""text"" maxlength=""50"" name=""strIntervention"" size=""35"" value=""" & rstStudent("strIntervention") & """ /></p>"
							elseif rstData("question") = 3 then 
								response.write "<input type=""text"" maxlength=""50"" name=""strSpecify"" size=""35"" value=""" & rstStudent("strSpecify") & """ /></p>"
							end if 
            
						case 5 
							response.write "</td>"
							' write the radio option buttons
							for intCol = 1 to intoptions
								'Response.Write "<td align=""center"" valign=""middle"" width=""70"">"
								'<input type=""radio"" title=""" & aHeader(intCol) &""" name=""" & strQuestion & """ value=""" & intCol & """"
								' check the selected box
								'if intCol = rstStudent(strQuestion) then 
								'	response.write " checked=""CHECKED"""
								'end if 
										
								'response.write " /></td>"
								Response.Write "<td align=""center"" valign=""middle"" width=""70"">"
								response.write buildRadioButton("horizontal", strQuestion, aHeader(intcol),intcol, rstStudent(strQuestion),"","")
								response.write "</td>"
							next 
							intQuestion = 11
						  'if session("Province") <> 5 then 
						  if rstData("option") = 0 then 

							  ' insert questions 6 - 11 
							    for intRow = 6 to 11
								    ' set the question string 
								    strQuestion = "intExtraQ" & intRow 
    								
								    if intRow mod 2 = 1 then 
									    strColour = "#f5f5f5"
								    else
									    strColour = "white"
								    end if 
    															
								    if rstAdditional("strQ" & intRow) <> "" then 
									    Response.Write "<tr bgcolor=""" & strColour & """><td align=""left"" valign=""top"" width=""40"">&nbsp;<font class=""boldTextBlack"">" 
									    ' only write the questions
									    if intRow < 10 then 
										    Response.Write "&nbsp;&nbsp;"
									    end if 
			                      
								  ' deal with french additional questions
								  if strLanguage = "English" Then
									strLang = ""
								  else
									strLang = "French"
								  end if 
			                      
													Response.Write intRow & " ) &nbsp;&nbsp;</font></td><td align=""left"" ><font class=""boldTextBlack"">" &  rstAdditional("strQ" & intRow & strLang) & "</font>"
													Response.Write "</td>"
			    									
													' write the radio option buttons
													for intCol = 1 to 3
														'Response.Write "<td align=""center"" valign=""middle"" width=""70""><input type=""radio"" title=""" & aHeader(intCol) &""" name=""" & strQuestion & """ value=""" & intCol & """"
			    										
														' check the selected box
														'if intCol = rstStudent(strQuestion) then 
														'	response.write " checked=""CHECKED"""
														'end if 
			    										
														'response.write " /></td>"
														Response.Write "<td align=""center"" valign=""middle"" width=""70"">"
														response.write buildRadioButton("horizontal", strQuestion, aHeader(intcol),intcol, rstStudent(strQuestion),"","")
														response.write "</td>"
													next 
								  response.write "</tr>"  
								else
								  intQuestion = intRow -1
								  Exit For
												end if 		
											next	
							'end if 
						  end if 
								case 6 
						  if rstData("Option") = 104 then   
							if strLanguage = "English" Then
											  response.write "<p align=""left""><font class=""boldTextBlack"">If yes, please specify:</font>"			
							else
											  response.write "<p align=""left""><font class=""boldTextBlack"">Si oui, veuillez préciser:</font>"			                        
										  end if 
			                            
							response.write "<input type=""text"" maxlength=""50"" name=""strOtherServices"" size=""45"" value=""" & rstStudent("strOtherServices") & """ /></p>"	
						  end if 
					end select
				' end if province = NWT
				end if 
					if rstData("question") <> 5 then 				
						Response.Write "</td>"
						
						if intOptions > 0 then 
							if intOptions = 1 then
								response.write "<td>"
															
								'<input type=""checkbox"" id=""intQ" & rstData("question")&  chr(rstData("option"))& """ name=""intQ" & rstData("question")&  chr(rstData("option"))&""""
								'if rstStudent(strQuestion) then 
								'  response.write " checked=""CHECKED"""
								'end if                                      
								'response.write " />"
								
								response.write buildCheckBox("intQ" & rstData("question")&  chr(rstData("option")), rstStudent(strQuestion))
								
								response.write "</td>"
							else
								' write the radio option buttons
								for intCol = 1 to intoptions
									Response.Write "<td align=""center"" valign=""middle"" width=""70"">"
									'response.write "<input type=""radio"" title=""" & aHeader(intCol) &""" name=""" & strQuestion & """ value=""" & intCol & """"
	    							
									' check the selected box
								'	if intCol = rstStudent(strQuestion) then 
								'		response.write " checked=""CHECKED"""
									'end if 
	    										
									'response.write " />
									response.write buildRadioButton("horizontal", strQuestion, aHeader(intcol),intcol, rstStudent(strQuestion),"","")
									response.write "</td>"
								next          
						  end if 
					end if 
			  	end if 
					response.write getHelpLink(session("Province"), "SectionE", strQuestion)				
					Response.Write "</tr>"
				end if 										
			end if 
      
			rstData.movenext
		loop
	   
		response.write "<tr><td colspan=""2"" align=""center""><textarea rows=""5"" cols=""90"" name=""strComments"">" & rstStudent("strComments") & "</textarea></td></tr>"
	
	
		' close and kill the additional questions recordset
		call close_adodb(rstAdditional)
	
		' close the last inner table if it is open
		if blnTable then 
			Response.Write "</table>"
		end if 
		if not blnLock then 
'			Response.Write "<a href=""javascript:goConfirm_Lock('" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','lock');"" class=""hugeLinkBlue""><img src=""images/lock.gif"" border=""0"" alt=""lock"" />" & lblCompletion & "</a>"
			Response.Write "<a href=""javascript:goConfirm_Lock('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','lock');"" class=""hugeLinkBlue""><img src=""images/lock.gif"" border=""0"" alt=""Lock"" />" & lblCompletion & "</a>"
		end if 
									
		' end the row
		Response.Write "</td></tr>"
		Response.Write "</table>"	


		
		' header
			Response.Write "<br />&nbsp;<input type=""button"" onclick=""javascript:goWindow('" & strLink & "','Guide','490','500','top=0,left=125,resizable=yes');"" name=""btnGuide"" value=""GUIDE"" title=""GUIDE"" />&nbsp;&nbsp;<a href=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','_Demographics','E');"" class=""bigLinkRed"">"
		if strLanguage="English" Then 
			response.write "Demographics"
		else
			response.write "Démographique"
		end if 
		response.write "</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','A','E')"" class=""bigLinkRed"">Section A</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','B','E')"" class=""bigLinkRed"">Section B</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','C','E')"" class=""bigLinkRed"">Section C</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','D','E')"" class=""bigLinkRed"">Section D</a>&nbsp;&nbsp; <font class=""subheaderBlue"">Section E</font>"
		if not blnLock then 
			response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type=""button"" onclick=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','E','E')"" name=""btnSave"" value=""" & lblSaveEDI & """ title=""" & lblSaveEDI & """ />"
		end if 
		response.write "<br />"		
	'else
	'	strError = "<font class=""regtextred"">No data on child - " & strEDIID & "</font>"
	'end if

	' close and kill the additional questions recordset
	call close_adodb(rstStudent)
	%>