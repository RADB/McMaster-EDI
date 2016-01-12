	<input type="hidden" name="CurrentSection" value="D" />
	<%
  select case session("province")
  	  ' Ontario
	    Case 1 					    
	        strWhere = " AND Ontario = 1 "
	    ' Alberta
	    case 3 
	        strWhere = " AND Alberta = 1 "
	    ' Manitoba
	    case 2 
	        strWhere= " AND Manitoba = 1 "
	    ' Saskatchewan
	    Case 5 
	        strWhere = " AND Saskatchewan = 1 "
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
	strSql = "SELECT * FROM sectionD WHERE strEDIID = '" & strEDIID & "'"
		
	' student recordset
	set rstStudent = server.createobject("adodb.recordset")
		
	rststudent.Open strSql, conn
	if not rstStudent.eof then 	
									
		' get all the demographic questions
    'strSql = "SELECT D.[English], D.[French], D.[isSectionHeader], D.[Section], D.[Question], D.[Option], D.[isHeader],  C.[Language], C.[intOptions], C.Col1, C.Col2, C.Col3, C.Col4, C.Col5, C.Col6 FROM Page_SectionD D  LEFT JOIN [Column Headers] C ON D.[HeaderType] = C.[HID] WHERE C.[Language]='" & strLanguage & "' OR C.[Language] Is Null ORDER BY D.[isSectionHeader], D.[Section], D.[Question], D.[Option], D.[isHeader]"
     strSql = "SELECT D.[English], D.[French], D.[isSectionHeader], D.[Section], D.[Question], D.[Option], D.[isHeader],  C.[Language], C.[intOptions], C.Col1, C.Col2, C.Col3, C.Col4, C.Col5, C.Col6, c.isCheckBox FROM Page_SectionD D  LEFT JOIN [Column Headers] C ON D.[HeaderType] = C.[HID] WHERE (C.[Language]='" & strLanguage & "' OR C.[Language] Is Null) " & strWhere & " ORDER BY D.Sequence"
				
    'response.write strSql
		'open the Section D questions 
		rstData.Open strSql, conn,1  
									
		' header
		Response.Write "<br />&nbsp;<input type=""button"" onclick=""javascript:goWindow('" & strLink & "','Guide','490','500','top=0,left=125,resizable=yes');"" name=""btnGuide"" value=""GUIDE"" title=""GUIDE"" />&nbsp;&nbsp;<a href=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','_Demographics','D');"" class=""bigLinkRed"">"
		if strLanguage="English" Then 
			response.write "Demographics"
		else
			response.write "Démographique"
		end if 
		response.write "</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','A','D')"" class=""bigLinkRed"">Section A</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','B','D')"" class=""bigLinkRed"">Section B</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','C','D')"" class=""bigLinkRed"">Section C</a>&nbsp;&nbsp; <font class=""subheaderBlue"">Section D</font>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','E','D')"" class=""bigLinkRed"">Section E</a>"
		if not blnLock then 
			response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type=""button"" onclick=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','D','D')"" name=""btnSave"" value=""" & lblSaveEDI & """ title=""" & lblSaveEDI & """ />"
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
          Response.Write "</td></tr>"
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
        isCheckBox = rstData("isCheckBox")
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
					Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""750"" align=""center"">"																					
						Response.Write "<tr><td align=""left"" valign=""middle""><font class=""boldTextBlack"">&nbsp;&nbsp;" 
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
							Response.Write "&nbsp;&nbsp;"
						end if 
						Response.Write rstData("question") & " &nbsp;&nbsp;</font></td><td align=""left""><font class=""boldTextBlack"">" &  rstData(strLanguage) & "</font>"
						strQuestion = "intQ" & rstData("question")
					elseif rstData("option") >= 97 then 
						if rstData("question") < 10 then 
							Response.Write "&nbsp;&nbsp;"
						end if 
            if rstData("option") = 97 then                                       
              ' colspan causing spacing issue - added to add width to second column                                         
						  Response.Write rstData("question") & chr(rstData("option")) & " &nbsp;&nbsp;</font></td><td align=""left"" style=""width:560px;""><font class=""boldTextBlack"">" &  rstData(strLanguage) & "</font>"
            else
              Response.Write rstData("question") & chr(rstData("option")) & " &nbsp;&nbsp;</font></td><td align=""left""><font class=""boldTextBlack"">" &  rstData(strLanguage) & "</font>"
            end if 
						strQuestion = "intQ" & rstData("Question") & chr(rstData("option"))
					end if
							
																	
					select case rstData("question")
						case 2
							if rstData("option")= 107 then 
								if strLanguage = "English" Then
									response.write "<p align=""left""><font class=""boldTextBlack"">If yes, please specify:</font>"			
								else
									response.write "<p align=""left""><font class=""boldTextBlack"">Si oui, veuillez préciser:</font>"			
								end if 							
							
								response.write "<input type=""text"" maxlength=""50"" name=""strProblem"" size=""45"" value=""" & rstStudent("strProblem") & """ /></p>"	
							end if 
 
            case 3
                dim aSpecial, rstSpecial               
				'if session("province")<>6 then 
					'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
					' get the special Problems for the drop down box
					'//////////////////////////////////////////////////////////////////////	
					set rstSpecial = server.CreateObject("adodb.recordset")

					' open all languages
					rstSpecial.Open "SELECT SPID, english, french FROM [LU SpecialProblems] ORDER BY sequence", conn  ' & strLanguage, conn

					' store all languages in array
					aSpecial = rstSpecial.GetRows 	

					' close and kill the langauges recordset
					call close_adodb(rstSpecial)
	            
	        
					  Response.Write "<br /><select name=""intQ3"">"
					  Response.Write "<option value=""0""></option>"										
					  for intRow = 0 to ubound(aSpecial,2)															
						  ' show the language
						  Response.write "<option value=""" & aSpecial(0,introw) & """"
						  if rstStudent(strQuestion) = aSpecial(0,intRow) then 
							  Response.Write " selected=""selected"""                                        
						  end if
						  Response.Write ">" & aSpecial(intLanguage,introw) & "</option>"
					  next
					  Response.Write "</select>"
					  
					  strQuestion = "intQ" & rstData("question") & "a"
					  Response.Write "<br /><select name=""intQ3a"">"
					  Response.Write "<option value=""0""></option>"										
					  for intRow = 0 to ubound(aSpecial,2)															
						  ' show the language
						  Response.write "<option value=""" & aSpecial(0,introw) & """"
						  if rstStudent(strQuestion) = aSpecial(0,intRow) then 
							  Response.Write " selected=""selected"""                                        
						  end if
						  Response.Write ">" & aSpecial(intLanguage,introw) & "</option>"
					  next
					  Response.Write "</select>"
					  
					  strQuestion = "intQ" & rstData("question") & "b"
					  Response.Write "<br /><select name=""intQ3b"">"
					  Response.Write "<option value=""0""></option>"										
					  for intRow = 0 to ubound(aSpecial,2)															
						  ' show the language
						  Response.write "<option value=""" & aSpecial(0,introw) & """"
						  if rstStudent(strQuestion) = aSpecial(0,intRow) then 
							  Response.Write " selected=""selected"""                                        
						  end if
						  Response.Write ">" & aSpecial(intLanguage,introw) & "</option>"
					  next
					  Response.Write "</select>"
					  
					if strLanguage = "English" Then
						response.write "<p align=""left""><font class=""boldTextBlack"">If other, please specify:</font>"			
					else
						response.write "<p align=""left""><font class=""boldTextBlack"">Si autre, s'il vous plaît précisez:</font>"			
					end if         
				'end if 								
                
                response.write "<input type=""text"" maxlength=""50"" name=""strOther"" size=""45"" value=""" & rstStudent("strOther") & """ /></p>"	
						case 5
              if rstData("option") = 99 then 
							  if strLanguage = "English" Then
								  response.write "<p align=""left""><font class=""boldTextBlack"">If yes, please specify:</font>"			
							  else
								  response.write "<p align=""left""><font class=""boldTextBlack"">Si oui, veuillez précisez:</font>"			
							  end if 	
  							
							  response.write "<input type=""text"" maxlength=""50"" name=""strAssess"" size=""45"" value=""" & rstStudent("strAssess") & """ /></p>"						
              end if 
					end select
					
					Response.Write "</td>"
												
					' write the radio option buttons or checkbox buttons
          if intOptions>0 then 
            if isCheckBox  then
				response.write "<td>"
				response.write buildCheckBox("intQ" & rstData("question")&  chr(rstData("option")),rstStudent(strQuestion))
				response.write "</td><td>"
				response.write buildCheckBox("intQ" & rstData("question")&  chr(rstData("option")) & "2" ,rstStudent(strQuestion & "2"))
				response.write "</td>"
				
				' Saskatchewan and Alberta
				if intOptions = 3 then
					Select case session("Province")
					' Alberta
					case 3
						'response.write "<td>"
						'response.write buildCheckBox("intQ" & rstData("question")&  chr(rstData("option")) & "Code" ,rstStudent(strQuestion & "Code"))
						'response.write "</td>"
					  ' saskatchewan
					  case 5
						response.write "<td>"
						response.write buildCheckBox("intQ" & rstData("question")&  chr(rstData("option")) & "RIS" ,rstStudent(strQuestion & "RIS"))
						response.write "</td>"                    
					end select
				end if 
			else              
				for intCol = 1 to intoptions
					Response.Write "<td align=""center"" valign=""middle"" width=""70"">"
					response.write buildRadioButton("horizontal", strQuestion, aHeader(intcol),intcol, rstStudent(strQuestion),"","")
					response.write "</td>"
				next 					  
            end if 
          end if
					response.write getHelpLink(session("Province"), "SectionD", strQuestion)
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
		
		'header
		Response.Write "<br />&nbsp;<input type=""button"" onclick=""javascript:goWindow('" & strLink & "','Guide','490','500','top=0,left=125,resizable=yes');"" name=""btnGuide"" value=""GUIDE"" title=""GUIDE"" />&nbsp;&nbsp;<a href=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','_Demographics','D');"" class=""bigLinkRed"">"
		if strLanguage="English" Then 
			response.write "Demographics"
		else
			response.write "Démographique"
		end if 
		response.write "</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','A','D')"" class=""bigLinkRed"">Section A</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','B','D')"" class=""bigLinkRed"">Section B</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','C','D')"" class=""bigLinkRed"">Section C</a>&nbsp;&nbsp; <font class=""subheaderBlue"">Section D</font>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','E','D')"" class=""bigLinkRed"">Section E</a>"
		if not blnLock then 
			response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type=""button"" onclick=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','D','D')"" name=""btnSave"" value=""" & lblSaveEDI & """ title=""" & lblSaveEDI & """ />"
		end if 		
		response.write "<br />"							
	else
		strError = "<font class=""regtextred"">No data on child - " & strEDIID & "</font>"
	end if
	' close and kill the additional questions recordset
	call close_adodb(rstStudent)

%>