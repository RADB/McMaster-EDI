	<input type="hidden" name="CurrentSection" value="_Demographics" />
	<%
   
  dim strWhere
    
	' get the demographic data
	strSql = "SELECT * FROM Student_Demographics WHERE strEDIID = '" & strEDIID & "' ORDER BY strEDIID"
	'response.write strSql	
	rstData.Open strSql, conn
	if not rstData.eof then 
		aStudent = rstData.getrows
	else
		strError = "<font class=""regtextred"">No data on child - " & strEDIID & "</font>"
	end if
		
	' close the recordset
	rstData.close

	' header							
	Response.Write "<br />&nbsp;<input type=""button"" onclick=""javascript:goWindow('" & strLink & "','Guide','490','500','top=0,left=125,resizable=yes');"" name=""btnGuide"" value=""GUIDE"" title=""GUIDE"" />&nbsp;&nbsp;<font class=""subheaderBlue"">"	
	if strLanguage="English" Then 
		response.write "Demographics"
	else
		response.write "Démographique"
	end if 
	response.write "</font>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','A','_Demographics');"" class=""bigLinkRed"">Section A</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','B','_Demographics')"" class=""bigLinkRed"">Section B</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','C','_Demographics')"" class=""bigLinkRed"">Section C</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','D','_Demographics')"" class=""bigLinkRed"">Section D</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','E','_Demographics')"" class=""bigLinkRed"">Section E</a>"
	if NOT blnLock then 
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type=""button"" onclick=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','_Demographics','_Demographics')"" name=""btnSave"" value=""" & lblSaveEDI & """ title=""" & lblSaveEDI & """ />"
	end if 
	response.write "<br />"

	select case session("province")

	' Alberta
	case 3
	strWhere = " WHERE Alberta = 1 "
	' Manitoba
	case 2
	strWhere= " WHERE Manitoba = 1 "
	' Saskatchewan
	Case 5
	strWhere = " WHERE Saskatchewan = 1 "
	' Ontario
	Case 1
	strWhere = " WHERE Ontario = 1 "
	Case 6
	strWhere = " WHERE NWT = 1 "
	Case 7
	strWhere = " WHERE Newfoundland = 1 "
	Case 8
	strWhere = " WHERE NovaScotia = 1 "
	' all others
	case else
	strWhere = ""
	end select


	' get all the demographic questions - changed May 24, 2004
	' strSql = "SELECT question, english, french,option FROM Page_Section_Demographics WHERE Question>0 ORDER BY Question, Option"
	'strSql = "SELECT question, english, french,option,Manitoba,Ontario,Alberta,Saskatchewan FROM Page_Section_Demographics ORDER BY Question, Option"
	' changed to allow province specific features
	strSql = "SELECT D.question, D.english, D.french,D.[option],C.[Language], C.[intOptions], C.Col1, C.Col2, C.Col3, C.Col4, C.Col5, C.Col6, c.isCheckBox FROM Page_Section_Demographics D LEFT JOIN [Column Headers] C ON D.[HeaderType] = C.[HID] " & strWhere & " AND (C.[Language]='" & strLanguage & "' OR C.[Language] Is Null) ORDER BY Sequence"

	'				response.write strsql
	'open the demographic questions
	rstData.Open strSql, conn,1

	' added May 24, 2004
	if not rstData.eof then
	Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""750"" align=""center"">"																					
		Response.Write "<tr><td colspan=""3""><br /></td></tr>"
		Response.write "<tr><td width=""24%"" align=""right"">&nbsp;" & strPrevious & "</td><td width=""55%"" align=""center"" valign=""middle""><font class=""headerBlue"">" &  rstData(strLanguage) & "</font></td><td width=""23%"" align=""left"">&nbsp;" & strNext & "</td></tr>"
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

    
      if session("province") = 3 then                        
        select case rstData("question") 
          case 7 
            strQuestion= "&nbsp;&nbsp;7" 
          case 8
            strQuestion  = "&nbsp;&nbsp;" & rstData("question") + 1  			
          case 9,10,11,12,13,14,15                              
           strQuestion  = rstData("question") + 1
          case 17
            strQuestion = "8a"                       
			strColour = "white"
          case 18            
            strQuestion = "8b"                     
			strColour = "whitesmoke"
          case else
            strQuestion = "&nbsp;&nbsp;" & rstData("question")                       
        end select
      ' if manitoba
      elseif session("province") = 2 then       
        Select Case rstData("Question") 
          case 16 
            strQuestion = "&nbsp;&nbsp;5"                     
          case 5, 6,7,8
            strQuestion =  "&nbsp;&nbsp;" & rstData("Question")+1
          case 9,10,11,12,13,14,15
            strQuestion = rstData("Question")+1
          case else
            strQuestion = "&nbsp;&nbsp;" & rstData("question")                       
        end select
      else
         strQuestion =  rstData("question") 
      end if   
               
		  Response.Write "<tr bgcolor=""" & strColour & """>"

		  ' only write the questions
		  if rstData("option") = 0 then 
			'2016-01-09 - changed again- this time removed date of completion for all provinces - need to remove question numbers.
			  '2015-12-10 - only write question numbers if not in alberta
				'if session("province") <>3 then				
				'	response.write "<td align=""left"" valign=""top"" width=""35"">&nbsp;<font class=""boldTextBlack"">" 
				'					
				'	if rstData("question") < 10 and Session("Province") > 3 and Session("Province") < 2 then 
				''		  Response.Write "&nbsp;&nbsp;"        
				'	end if 
				'		
				'	Response.Write strQuestion & " &nbsp;&nbsp;</font></td>"
				'end if 
			' added 2012-10-02
			'2015-12-10 - do not display for alberta
			' 2016-01-09 - see above - no numbers for any province
			'if session("province") =3 then 
			'	if (rstdata("question")< 6 or rstData("question") >8) AND rstData("question")<> 11 AND rstData("question")<17 then
			'		if session("province") = 3 and (rstData("question")<>7 and rstData("question")<>17 and rstData("question")<>18) then
			'				response.write "<td colspan=""2""><font class=""boldTextBlack"">" &  rstData(strLanguage) & ": </font>"
			'		else 
			'			response.write "<td><font class=""boldTextBlack"">" &  rstData(strLanguage) & ": </font>"
			'		end if 
			'	end if 			
			'else
				'if session("province") = 3 and (rstData("question")<>7 and rstData("question")<>17 and rstData("question")<>18) then
					response.write "<td colspan=""2""><font class=""boldTextBlack"">" &  rstData(strLanguage) & ": </font>"
				'else 
				'	response.write "<td><font class=""boldTextBlack"">" &  rstData(strLanguage) & ": </font>"
				'end if 
			'end if
		end if

		' NWT Demographics
		if session("province") = 6 then		
		select case rstData("question")			
			' Class Assignment
			case 1
				' check value - if null default to kindergarten 2015-01-13
				if isnull(aStudent(16,0)) then 
					aStudent(16,0) = 2
				end if 
				
				rstData.MoveNext 																		
				Response.Write "<select name=""intClassAssignment"">"
				do while rstData("question") = 1	
					Response.write "<option value=""" & rstData("option") & """" 
						if aStudent(16,0) = rstData("option") then 
							Response.write " selected=""selected"""
						end if 
						response.write ">" & rstData(strLanguage) & "</option>"  
					rstData.MoveNext
				loop
				Response.Write "</select>"
													
				rstData.MovePrevious 
			' dob
			  case 2
				rstData.MoveNext
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
						Response.write " selected=""selected"""
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
						Response.write " selected=""selected"""
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
				for introw = 1 to (year(date)-1997)
					Response.Write "<option value = """ & intRow + 1994 & """"
					if intYear = intRow + 1994 then 
						Response.write " selected=""selected"""
					end if 
					' write the day
					Response.Write ">" & intRow + 1994 & "</option>"
				next
				%>
				</select>
				<%  
				rstData.MovePrevious 
			' sex
			case 3							
				rstData.MoveNext 
													
				Response.Write "<select name=""intSex"">"
				Response.Write "<option value=""-1""></option>"										
				do while rstData("question") = 3															
					' show the language
					Response.write "<option value=""" & rstData("option") & """"
					if aStudent(4,0) = rstData("option") then 
						Response.Write " selected=""selected"""
					end if 
					Response.Write ">" & rstData(strLanguage) & "</option>"
					rstData.MoveNext 
				loop
				Response.Write "</select>"
													
				rstData.MovePrevious 
			' postal code
			case 4 
				Response.Write "<input type=""text"" size=""10"" name=""strPostal"" value=""" & aStudent(6,0) & """ maxlength=""7"" />" 
			' class type
			case 5
					rstData.MoveNext 
				
				Response.Write "<select name=""intClassType"">"
				Response.Write "<option value=""-1""></option>"										
				do while rstData("question") = 5	          
					Response.write "<option value=""" & rstData("option") & """"
					      if aStudent(17,0) = rstData("option") then 
						      Response.Write " selected=""selected"""
					      end if 
					      Response.Write ">" & rstData(strLanguage) & "</option>"
					
					
					rstData.MoveNext 
				loop
				Response.Write "</select>"
													
				rstData.MovePrevious
			' date of completion 
			case 6 
				Response.Write "<font class=""lrgRegTextBlack"">"
				if not isnull(aStudent(8,0)) then 
					if strLanguage = "English" then 
						Response.Write right("00" & day(aStudent(8,0)),2) & "-" & monthname(datepart("m",aStudent(8,0)),true) & "-" & year(aStudent(8,0)) 		
					else
						Response.Write right("00" & day(aStudent(8,0)),2) & "-" & left(French_Month(datepart("m",aStudent(8,0))),3) & "-" & year(aStudent(8,0)) 		
					end if 						
				else
					if strLanguage = "English" then	
						Response.write "Incomplete"
					else
						response.write "Incomplète"
					end if
				end if 
				Response.Write "</font>"
			' special needs
			case 7
				rstData.MoveNext 
													
				Response.Write "<select name=""intSpecial"">"
				Response.Write "<option value=""-1""></option>"										
				do while rstData("question") = 7															
					' show the language
					Response.write "<option value=""" & rstData("option") & """"
					if aStudent(20,0) = rstData("option") then 
						Response.Write " selected=""selected"""
					end if 
					Response.Write ">" & rstData(strLanguage) & "</option>"
					rstData.MoveNext 
				loop
				Response.Write "</select>"
													
				rstData.MovePrevious
			' ESL
			case 8 
				rstData.MoveNext 
													
				Response.Write "<select name=""intESL"">"
				Response.Write "<option value=""-1""></option>"										
				do while rstData("question") = 8															
					' show the language
					Response.write "<option value=""" & rstData("option") & """"
					if aStudent(19,0) = rstData("option") then 
						Response.Write " selected=""selected"""
					end if 
					Response.Write ">" & rstData(strLanguage) & "</option>"
					rstData.MoveNext 
				loop
				Response.Write "</select>"
													
				rstData.MovePrevious 
			' french immersion
			case 9
				rstData.MoveNext 
													
				Response.Write "<select name=""intFrench"">"
				Response.Write "<option value=""-1""></option>"										
				do while rstData("question") = 9															
					' show the language
					Response.write "<option value=""" & rstData("option") & """"
					if aStudent(22,0) = rstData("option") then 
						Response.Write " selected=""selected"""
					end if 
					Response.Write ">" & rstData(strLanguage) & "</option>"
					rstData.MoveNext 
				loop
				Response.Write "</select>"
													
				rstData.MovePrevious 
			' other immersion
			case 10
				rstData.MoveNext 
													
				Response.Write "<select name=""intOther"">"
				Response.Write "<option value=""-1""></option>"										
				do while rstData("question") = 10
						' show the language
						Response.write "<option value=""" & rstData("option") & """"
						if aStudent(23,0) = rstData("option") then 
							Response.Write " selected=""selected"""
						end if 
						Response.Write ">" & rstData(strLanguage) & "</option>"
					rstData.MoveNext 
				loop
				Response.Write "</select>"
				rstData.MovePrevious 			
			'Ethnic Status
			case 11
				rstData.MoveNext		
				Response.Write "<select name=""intEthnicStatus"">"
				Response.Write "<option value=""-1""></option>"			
				do while rstData("question") = 11	
					  Response.write "<option value=""" & rstData("option") & """"
					  if aStudent(30,0) = rstData("option") then 
						  Response.Write " selected=""selected"""
					  end if 
					  Response.Write ">" & rstData(strLanguage) & "</option>"					
				  rstData.MoveNext                
				loop
				rstData.MovePrevious 
			' communicates adequately
			case 12
				rstData.MoveNext								
				Response.Write "<select name=""intCommunicates"">"
				Response.Write "<option value=""-1""></option>"										
				do while rstData("question") = 12
					' show the language
					Response.write "<option value=""" & rstData("option") & """"
					if aStudent(25,0) = rstData("option") then 
						Response.Write " selected=""selected"""
					end if 
					Response.Write ">" & rstData(strLanguage) & "</option>"
					rstData.MoveNext 
				loop
				Response.Write "</select>"

				rstData.MovePrevious  
			' language
			case 13
			' status
				rstData.MoveNext
				response.write "</td></tr>"
				do while rstData("question") = 13
					response.write "<tr><td>&nbsp;</td><td align=""left"">"
					
					if rstData("Option") < 42 then						
						' buildcheckbox
						response.write buildCheckBox("intNWT" & replace(replace(replace(rstData("English"),"'","")," ",""),"(Dogrib)",""), aStudent(rstData("Option"),0))		
					end if 
					
					Response.Write "<font class=""boldTextBlack"">" & rstData(strLanguage) & "</font>" 
					
					if rstData("Option") = 42 then 
						Response.Write "<select name=""intNWTOther"">"
						Response.Write "<option value=""-1""></option>"										
						for intRow = 0 to ubound(aLanguages,2)															
							' show the language
							Response.write "<option value=""" & aLanguages(0,introw) & """"
							if aStudent(42,0) = aLanguages(0,intRow) then 
								Response.Write " selected=""selected"""
							end if 
							Response.Write ">" & aLanguages(intLanguage,introw) & "</option>"
						next
						Response.Write "</select>"
					end if 
					rstData.movenext
				loop
				

				rstData.MovePrevious  
			case 14
				rstData.MoveNext 
				'2012-11-26 - removed popup from NWT
				if session("province") = 6 then 
					Response.Write "<select name=""intStatus"">"									
				elseif session("province") = 3 then 
					Response.Write "<select name=""intStatus"" onchange=""javascript:checkStatus(this.selectedIndex,'" & strAlbertaWarning & "', '" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "');"">"									
				else
					Response.Write "<select name=""intStatus"" onchange=""javascript:checkStatus(this.selectedIndex,'" & strWarning & "', '" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "');"">"									
				end if 
				
				Response.Write "<option value=""-1""></option>"										
				do while rstData("question") = 14															
					' show the language
					Response.write "<option value=""" & rstData("option") & """"
					if aStudent(18,0) = rstData("option") then 
						Response.Write " selected=""selected"""
					end if 
					Response.Write ">" & rstData(strLanguage) & "</option>"
					rstData.MoveNext 
				loop
				Response.Write "</select>"
													
				rstData.MovePrevious
			' repeat grade
			case 15
				rstData.MoveNext		
				Response.Write "<select name=""intRepeat"">"
				Response.Write "<option value=""-1""></option>"			
				do while rstData("question") = 15	
					  Response.write "<option value=""" & rstData("option") & """"
					  if aStudent(24,0) = rstData("option") then 
						  Response.Write " selected=""selected"""
					  end if 
					  Response.Write ">" & rstData(strLanguage) & "</option>"					
				  rstData.MoveNext                
				loop
				rstData.MovePrevious 
			' beneficiary
			case 16
				rstData.MoveNext		
				Response.Write "<select name=""intBeneficiary"">"
				Response.Write "<option value=""-1""></option>"			
				do while not rstData.eof	
					  Response.write "<option value=""" & rstData("option") & """"
					  if aStudent(29,0) = rstData("option") then 
						  Response.Write " selected=""selected"""
					  end if 
					  Response.Write ">" & rstData(strLanguage) & "</option>"					
				  rstData.MoveNext                
				loop
				rstData.MovePrevious
		end select
		
	  else
	  
	  
		select case rstData("question") 
      										
			' Class Assignment
			case 1
				rstData.MoveNext 
					'********************************************
				' 2008-11-14
				' added provincial specs
				'********************************************																				
				Response.Write "<select name=""intClassAssignment"">"
        ' only include null option for BC and Ontario
        if session("province") = 1 or session("province") = 4 then 
				  Response.Write "<option value=""-1""></option>"										
        end if 
				do while rstData("question") = 1															
					select case session("province")
               ' Ontario
               'Case 1 					    
                 ' only include SK
		 ' change 2011-04-28 to include JK
                 'if rstData("option") = 2 then 
		 'Response.write "<option value=""" & rstData("option") & """ selected=""selected"">" & rstData(strLanguage) & "</option>"        
                 'end if 					   

					    ' Manitoba, Alberta, Saskatchewan
					   case 2,3,5
					           ' only include K
                 if rstData("option") = 3 then 
					          Response.write "<option value=""" & rstData("option") & """ selected=""selected"">" & rstData(strLanguage) & "</option>"        
                 end if 		
			' Ontario			   					  					        					    					  
					    case else
					        Response.write "<option value=""" & rstData("option") & """"
					          if aStudent(16,0) = rstData("option") then 
						          Response.Write " selected=""selected"""
					          end if 
					          Response.Write ">" & rstData(strLanguage) & "</option>"        
					end select 
					
					rstData.MoveNext 
				loop
				Response.Write "</select>"
													
				rstData.MovePrevious 
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
						Response.write " selected=""selected"""
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
						Response.write " selected=""selected"""
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
				for introw = 1 to (year(date)-1997)
					Response.Write "<option value = """ & intRow + 1994 & """"
					if intYear = intRow + 1994 then 
						Response.write " selected=""selected"""
					end if 
					' write the day
					Response.Write ">" & intRow + 1994 & "</option>"
				next
				%>
				</select>
				<%                    
      ' Class Type
			case 5
				rstData.MoveNext 
				
				Response.Write "<select name=""intClassType"">"
				Response.Write "<option value=""-1""></option>"										
				do while rstData("question") = 5	          
		select case session("province")
               ' Ontario
               Case 1 					    
                 ' only include Options 1 - 5 and 12
                 if (rstData("option") >= 1 AND rstData("option") <= 5) OR rstData("option") = 12 then 
					           Response.write "<option value=""" & rstData("option") & """"
					            if aStudent(17,0) = rstData("option") then 
						            Response.Write " selected=""selected"""
					            end if 
					            Response.Write ">" & rstData(strLanguage) & "</option>"      
                 end if 					   
					    ' Manitoba
					    case 2
					           ' only include Options 7, 9 and 12
                 if rstData("option") = 7 OR rstData("option") = 9 OR rstData("option") = 12 then 
					           Response.write "<option value=""" & rstData("option") & """"
					            if aStudent(17,0) = rstData("option") then 
						            Response.Write " selected=""selected"""
					            end if 
					            Response.Write ">" & rstData(strLanguage) & "</option>"      
                 end if 
              'Alberta
              Case 3
              	 ' only include Options 7-9,12
                 if (rstData("option") >= 7 AND rstData("option") <= 9) OR rstData("option") = 12 then 
					           Response.write "<option value=""" & rstData("option") & """"
					            if aStudent(17,0) = rstData("option") then 
						            Response.Write " selected=""selected"""
					            end if 
					            Response.Write ">" & rstData(strLanguage) & "</option>"      
                 end if 
              'Saskatchewan
              Case 5
              ' only include Options 7, 9-12
                 if (rstData("option") >= 9 AND rstData("option") <= 12) OR rstData("option") = 7 then 
					           Response.write "<option value=""" & rstData("option") & """"
					            if aStudent(17,0) = rstData("option") then 
						            Response.Write " selected=""selected"""
					            end if 
					            Response.Write ">" & rstData(strLanguage) & "</option>"      
                 end if 
					    case else
					      Response.write "<option value=""" & rstData("option") & """"
					      if aStudent(17,0) = rstData("option") then 
						      Response.Write " selected=""selected"""
					      end if 
					      Response.Write ">" & rstData(strLanguage) & "</option>"
					end select 
					
					rstData.MoveNext 
				loop
				Response.Write "</select>"
													
				rstData.MovePrevious 
            
			' Date of completion
			case 6
				'2015-12-10 - do not display for alberta
				if session("province") <>3 then 
					Response.Write "<font class=""lrgRegTextBlack"">"
					if not isnull(aStudent(8,0)) then 
						if strLanguage = "English" then 
							Response.Write right("00" & day(aStudent(8,0)),2) & "-" & monthname(datepart("m",aStudent(8,0)),true) & "-" & year(aStudent(8,0)) 		
						else
							Response.Write right("00" & day(aStudent(8,0)),2) & "-" & left(French_Month(datepart("m",aStudent(8,0))),3) & "-" & year(aStudent(8,0)) 		
						end if 						
					else
						if strLanguage = "English" then	
							Response.write "Incomplete"
						else
							response.write "Incomplète"
						end if
					end if 
					Response.Write "</font>"
				end if 
			' DOB
			
			' Language Section
			case 12
				Response.Write "<select name=""intLanguage1"">"
				Response.Write "<option value=""-1""></option>"										
				for intRow = 0 to ubound(aLanguages,2)															
					' show the language
					Response.write "<option value=""" & aLanguages(0,introw) & """"
					if aStudent(14,0) = aLanguages(0,intRow) then 
						Response.Write " selected=""selected"""
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
						Response.Write " selected=""selected"""
					end if 
					Response.Write ">" & aLanguages(intLanguage,introw) & "</option>"
				next
				Response.Write "</select>"	
				
				' allow 3rd language for all provinces except NWT
				if session("province")<>6 then 
					Response.Write "<select name=""intLanguage3"">"
					Response.Write "<option value=""-1""></option>"										
					for intRow = 0 to ubound(aLanguages,2)															
						' show the language
						Response.write "<option value=""" & aLanguages(0,introw) & """"
						if aStudent(43,0) = aLanguages(0,intRow) then 
							Response.Write " selected=""selected"""
						end if 
						Response.Write ">" & aLanguages(intLanguage,introw) & "</option>"
					next
					Response.Write "</select>"		
				end if
			
			' Student Status
			case 14
				rstData.MoveNext 
				'********************************************
				' 2004 11 24
				' added javascript for Check Status			
				'********************************************						
				if session("province") = 3 then 
					Response.Write "<select name=""intStatus"" onchange=""javascript:checkStatus(this.selectedIndex,'" & strAlbertaWarning & "', '" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "');"">"									
				else
					Response.Write "<select name=""intStatus"" onchange=""javascript:checkStatus(this.selectedIndex,'" & strWarning & "', '" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "');"">"														
				end if 
				Response.Write "<option value=""-1""></option>"										
				do while rstData("question") = 14															
					' show the language
					Response.write "<option value=""" & rstData("option") & """"
					if aStudent(18,0) = rstData("option") then 
						Response.Write " selected=""selected"""
					end if 
					Response.Write ">" & rstData(strLanguage) & "</option>"
					rstData.MoveNext 
				loop
				Response.Write "</select>"
													
				rstData.MovePrevious 
			' Sex											
			case 3
				rstData.MoveNext 
													
				Response.Write "<select name=""intSex"">"
				Response.Write "<option value=""-1""></option>"										
				do while rstData("question") = 3															
					' show the language
					Response.write "<option value=""" & rstData("option") & """"
					if aStudent(4,0) = rstData("option") then 
						Response.Write " selected=""selected"""
					end if 
					Response.Write ">" & rstData(strLanguage) & "</option>"
					rstData.MoveNext 
				loop
				Response.Write "</select>"
													
				rstData.MovePrevious 
			' Postal Code
			case 4
				Response.Write "<input type=""text"" size=""10"" name=""strPostal"" value=""" & aStudent(6,0) & """ maxlength=""7"" />" 
			' ESL
			case 8
				rstData.MoveNext 
				
				'2015-12-10 - do not display for alberta
				if session("province") <>3 then 														
					Response.Write "<select name=""intESL"">"
					Response.Write "<option value=""-1""></option>"										
					do while rstData("question") = 8															
						' show the language
						Response.write "<option value=""" & rstData("option") & """"
						if aStudent(19,0) = rstData("option") then 
							Response.Write " selected=""selected"""
						end if 
						Response.Write ">" & rstData(strLanguage) & "</option>"
						rstData.MoveNext 
					loop
					Response.Write "</select>"
														
					rstData.MovePrevious 
				end if 
			' Special Needs
			case 7
				rstData.MoveNext 
				'2015-12-10 - do not display for alberta
				if session("province") <>3 then 									
				Response.Write "<select name=""intSpecial"" id=""intSpecial"" onChange=""javascript:chooseOnlyOne(this.name,'" & strLanguage & "','" & session("province") & "');"">"
				Response.Write "<option value=""-1""></option>"										
				do while rstData("question") = 7															
					' show the language
					Response.write "<option value=""" & rstData("option") & """"
					if aStudent(20,0) = rstData("option") then 
						Response.Write " selected=""selected"""
					end if 
					Response.Write ">" & rstData(strLanguage) & "</option>"
					rstData.MoveNext 
				loop
				Response.Write "</select>"
				
				if session("province") = 3 then 
					Response.Write "</font></td>"
					Response.Write "<td rowspan=""3"" align=""center"">"
					response.write "<table border=""1"" width=""250""><tr><td><font class=""subheaderBlue"">" & strABSpecialEduCode & "</font></td></tr></table>"
				end if
				end if 
				rstData.MovePrevious
				
				' Aboriginal
				case 11
				rstData.MoveNext
				'2015-12-10 - do not display for alberta
				if session("province") <>3 then 
					Response.Write "<select name=""intAboriginal"">"
				Response.Write "<option value=""-1""></option>"										
				do while rstData("question") = 11															
					' show the language
					Response.write "<option value=""" & rstData("option") & """"
					if aStudent(21,0) = rstData("option") then 
						Response.Write " selected=""selected"""
					end if 
					Response.Write ">" & rstData(strLanguage) & "</option>"
					rstData.MoveNext 
				loop
				Response.Write "</select>"
				end if 									
				rstData.MovePrevious 
			' French Immersion
			case 9
				rstData.MoveNext 
													
				Response.Write "<select name=""intFrench"">"
				Response.Write "<option value=""-1""></option>"										
				do while rstData("question") = 9															
					' show the language
					Response.write "<option value=""" & rstData("option") & """"
					if aStudent(22,0) = rstData("option") then 
						Response.Write " selected=""selected"""
					end if 
					Response.Write ">" & rstData(strLanguage) & "</option>"
					rstData.MoveNext 
				loop
				Response.Write "</select>"
													
				rstData.MovePrevious 
			' Other Immersion
			case 10
				rstData.MoveNext 
													
				Response.Write "<select name=""intOther"">"
				Response.Write "<option value=""-1""></option>"										
				do while rstData("question") = 10
						' show the language
						Response.write "<option value=""" & rstData("option") & """"
						if aStudent(23,0) = rstData("option") then 
							Response.Write " selected=""selected"""
						end if 
						Response.Write ">" & rstData(strLanguage) & "</option>"
					rstData.MoveNext 
				loop
				Response.Write "</select>"
													
				rstData.MovePrevious 
			' Communicates
			case 13			
				rstData.MoveNext								
				Response.Write "<select name=""intCommunicates"">"
				Response.Write "<option value=""-1""></option>"										
				do while rstData("question") = 13
					' show the language
					Response.write "<option value=""" & rstData("option") & """"
					if aStudent(25,0) = rstData("option") then 
						Response.Write " selected=""selected"""
					end if 
					Response.Write ">" & rstData(strLanguage) & "</option>"
					rstData.MoveNext 
				loop
				Response.Write "</select>"

				rstData.MovePrevious 
            
      ' Repeat the grade
			case 15				
				rstData.MoveNext		
				Response.Write "<select name=""intRepeat"">"
				Response.Write "<option value=""-1""></option>"			
        ' not alberta or manitoba then last question
        'if session("province") <> 3  then 
				  do while not rstData.eof	
                ' show the language
						  Response.write "<option value=""" & rstData("option") & """"
						  if aStudent(24,0) = rstData("option") then 
							  Response.Write " selected=""selected"""
						  end if 
						  Response.Write ">" & rstData(strLanguage) & "</option>"					
					  rstData.MoveNext                
				  loop
          rstData.MovePrevious     
        'else
         ' do while rstData("question") = 15					
         '     ' show the language
				'		  Response.write "<option value=""" & rstData("option") & """"
				'		  if aStudent(24,0) = rstData("option") then 
				'			  Response.Write " selected=""selected"""
				'		  end if 
				'		  Response.Write ">" & rstData(strLanguage) & "</option>"					
				'	  rstData.MoveNext                
				'  loop
        'end if               
						
				Response.Write "</select>"
				
      ' language delay
      case 17				
				rstData.MoveNext	
'2015-12-10 - do not display for alberta
				if session("province") <>3 then 				
				Response.Write "<select name=""intLanguageDelay"" id=""intLanguageDelay"" onChange=""javascript:chooseOnlyOne(this.name,'" & strLanguage & "','" & session("province") & "');"">"
				Response.Write "<option value=""-1""></option>"										
				do while rstData("question") = 17					
						' show the language
						Response.write "<option value=""" & rstData("option") & """"
						if aStudent(27,0) = rstData("option") then 
							Response.Write " selected=""selected"""
						end if 
						Response.Write ">" & rstData(strLanguage) & "</option>"					
					rstData.MoveNext                
				loop
				Response.Write "</select>"
				end if 
				rstData.MovePrevious    
       ' disability
       case 18				
				rstData.MoveNext		
				'2015-12-10 - do not display for alberta
				if session("province") <>3 then 
				Response.Write "<select name=""intDisability"" id=""intDisability"" onChange=""javascript:chooseOnlyOne(this.name,'" & strLanguage & "','" & session("province") & "');"">"
				Response.Write "<option value=""-1""></option>"										
				do while rstData("question") = 18					
						' show the language
						Response.write "<option value=""" & rstData("option") & """"
						if aStudent(28,0) = rstData("option") then 
							Response.Write " selected=""selected"""
						end if 
						Response.Write ">" & rstData(strLanguage) & "</option>"					
					rstData.MoveNext                
				loop
				Response.Write "</select>"
				end if 
				rstData.MovePrevious         
       
      ' immigrant 
       case 16
        rstData.MoveNext	
        if session("province") = 2 then 
          Response.Write "<select name=""intImmigrant"">"
				  Response.Write "<option value=""-1""></option>"
          '****************************************
          ' note the .eof - if a question is added you must change this to while question = 16
          '****************************************
          do while rstData("question") = 16      
            Response.write "<option value=""" & rstData("option") & """"
            if aStudent(26,0) = rstData("option") then 
              Response.Write " selected=""selected"""
            end if 
            Response.Write ">" & rstData(strLanguage) & "</option>"       					    
    
					  rstData.MoveNext                
				  loop
           
				  Response.Write "</select>"
           
          rstData.MovePrevious   
  
        end if            
		end select
     end if  
    'if rstData("question") < 16 then
		  Response.Write "</font></td>"
		  'response.write getHelpLink
		if session("province") =3 then 
			if (rstdata("question")< 6 or rstData("question") >8) AND rstData("question")<> 11 AND rstData("question")<17 then			
				response.write getHelpLink(session("Province"), "Demographics", strQuestion)
			end if 
		else
			response.write getHelpLink(session("Province"), "Demographics", strQuestion)
		end if 
		  Response.Write "</tr>"
    'else
     ' if session("province") = 2 then 
     '   Response.Write "</font></td>"
		 '   Response.Write "</tr>"
     ' end if
    'end if 
    
		rstData.movenext

	loop
	Response.Write "</table>"	
  
  
  
	' header							
	Response.Write "<br />&nbsp;<input type=""button"" onclick=""javascript:goWindow('" & strLink & "','Guide','490','500','top=0,left=125,resizable=yes');"" name=""btnGuide"" value=""GUIDE"" title=""GUIDE"" />&nbsp;&nbsp;<font class=""subheaderBlue"">"	
	if strLanguage="English" Then 
		response.write "Demographics"
	else
		response.write "Démographique"
	end if 
	response.write "</font>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','A','_Demographics');"" class=""bigLinkRed"">Section A</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','B','_Demographics')"" class=""bigLinkRed"">Section B</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','C','_Demographics')"" class=""bigLinkRed"">Section C</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','D','_Demographics')"" class=""bigLinkRed"">Section D</a>&nbsp;&nbsp; <a href=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','E','_Demographics')"" class=""bigLinkRed"">Section E</a>"
	if NOT blnLock then 
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type=""button"" onclick=""javascript:goSaveEDI('" & strEDIYear & "','" & strSite & "','" & strSchool & "','" & strTeacher & "','" & strClass & "','" & strChild & "','_Demographics','_Demographics')"" name=""btnSave"" value=""" & lblSaveEDI & """ title=""" & lblSaveEDI & """ />"
	end if 
	response.write "<br />"
	%>
	
		