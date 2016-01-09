
<%
	if not rstFeedback.eof then 	

		' get all the feedback questions
		strSql = "SELECT F.[English], F.[French], F.[isSectionHeader], F.[Section], F.[Question], F.[Option], F.[isHeader],  C.[Language], C.[intOptions], C.Col1, C.Col2, C.Col3, C.Col4, C.Col5, C.Col6 FROM Page_TeacherFeedback F  LEFT JOIN [Column Headers] C ON F.[HeaderType] = C.[HID] WHERE (C.[Language]='" & strLanguage & "' OR C.[Language] Is Null) AND f.iid >= 6 ORDER BY F.[isSectionHeader], F.[Section], F.[Question], F.[Option], F.[isHeader] DESC"
  
		'open the Teacher Feedback Questions
		rstQuestions.Open strSql, conn,1  
  
  	' bln for inner table existance
		blnTable = false
		intOptions = 0 
		intRow = 0

		' table will be <spacer Column><question><6 columns>
		Response.Write "<table border=""0"" cellpadding=""0"" cellspacing=""0"" width=""750"" align=""center"">"			      
		do while not rstQuestions.EOF 
      ' write the header row if it is a header
      if rstQuestions("isHeader") then 
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
        Response.Write "<td align=""left"" colspan=""2""><font class=""subHeaderBlue"">" & rstQuestions(strLanguage) & "</font></td>" 
        
    		select case rstQuestions("Section")
				  case 8            
			      	response.write "<tr><td colspan=""2"" align=""center""><textarea rows=""5"" cols=""90"" name=""strmostBeneficial"">" &  rstFeedback("strMostBeneficial") & "</textarea></td></tr>"
          case 9
			      	response.write "<tr><td colspan=""2"" align=""center""><textarea rows=""5"" cols=""90"" name=""strCommentse-EDI"">" &  rstFeedback("strCommentse-EDI") & "</textarea></td></tr>"  	  		 					          
			  end select
        ' get the number of options 
        if not isnull(rstQuestions("intOptions")) Then 
          intOptions = rstQuestions("intOptions")
          for intCol = 1 to intoptions
            Response.Write "<td align=""center"" valign=""middle"" width=""70""><font class=""boldTextBlack"">" & rstQuestions("Col" & intCol) & "</font></td>"
            ' store the header values for the title below 2005 09 15
            aHeader(intCol) = rstQuestions("Col" & intCol)
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
						Response.Write "<tr><td align=""left"" valign=""middle""><font class=""boldTextBlack"">&nbsp;&nbsp;" 
						' only write the questions
						if rstQuestions("option") = 0 then 
							if rstQuestions("question") < 10 then 
								Response.Write "&nbsp;&nbsp;"
							end if 
							Response.Write rstQuestions("question") & " ) &nbsp;&nbsp;</td><td><font class=""boldTextBlack"">" &  rstQuestions(strLanguage) & "</font>"
						end if
																	
						'select case rstQuestions("question")
														
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
					if rstQuestions("option") = 0 then 
						if rstQuestions("question") < 10 then 
							Response.Write "&nbsp;&nbsp;"
						end if 
						Response.Write rstQuestions("question") & " ) &nbsp;&nbsp;</font></td><td align=""left""><font class=""boldTextBlack"">" &  rstQuestions(strLanguage) & "</font>"
						strQuestion = "intQ" & rstQuestions("question")
					elseif rstQuestions("option") >= 97 then 
						if rstQuestions("question") < 10 then 
							Response.Write "&nbsp;&nbsp;"
						end if 
						Response.Write rstQuestions("question") & chr(rstQuestions("option")) & " ) &nbsp;&nbsp;</font></td><td align=""left""><font class=""boldTextBlack"">" &  rstQuestions(strLanguage) & "</font>"
						strQuestion = "intQ" & rstQuestions("Question") & chr(rstQuestions("option"))
					end if
					
					
																	
					select case rstQuestions("question")
						case 3
              if rstQuestions("option")= 104 then 
							  if strLanguage = "English" Then
								  response.write "<p align=""left""><font class=""boldTextBlack"">If yes, please specify:</font>"			
							  else
								  response.write "<p align=""left""><font class=""boldTextBlack"">Dans l’affirmative, précisez:</font>"			
							  end if 							
							response.write "<input type=""text"" maxlength=""50"" name=""strQ3Other"" size=""45"" value=""" & rstFeedback("strQ3Other") & """></p>"						
              end if 
					end select
					
					Response.Write "</td>"
												
					' write the radio option buttons
					for intCol = 1 to intoptions
						Response.Write "<td align=""center"" valign=""middle"" width=""70""><input type=""radio"" title=""" & aHeader(intCol) &""" name=""" & strQuestion & """ value=""" & intCol & """"
						if rstFeedback(strQuestion) = intCol then 
							response.write " checked"
						end if 
						response.write "></td>"
					next 
					Response.Write "</tr>"
				end if 										
			end if 
          
			rstQuestions.movenext
		loop
		' close the last inner table if it is open
      
		if blnTable then 
			Response.Write "</table>"
		end if 
									
	' end the row
	Response.Write "</td></tr>"
	Response.Write "</table>"	
			
else
  strError = "<font class=""regtextred"">No data on teacher - " & format(intTeacher,"00000000") & "</font>"
end if
' close and kill the additional questions recordset
call close_adodb(rstQuestions)
      %>