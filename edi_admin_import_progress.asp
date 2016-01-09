<!-- #include virtual="/shared/admin_security.asp" -->
<%
' if the user has not logged in they will not be able to see the page
if blnSecurity then
	dim connXLS
	dim rstXLS
	dim aSites,aSchools,aTeachers,aClasses,aChildren
	dim intSites, intSchools, intTeachers, intClasses, intChildren
	 
	on error resume next 
%>
	<html>
	<head
	<!-- Start CSS files-->
		<link rel="stylesheet" type="text/css" href="Styles/edi.css">
	<!-- End CSS files -->
		<title>e-EDI Import Summary Report</title>
	</head>
	<body>
		<!-- #include virtual="/shared/page_header.inc" -->
		
		<form name="Screens" method="POST" action="edi_admin_import.asp"> 
			<a class="reglinkMaroon" href="edi_admin.asp">Home</a>&nbsp;<font class="regtextblack">></font>&nbsp;<a class="reglinkMaroon" href="edi_admin_import.asp">Data Import File Selection</a>&nbsp;<font class="regtextblack">></font>&nbsp;<font class="boldtextblack">Import Summary Report</font>
			<table border="1" bordercolor="006600" cellpadding="0" cellspacing="0" width="760">
			<tr><td>
				<table border="0" cellpadding="0" cellspacing="0" width="750" align="center">
					<tr><td align="center"><font class="headerBlue">Import Summary Report</font></td></tr>
					
					<%
					Set Upload = Server.CreateObject("Persits.Upload")
					' This is needed to enable the progress indicator
					Upload.ProgressID = Request.QueryString("PID")
					
					' path to the upload directory
					strPath = server.mappath("edi_uploads/")

					' count the files being uploaded
					Count = Upload.Save(strPath)

					' display success
					'Response.Write Count & " file uploaded successfully " & strPath
		
					' kill the upload object
					'set upload = nothing
					'For Each File in Upload.Files
					'	Response.Write File.Name & "= " & File.Path & " (" & File.Size &" bytes)<BR>"
					'Next
					Set File = Upload.Files("FILE1")
					If Not File Is Nothing Then
					   strFileName = File.Path 'File.filename	
					   Response.Write "<font class=""regtextgreen"">File (" & File.filename & ") Uploaded Successfully (" & File.Size/1024 & " kb )</font><br />"
						'response.write strFilename
					End If 
					set file = nothing
					set upload = nothing 
					
					'Response.Redirect "edi_admin.asp"
					'For Each Item in Upload.Form
					'	Response.Write Item.Name & "= " & Item.Value & "<BR>"
					'Next

					'Response.Write strPath 
					'response.write "<script language=""javascript"">"
					'	response.write "document.forms.Screens.filepath.value = '" & strPath & "';"
				'		Response.write "document.forms.Screens.submit();"
				'	Response.Write "</script>"
		
					' connect to EDI Data
					'call open_adodb(conn, "DATA")
                    call open_adodb(conn, "MACEDI")
					
					' connect to xls file
					call XLSconnection(strFilename)
					
					call get_Data
									
					intErrors = 0
					Response.Write "<tr><td><br><br><font class=""boldtextBlack"">Sites</font><br></td></tr>"
					for intRow = 0 to intSites - 1
						'strSQL = "INSERT INTO sites ( intSiteID, strName, strCoordinator, strEmail, strAddress, strCity, intProvince, strPostal, strPhone, strFax) VALUES" & _
						'			"(" & aSites(0,intRow) & "," & checkNull(aSites(1,intRow)) & "," & checkNull(aSites(2,intRow)) & "," & checkNull(aSites(3,intRow)) & "," & checkNull(aSites(4,intRow)) & "," & checkNull(aSites(5,intRow)) & "," & checknull(aSites(6,intRow)) & "," & checkNull(aSites(7,intRow)) & "," & checkNull(aSites(8,intRow)) & "," & checkNull(aSites(9,intRow)) & ")"
						strSQL = "INSERT INTO sites ( intSiteID, strName, strCoordinator, strEmail) VALUES" & _
									"(" & aSites(0,intRow) & "," & checkNull(aSites(1,intRow)) & "," & checkNull(aSites(2,intRow)) & "," & checkNull(aSites(3,intRow)) & ")"
						
						'Response.Write strSql & "<br>"		
						conn.execute strSql	
						
						if conn.errors.count > 0 then 
							if conn.errors(0).number = -2147467259 AND instr(conn.errors(0).description, "duplicate") > 0 then							
								Response.Write "<tr><td><font class=""regtextred"">The site " & aSites(0,intRow) & " is already in the database.</font></td></tr>"							
							else
								Response.Write "<tr><td><font class=""regtextred"">The site " & aSites(0,intRow) & " could not be inserted due to error number " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/>" & strSql & "</font></td></tr>"									
							end if 
						
							intErrors = intErrors + 1
							conn.errors.clear
						else
							Response.Write "<tr><td><font class=""regtextgreen"">The site " & aSites(0,intRow) & " was successfully inserted.<br/><br/></font></td></tr>"
						end if 
					next 
				
					
					erase aSites
					Response.Write "<tr><td><font class=""regtextblack"">" & intSites & " Sites, " & intErrors & " Errors<br/><br/></font></td></tr>"
											
					Response.Write "<tr><td><br><br><font class=""boldtextBlack"">Schools</font><br></td></tr>"
					
					intErrors = 0
					for intRow = 0 to intSchools - 1
						'strSQL = "INSERT INTO schools ( intSiteID, intSchoolID, strName, strAddress, strCity, intProvince, strPostal, strPhone, strFax  ) VALUES" & _
						'			"(" & aSchools(0,intRow) & "," & aSchools(1,intRow) & "," & checkNull(aSchools(2,intRow)) & "," & checkNull(aSchools(3,intRow)) & "," & checkNull(aSchools(4,intRow)) & "," & checkNull(aSchools(5,intRow)) & "," & checkNull(aSchools(6,intRow)) & "," & checkNull(aSchools(7,intRow)) & "," & checkNull(aSchools(8,intRow)) & ")"
						strSQL = "INSERT INTO schools ( intSiteID, intSchoolID, strName, intProvince,intELP) VALUES" & _
									"(" & aSchools(0,intRow) & "," & aSchools(1,intRow) & "," & checkNull(aSchools(2,intRow)) & "," & checkNull(aSchools(3,intRow)) & ",0)"

						'Response.Write strSql & "<br>"			
						conn.execute strSql	
						
						if conn.errors.count > 0 then 
							if conn.errors(0).number = -2147467259 AND instr(conn.errors(0).description, "duplicate") > 0 then							
								Response.Write "<tr><td><font class=""regtextred"">The school " & aSchools(1,intRow) & " is already in the database.</font></td></tr>"							
							else
								Response.Write "<tr><td><font class=""regtextred"">The school " & aSchools(1,intRow) & " could not be inserted due to error number " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/></font></td></tr>"
							end if 
							intErrors = intErrors + 1
							conn.errors.clear
						else
							Response.Write "<tr><td><font class=""regtextgreen"">The school " & aSchools(1,intRow) & " was successfully inserted.<br/><br/></font></td></tr>"
						end if 
					next 
					Response.Write "<tr><td><font class=""regtextblack"">" & intSchools & " Schools, " & intErrors & " Errors<br/><br/></font></td></tr>"
					erase aSchools
					
					Response.Write "<tr><td><br><br><font class=""boldtextBlack"">Teachers</font><br></td></tr>"
					
					intErrors = 0
					for intRow = 0 to intTeachers - 1
						'strSQL = "INSERT INTO teachers ( intSchoolID, intTeacherID, strName, strEmail, strPassword ) VALUES" & _
						'			"(" & aTeachers(0,intRow) & "," & aTeachers(1,intRow) & "," & checkNull(aTeachers(2,intRow)) & "," & checkNull(aTeachers(3,intRow)) & "," & checknull(right("000" & aTeachers(1,intRow),8)) & ")"
						strSQL = "INSERT INTO teachers ( intSchoolID, intTeacherID, strName, strEmail, strPassword ) VALUES" & _
									"(" & aTeachers(0,intRow) & "," & aTeachers(1,intRow) & "," & checkNull(aTeachers(2,intRow)) & "," & checkNull(aTeachers(3,intRow)) & "," & checknull(right("000" & aTeachers(1,intRow),8)) & ")"
'						'Response.Write strSql & "<br>"			
						conn.execute strSql	
						
						if conn.errors.count > 0 then 
							if conn.errors(0).number = -2147467259 AND instr(conn.errors(0).description, "duplicate") > 0 then							
								Response.Write "<tr><td><font class=""regtextred"">The teacher " & aTeachers(1,intRow) & " is already in the database.</font></td></tr>"							
							else
								Response.Write "<tr><td><font class=""regtextred"">The teacher " & aTeachers(1,intRow) & " could not be inserted due to error number " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/></font></td></tr>"
							end if 
							intErrors = intErrors + 1
							conn.errors.clear
						else
							Response.Write "<tr><td><font class=""regtextgreen"">The teacher " & aTeachers(1,intRow) & " was successfully inserted.<br/><br/></font></td></tr>"
						end if 
						'***********************************************************
						' insert teacher participation records - added  - 2007-11-20
						strSQL = "INSERT INTO teacherParticipation ( intTeacherID) VALUES" & _
									"(" & aTeachers(1,intRow) & ")"
'						'Response.Write strSql & "<br>"			
						conn.execute strSql	
						
						if conn.errors.count > 0 then 
							if conn.errors(0).number = -2147467259 AND instr(conn.errors(0).description, "duplicate") > 0 then							
								Response.Write "<tr><td><font class=""regtextred"">The teacher participation record " & aTeachers(1,intRow) & " is already in the database.</font></td></tr>"							
							else
								Response.Write "<tr><td><font class=""regtextred"">The teacher participation record " & aTeachers(1,intRow) & " could not be inserted due to error number " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/></font></td></tr>"
							end if 
							intErrors = intErrors + 1
							conn.errors.clear
						else
							Response.Write "<tr><td><font class=""regtextgreen"">The teacher participation record " & aTeachers(1,intRow) & " was successfully inserted.<br/><br/></font></td></tr>"
						end if 
						'***********************************************************
						
						'***********************************************************
						' insert teacher training feedback records - added  - 2008-11-24
						strSQL = "INSERT INTO teacherTrainingFeedback ( intTeacherID) VALUES" & _
									"(" & aTeachers(1,intRow) & ")"
'						'Response.Write strSql & "<br>"			
						conn.execute strSql	
						
						if conn.errors.count > 0 then 
							if conn.errors(0).number = -2147467259 AND instr(conn.errors(0).description, "duplicate") > 0 then							
								Response.Write "<tr><td><font class=""regtextred"">The teacher training feedback record " & aTeachers(1,intRow) & " is already in the database.</font></td></tr>"							
							else
								Response.Write "<tr><td><font class=""regtextred"">The teacher training feedback record " & aTeachers(1,intRow) & " could not be inserted due to error number " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/></font></td></tr>"
							end if 
							intErrors = intErrors + 1
							conn.errors.clear
						else
							Response.Write "<tr><td><font class=""regtextgreen"">The teacher training feedback record " & aTeachers(1,intRow) & " was successfully inserted.<br/><br/></font></td></tr>"
						end if 
						'***********************************************************
					next 
					erase aTeachers
					Response.Write "<tr><td><font class=""regtextblack"">" & intTeachers & " Teachers, " & intErrors & " Errors<br/><br/></font></td></tr>"
						
					Response.Write "<tr><td><br><br><font class=""boldtextBlack"">Classes</font><br></td></tr>"
					
					intErrors = 0
					for intRow = 0 to intClasses - 1
						'strSQL = "INSERT INTO classes ( intTeacherID, intClassID, intLanguage ) VALUES" & _
						'	"(" & aClasses(0,intRow) & "," & aClasses(1,intRow) & "," & aClasses(2,intRow) & ")"
						strSQL = "INSERT INTO classes ( intTeacherID, intClassID) VALUES" & _
									"(" & aClasses(0,intRow) & "," & aClasses(1,intRow) & ")"
						'Response.Write strSql & "<br>"			
						conn.execute strSql	
						
						if conn.errors.count > 0 then 
							if conn.errors(0).number = -2147467259 AND instr(conn.errors(0).description, "duplicate") > 0 then							
								Response.Write "<tr><td><font class=""regtextred"">The class " & aClasses(1,intRow) & " is already in the database.</font></td></tr>"							
							else
								Response.Write "<tr><td><font class=""regtextred"">The class " & aClasses(1,intRow) & " could not be inserted due to error number " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/></font></td></tr>"
							end if 
							intErrors = intErrors + 1
							conn.errors.clear
						else
							Response.Write "<tr><td><font class=""regtextgreen"">The class " & aClasses(1,intRow) & " was successfully inserted.<br/><br/></font></td></tr>"
						end if 
					next 
					erase aClasses
					
					Response.Write "<tr><td><font class=""regtextblack"">" & intClasses & " Classes, " & intErrors & " Errors<br/><br/></font></td></tr>"
					Response.Write "<tr><td><br><br><font class=""boldtextBlack"">Children</font><br></td></tr>"
					
					intErrors = 0
					'Class ID - Child ID - EDI ID - local ID - sex - dob - postal - French immersion
					for intRow = 0 to intChildren - 1
						strSQL = "INSERT INTO children ( intClassID, intChild, strEDIID,strLocalID, intSex,dtmDob,strPostal) VALUES" & _
									"(" & aChildren(0,intRow) & "," & aChildren(1,intRow) & "," & checkNull(aChildren(2,intRow)) & "," & checkNull(aChildren(3,intRow)) & "," & aChildren(4,intRow) & "," & checkNull(aChildren(5,intRow)) & "," & checkNull(aChildren(6,intRow)) & ")"
						'Response.Write strSql & "<br>"			
						conn.execute strSql	
			
						if conn.errors.count > 0 then 
							if conn.errors(0).number = -2147467259 AND instr(conn.errors(0).description, "duplicate") > 0 then							
								Response.Write "<tr><td><font class=""regtextred"">The child " & aChildren(2,intRow) & " is already in the database.</font></td></tr>"							
							else					
								Response.Write "<tr><td><font class=""regtextred"">The child " & aChildren(2,intRow) & " could not be inserted due to error number " & conn.errors(0).number & "<br /><br />Description : " & makeReadable(conn.errors(0).description) & "<br/><br/>" & strSql & "</font></td></tr>"
							end if 
							intErrors = intErrors + 1
							conn.errors.clear
						else
							' insert to other tables 
							conn.execute "INSERT INTO demographics (strEDIID) VALUES(" & checkNull(aChildren(2,intRow)) & ")"
							conn.execute "INSERT INTO sectionA (strEDIID) VALUES(" & checkNull(aChildren(2,intRow)) & ")"
							conn.execute "INSERT INTO sectionB (strEDIID) VALUES(" & checkNull(aChildren(2,intRow)) & ")"
							conn.execute "INSERT INTO sectionC (strEDIID) VALUES(" & checkNull(aChildren(2,intRow)) & ")"
							conn.execute "INSERT INTO sectionD (strEDIID) VALUES(" & checkNull(aChildren(2,intRow)) & ")"
							conn.execute "INSERT INTO sectionE (strEDIID) VALUES(" & checkNull(aChildren(2,intRow)) & ")"
							'response.write "INSERT EDI ID " & strEDIID
							Response.Write "<tr><td><font class=""regtextgreen"">The child " & aChildren(2,intRow) & " was successfully inserted.<br/><br/></font></td></tr>"
						end if 
					next 			
					erase aChildren
					
					Response.Write "<tr><td><font class=""regtextblack"">" & intChildren & " Children, " & intErrors & " Errors<br/><br/></font></td></tr>"
					call close_adodb(conn)
					call close_adodb(rstXLS)
					call close_adodb(connXLS)
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
end if


sub XLSconnection(strFile)
	'Creates an instance of an Active Server Component
	Set connXLS = Server.CreateObject("ADODB.Connection")
	
	With connXLS
		.Provider = "Microsoft.Jet.OLEDB.4.0"
		.Properties("Extended Properties").Value = "Excel 8.0;hdr=yes;"
		.Open strFile
		'"c:\websites\e-edica\www\edi_uploads\" 
	end with
	
	Set rstXLS = Server.CreateObject("ADODB.Recordset")
end sub

sub get_Data()
	' sites
	'strSQL = "SELECT DISTINCT  left(format(sch_id,'000000'),3) as siteID, site, coordinator, [Coordinator Email] as email, [Site Address] as address, [Site City] as city, [Site Province] as province,[Site Postal Code] as postal, [Site Phone] as phone,[Site Fax] as fax FROM [Sheet1$] WHERE sch_id <> 0 AND not isnull(site)"       
	strSQL = "SELECT DISTINCT  mid(format(edi_id,'000000000000000'),5,3) as siteID, site, coordinator, [Coordinator Email] as email FROM [Sheet1$] WHERE Not isnull(site)"       
    
	rstXLS.Open strSql, connXLS
	if not rstXLS.eof then 
		aSites = rstXLS.GetRows 
		
		' count of sites
		intSites = ubound(aSites,2) + 1
		rstXLS.Close
		
		' schools
		'strSQL = "SELECT DISTINCT  left(format(sch_id,'000000'),3) as siteID, format(sch_id,'000000') as schoolID, [school name] as school, [School Address] as address, [School City] as city, [School Province] as province,[School Postal Code] as postal, [School Phone] as phone,[School Fax] as fax FROM [Sheet1$] WHERE sch_id <> 0 AND not isnull([School Address])"       		
		'strSQL = "SELECT DISTINCT  left(format(edi_id,'00000000000'),3) as siteID, left(format(edi_id,'00000000000'),6) as schoolID, [school name] as school FROM [Sheet1$] WHERE not isnull(Site)"       		
		strSQL = "SELECT DISTINCT  mid(format(edi_id,'000000000000000'),5,3) as siteID, mid(format(edi_id,'000000000000000'),5,6) as schoolID, [school name] as school, [school province] as province FROM [Sheet1$] WHERE not isnull(Site)"       		

		rstXLS.Open strSql, connXLS
		if not rstXLS.eof then 
			aSchools = rstXLS.GetRows 
				
			' count of schools
			intSchools = ubound(aSchools,2) + 1
			rstXLS.Close
				
			' teachers
			'strSQl = "SELECT DISTINCT  format(sch_id,'000000') as schoolID, format(sch_id,'000000') & format(teach_id,'00') as teacherID, [teacher name] as teacher, [teacher email] as email FROM [Sheet1$] WHERE sch_id<>0 AND not isnull([teacher email])"       		
			strSQl = "SELECT DISTINCT  mid(format(edi_id,'000000000000000'),5,6) as schoolID, mid(format(edi_id,'000000000000000'),5,8)  as teacherID, [teacher name] as teacher, [teacher email] as email FROM [Sheet1$] WHERE not isnull([teacher email])"       		

			rstXLS.Open strSql, connXLS	
			if not rstXLS.eof then 
				aTeachers = rstXLS.GetRows 
							
				' count of teachers
				intTeachers = ubound(ateachers,2) + 1
				rstXLS.Close
							
				' classes
				'strSQl = "SELECT DISTINCT  format(sch_id,'000000') & format(teach_id,'00') as teacherID, format(class_id,'000000000'), iif(immersion<>'',2,1) FROM [Sheet1$] WHERE sch_id <> 0"       		
				strSQl = "SELECT DISTINCT  mid(format(edi_id,'000000000000000'),5,8) as teacherID, mid(format(edi_id,'000000000000000'),5,9) as class_id FROM [Sheet1$] WHERE Not isnull(site)"       		

				rstXLS.Open strSql, connXLS				
				if not rstXLS.eof then 
					aClasses = rstXLS.GetRows 
									
					' count of classes
					intClasses = ubound(aClasses,2) + 1
					rstXLS.Close
									
					' children
					'strSQl = "SELECT DISTINCT  format(class_id,'000000000'), format(stud_id,'00') ,format(edi_id,'00000000000')  , [child's local id] as localID,iif(gender='m', 1,2) ,format(dob,'dd-mmm-yyyy'), [postal code] FROM [Sheet1$] WHERE sch_id <> 0"       		
					strSQl = "SELECT DISTINCT  mid(format(edi_id,'000000000000000'),5,9) as class_id, right(format(edi_id,'000000000000000'),2) as stud_id ,format(edi_id,'000000000000000')  , [child's local id] as localID,iif(gender='m', 1,2) ,format(dob,'dd-mmm-yyyy'), [postal code] FROM [Sheet1$] WHERE Not isnull(site)"       		
					rstXLS.Open strSql, connXLS
					if not rstXLS.eof then 
						aChildren = rstXLS.GetRows 
										
						' count of classes
						intChildren = ubound(aChildren,2) + 1
						rstXLS.Close
					else
						intChildren = 0
					end if
				else
					intClasses = 0
					intChildren = 0
				end if  		
			else
				intTeachers = 0
				intClasses = 0
				intChildren = 0
			end if  				
		else
			intSchools = 0
			intTeachers = 0
			intClasses = 0
			intChildren = 0
		end if  		
	else
		intSites = 0
		intSchools = 0
		intTeachers = 0
		intClasses = 0
		intChildren = 0
	end if 
end sub 
%>