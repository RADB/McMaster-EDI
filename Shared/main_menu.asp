<%
Sub build_IE(intActive)
	Response.Write "<table width=""100%"" height=""60%"" cellpadding=""0"" cellspacing=""10"" border=""0"">"
		Response.Write "<tr><td width=""11%"" valign=""top"">"
			Response.Write "<table width=""100%"" cellpadding=""0"" cellspacing=""0"" border=""0"" bordercolor=""white"">"
				aName = array("Home", "NCAA", "NFL", "Resume", "Next")
				aLink = array("http://www.renner.ca/default.asp", "http://www.renner.ca/ncaa2002", "http://www.renner.ca/nfl2001", "http://www.renner.ca/resume", "http://www.renner.ca/1161")				
				call build_menu(intActive, aName, aLink)	
			Response.Write "</table>"
		Response.Write "</td>"
end sub

sub build_Other(intActive)
	Response.Write "<table width=""100%"" height=""60%"" cellpadding=""0"" cellspacing=""10"" border=""0"">"
		Response.Write "<tr><td width=""11%"" valign=""top"">"
			Response.Write "<table width=""100%"" cellpadding=""0"" cellspacing=""0"" border=""0"" bordercolor=""white"">"
				aName = array("Home", "NCAA", "NFL", "Resume", "Next")
				aLink = array("http://www.renner.ca/default.asp", "http://www.renner.ca/ncaa2002", "http://www.renner.ca/nfl2001", "http://www.renner.ca/resume", "http://www.renner.ca/1161")
				call build_menu(intActive, aName, aLink)	
			Response.Write "</table>"
		Response.Write "</td>"
end sub

sub build_menu(intActive, aName, aLink)
	for row = 0 to ubound(aName)
		if intActive = row then 
			Response.Write "<tr><td nowrap bgcolor=""slategray"" id=""td" & aName(row) & """>"
			Response.Write "<font class=""boldTextWhite"">" & aName(row) & "</font>"
		else
			Response.Write "<tr><td nowrap bgcolor=""black"" onMouseOver=""javascript:this.bgColor='slateGray';td" & aName(intActive) & ".bgColor='black';"" onMouseOut=""javascript:this.bgColor='Black';td" & aName(intActive) & ".bgColor='slateGray';"" onclick=""javascript:window.location='" & aLink(row) & "'"" id=""td" & aName(row) & """>"
			Response.Write "<a class=""menuLink"" href=""" & aLink(row) & """>" & aName(row) & "</a>"
		end if 
		Response.Write "</td></tr>"
	next
end sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Changed November 13, 2001
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'Sub build_IE(intActive)
'	Response.Write "<table width=""100%"" height=""60%"" cellpadding=""0"" cellspacing=""10"" border=""0"">"
'		Response.Write "<tr><td width=""11%"" valign=""top"">"
'			Response.Write "<table width=""100%"" cellpadding=""0"" cellspacing=""0"" border=""0"" bordercolor=""white"">"
'				aName = array("Home", "1161", "Information", "Mileage", "NCAA", "NFL", "Resume", "Pathfinder")
'				aLink = array("http://www.renner.ca/default.asp", "http://www.renner.ca/1161", "http://www.renner.ca/your_info.asp", "http://www.renner.ca/mileage", "http://www.renner.ca/ncaa2001", "http://www.renner.ca/nfl2001", "http://www.renner.ca/resume", "http://www.renner.ca/pathfinder")
'				call build_menu(intActive, aName, aLink)	
'			Response.Write "</table>"
'		Response.Write "</td>"
'end sub

'sub build_Other(intActive)
'	Response.Write "<table width=""100%"" height=""60%"" cellpadding=""0"" cellspacing=""10"" border=""0"">"
'		Response.Write "<tr><td width=""11%"" valign=""top"">"
'			Response.Write "<table width=""100%"" cellpadding=""0"" cellspacing=""0"" border=""0"" bordercolor=""white"">"
'				aName = array("Home", "1161", "Mileage", "NCAA", "NFL", "Resume", "Pathfinder")
'				aLink = array("http://www.renner.ca/default.asp", "http://www.renner.ca/1161", "http://www.renner.ca/mileage", "http://www.renner.ca/ncaa2001", "http://www.renner.ca/nfl2001", "http://www.renner.ca/resume", "http://www.renner.ca/pathfinder")
'				call build_menu(intActive, aName, aLink)	
'			Response.Write "</table>"
'		Response.Write "</td>"
'end sub

'sub build_menu(intActive, aName, aLink)
'	for row = 0 to ubound(aName)
'		if intActive = row then 
'			Response.Write "<tr><td nowrap bgcolor=""slategray"" id=""td" & aName(row) & """>"
'			Response.Write "<font class=""boldTextWhite"">" & aName(row) & "</font>"
'		else
'			Response.Write "<tr><td nowrap bgcolor=""black"" onMouseOver=""javascript:this.bgColor='slateGray';td" & aName(intActive) & ".bgColor='black';"" onMouseOut=""javascript:this.bgColor='Black';td" & aName(intActive) & ".bgColor='slateGray';"" onclick=""javascript:window.location='" & aLink(row) & "'"" id=""td" & aName(row) & """>"
'			Response.Write "<a class=""menuLink"" href=""" & aLink(row) & """>" & aName(row) & "</a>"
'		end if 
'		Response.Write "</td></tr>"
'	next
'end sub
%>
