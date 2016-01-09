<%
' updates the counter
call open_adodb(conn,"MACEDI")
		
strQuery = "INSERT INTO VisitorStatistics (IPAddress,VisitDate,VisitTime, Country, ProvinceState,City, Browser, Version, Agent,URL,Querystring, sessionID) VALUES('" & strClient & "','" & date() &"','" & time() &"','" & strCountry &"','" & strProvince &"','" & strCity &"','" & strAgentName &"','" & strVersion &"','" & strAgent &"','" & strURL & "','" & strQueryString &"','" & Session.sessionID & "')" 
response.write strQuery & "<br>" & request.querystring
Conn.execute(strQuery)		

set rstCounter = Server.CreateObject("ADODB.Recordset")
strQuery = "SELECT count(1) as intVis FROM VisitorStatistics"
rstCounter.Open strQuery, conn
		
intCounter = rstCounter("intVis")
		
call close_adodb(rstCounter)		
call close_adodb(conn)

if intVersion > 0 then 
	Response.Write "<a href=""javascript:goLayer('idVisitor');"" class=""regLink"">Visitor Information :</a>"
	Response.Write "<br />"
	Response.Write "<layer name=""idVisitor"" visibility=""hide"" zIndex=""0"">"
else
	Response.Write "<a href=""javascript:goDiv('idVisitor');goArrow();"" class=""regLink""><img src=""../files/dn.ico"" border=""no"" height=""15"" id=""imgArrow1"" /> Visitor Information <img src=""../files/dn.ico"" border=""no"" height=""15"" id=""imgArrow"" /></a>"
	Response.Write "<br />"
	Response.Write "<div id=""idVisitor"" style=""display:'none'"">"
end if 
%>

	<dd /> 
		<table border="1" cellspacing="1" cellpadding="0" bordercolor="navy" valign="bottom">
			<tr>
				<td>
					<table cellspacing="0" cellpadding="0" border="1" valign="top" bgcolor="WhiteSmoke">
						<tr>
							<td width="150" nowrap align="center">
								<font class="regHeader">&nbsp;IP Address:</font>
							</td>
							<td width="165" nowrap align="center">
								<font class="regHeader">&nbsp;Browser :</font>
							</td>
							<td width="100" nowrap align="center">
								<font class="regHeader">&nbsp;Number :</font>
							</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td>
					<table cellspacing="0" cellpadding="0" border="1" valign="top">
						<tr>
							<td width="150" nowrap align="center">
								<font class="regText">&nbsp;<%=strClient%>&nbsp;</font>
							</td>
							<td width="165" nowrap align="center">
								<font class="regText">&nbsp;<%=strVersion%>&nbsp;</font>
							</td>
							<td width="100" nowrap align="center">
								<font class="regText">&nbsp;<%=intCounter%>&nbsp;</font>
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>	
<%
if intVersion > 0 then 
	Response.Write "</layer>"
else
	Response.Write "</div>"
end if
%>