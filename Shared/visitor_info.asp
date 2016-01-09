<%
' updates the counter
call open_adodb(connVisitor,"MACEDI")
strQuery = "INSERT INTO VisitorStatistics (IPAddress,VisitDate,VisitTime, Country, ProvinceState,City, Browser, Version, Agent,URL,Querystring,Form, Email, sessionID) VALUES(" & checkValue(strClient)& "," &checkValue(date())& "," &checkValue(time())& "," &checkValue(strCountry)& "," &checkValue(strProvince)& "," &checkValue(strCity)& "," &checkValue(strAgentName)& "," &checkValue(strVersion)& "," &checkValue(strAgent)& "," &checkValue(strURL)& "," &checkValue(strQueryString)& "," &checkValue(request.form)& "," &checkValue(session("id"))& "," &checkValue(Session.sessionID) &")" 
'response.write strQuery & "<br>" & request.querystring
ConnVisitor.execute(strQuery)		

'set rstCounter = Server.CreateObject("ADODB.Recordset")
'strQuery = "SELECT count(1) as intVis FROM VisitorStatistics"
'rstCounter.Open strQuery, conn
		
'intCounter = rstCounter("intVis")
		
'call close_adodb(rstCounter)		
call close_adodb(connVisitor)


' for each x in Request.ServerVariables
'   response.write(x & "=" & request.servervariables(x) & "<br>")
' next
' 
' response.write "Agent = " & strAgent & "<br>"
' response.write "Agent Name = " & strAgentName & "<br>"
' response.write "Version = " & strVersion & "<br>"
' response.write "Country = " & strCountry & "<br>"
' response.write "Province = " & strProvince & "<br>"
' response.write "City = " & strCity & "<br>"		 
'response.write request.form & "<br>" & len(request.form)
%>
