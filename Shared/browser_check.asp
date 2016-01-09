<%
' gets the clients IP address
strClient = Request.ServerVariables("REMOTE_ADDR")
strURL =  Request.ServerVariables("PATH_INFO")
strQueryString = Request.ServerVariables("QUERY_STRING")
' removes "mozilla/" from the strAgent specs
strAgent = replace(Request.ServerVariables("HTTP_USER_AGENT"), "Mozilla/", "")
'Dim aryAgentElems, strOS
'aryAgentElems = Split(strAgent, ";")
'strOS = aryAgentElems(2)
'Parse = strOS 

'determines Netscape or Internet Explorer or Unknown
if instr(1,strAgent,"MSIE",1) then 	
	'EXPLORER : VERSION
	strAgentName = "Internet Explorer"	
	strVersion = Mid(strAgent,instr(1,strAgent,"MSIE ",1)+5,4)	
elseif instr(1,strAgent,"Trident",1) then 	
	'EXPLORER : VERSION
	strAgentName = "Internet Explorer"	
	'strVersion = Mid(strAgent,instr(1,strAgent,"rv: ",1)+3,4)	
	strVersion = Mid(strAgent,instr(1,strAgent,"rv:",1)+3,4)	
elseif instr(1,strAgent,"Firefox",1) then 		
	strAgentName = "Firefox"		
	strVersion = Mid(strAgent,instr(1,strAgent,"Firefox/",1)+8,5)		
elseif instr(1,strAgent,"Chrome",1) then 		
	strAgentName = "Chrome"		
	strVersion = Mid(strAgent,instr(1,strAgent,"Chrome/",1)+7,13)		
else
	'UNKNOWN : VERSION
	strAgentName = "Unknown"		
	strVersion = "Unknown"
end if 	
	
' returns:
'		strClient - IP address
'		strAgent - full browser string
'		intNumber - browser version
'		strVersion - version of browser
'		intVersion - (0) IE (1) NETSCAPE (2) OTHER

On Error Resume Next ' prevent tossing unhandled exception
'response.write session.sessionID & "<br>"
'response.write Request.Cookies("e-EDI")("SessionID") & "<br>"
'if session.sessionID <> Request.Cookies("e-EDI")("SessionID") then 
'response.write "not same"
'else
'response.write "same"'
'end if 
' check the cookie - if same session id than do not look it up
if session.sessionID <> Request.Cookies("e-EDI")("SessionID") then 
	 Dim URL, objXML, value

	 Set objXML = Server.CreateObject("MSXML2.DOMDocument.6.0")
	 URL = "http://api.ipinfodb.com/v3/ip-city/?key=1d77477f9d8620e5523f053645dd892fe97374f6b666225dc92f28f57c0e733d&&format=xml&ip=" & strClient
	 
	'<statusCode>OK</statusCode>
	'<statusMessage/>
	'<ipAddress>72.143.61.138</ipAddress>
	'<countryCode>CA</countryCode>
	'<countryName>CANADA</countryName>
	'<regionName>ONTARIO</regionName>
	'<cityName>TORONTO</cityName>
	'<zipCode>M3B 0A3</zipCode>
	'<latitude>43.7001</latitude>
	'<longitude>-79.4163</longitude>
	'<timeZone>-04:00</timeZone>

	 if (NOT Err.Description = "") then
		 'You could use the following to email an alert
		 'Response.Write("An error occured when retrieving data from an external source.")
		 'Response.Write(Err.Description)
		 'Response.End
	 else
		 objXML.setProperty "ServerHTTPRequest", True
		 objXML.async = False
		 objXML.Load URL
		 'Set oRoot = objXML.selectSingleNode("//response")
		'strIP = objXML.documentElement.childNodes(0).text
		 strCountry=objXML.documentElement.childNodes(4).text 
		 strProvince= objXML.documentElement.childNodes(5).text 
		 'strRegion = objXML.documentElement.childNodes(5).text
		 strCity = objXML.documentElement.childNodes(6).text
		 'strRcode = objXML.documentElement.childNodes(4).text	 
		 'Response.Write "after setting value: " & Err.Description & "<br>"
		 Response.Cookies("e-EDI")("SessionID") = session.sessionID
		 Response.Cookies("e-EDI")("Country") = strCountry
		 Response.Cookies("e-EDI")("Province") = strProvince
		 Response.Cookies("e-EDI")("City") = strCity
	end if 
else
	strCountry = Request.Cookies("e-EDI")("Country")
	strProvince = Request.Cookies("e-EDI")("Province")
	strCity = Request.Cookies("e-EDI")("City")
end if 
%>