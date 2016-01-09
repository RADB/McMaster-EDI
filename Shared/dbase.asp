<%
' allows redirects to function properly
' causes pages not to cache
Response.Buffer = true
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"

'strSMTPServer = "smtp1.cis.mcmaster.ca"
strSMTPServer = "outgoingmail.phri.ca"

sub open_adodb(ByRef conn,strDB)
	set conn = Server.CreateObject("ADODB.Connection")
	
	select case strDB
		' database -- edi.mdb
		'case "EDI"
			'conn.open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=d:\websites\e-edica\data\edi.mdb;PWD=BS6464"
		'	conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\websites\e-edica\data\edi.mdb;Jet OLEDB:Database Password=BS6464;"
		'case "DATA"
			'conn.open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=d:\websites\e-edica\data\edi_data.mdb;PWD="
		'	conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\Websites\e-edica\data\edi_data.mdb;User Id=admin;Password="
		'case "TABLES"
			'conn.open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=d:\websites\e-edica\data\edi_tables.mdb;PWD="	
			'conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\Websites\e-edica\data\edi_tables.mdb;User Id=admin;Password="					
			'conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=d:\websites\e-edica\data\edi_tables.mdb;User Id=admin;Password="

        case "MACEDI"
		    conn.Open "Provider=SQLOLEDB;Data Source=RIOffordP1;Initial Catalog=EDI;User Id=macwebagent;Password=tr2003B$"
	end select
end sub

sub close_adodb(strName)
on error resume next
	strName.Close
	set strName = nothing
end sub

function checknull(strTemp)
	if isnull(strTemp) or len(strTemp) = 0 then 
		checknull = "null"	
	else
		checknull = "'" & replace(strTemp,"'","''") & "'"
	end if 
end function

function checknumber(intTemp)
	if isnull(intTemp) or intTemp = "-1" then 
		checknumber = "null"
	else
		checknumber = intTemp
	end if 
end function

function checkvalue(strTemp)
	if isnull(strTemp) or len(strTemp) = 0 or strTemp = "-1" then 
		checkvalue = "null"	
	elseif isdate(strTemp) then 
		checkvalue = "'" & strTemp & "'"	
	elseif strTemp = "on" then 
		checkvalue = "1"
	else
		checkvalue = "'" & replace(strTemp,"'","''") & "'"
	end if 
end function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Function Name:    makeReadable
'  Author:           Andrew Renner
'  Date:             June 4, 2001
'  Variables Passed: None
'  Pseudo Code:      1) Removes the jargon from db error messages
'
'  Revision List:    |Author           |Date              |Modifications
'                    |-----------------+------------------+---------------
'                    |Andrew Renner    |July 11, 2001     |Added Comments
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function makeReadable(strTest)
   ' replaces the jsrgon in the error string with ""
   strTest = Replace(strTest, "[Microsoft]", "")
   strTest = Replace(strTest, "[SQL Server]", "")
   strTest = Replace(strTest, "[ODBC SQL Server Driver]", "")
   strTest = Replace(strTest, "[ODBC Microsoft Access Driver]", "")
   strTest = Replace(strTest, "[ODBC Driver Manager]", "")
   strTest = Replace(strTest, "[ODBC]", "")
   strTest = Replace(strTest, "[Oracle]", "")
   strTest = Replace(strTest, "[Ora]", "")
      
   ' returns the clean error string
   makeReadable = strTest
End Function
	
%>
