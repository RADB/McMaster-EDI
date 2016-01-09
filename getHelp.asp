<!-- #include virtual="/shared/security.asp" -->
<%
' if the user has not logged in they will not be able to see the page
if blnSecurity then
	' open edi connection
	 call open_adodb(conn, "MACEDI")
	 set rstHelp = server.CreateObject("adodb.recordset")
	
	' find the record 	
	strSql = "SELECT " & session("Language") & " FROM Help WHERE ProvinceID = " & checkValue(request.querystring("ProvinceID")) & " AND Section = " & checkValue(request.querystring("Section")) & " AND Question = " & checkValue(request.querystring("Question"))
		'response.write strSql
	' open the recordset
	rstHelp.Open strSql, conn
	if not rstHelp.EOF then 
		strEnglish = rstHelp(session("Language"))
	end if 
	
	call close_adodb(rstHelp)
	call close_adodb(conn)
%>
	<div class="modal-header">
		<button type="button" class="close" data-dismiss="modal"><span aria-hidden="true">&times;</span><span class="sr-only">Close</span></button>   
		<h4 class="modal-title" id="myModalLabel"><img border="0" src="images\Help.png" alt="Help" name="Help" title="Help" height="40"/>  e-EDI Help</h4>
	</div>
	<div class="modal-body">
		<%  
		response.write strEnglish & "<br />"			
		%>
	</div>
	<div class="modal-footer">
		<button type="button" class="btn btn-default" data-dismiss="modal">Close</button>        
	</div>

<%
end if
%>