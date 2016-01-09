<%@EnableSessionState=False%>
<% Response.Expires = -1 %>

<html>
<head>
<title>uploading files</title>
<style type='text/css'>td {font-family:arial; font-size: 9pt }</style>
</head>

<% If Request("b") = "IE" Then %> <!-- Internet Explorer -->
<body bgcolor="#c0c0b0">
<iframe src="edi_upload_bar.asp?PID=<%= Request("PID") & "&to=" & Request("to") %>" title="Upload Progress" noresize scrolling=no frameborder=0 framespacing=10 width=369 height=65></iframe>
<table border="0" width="100%" cellpadding="2" cellspacing="0">
  <tr><td align="center">
     To cancel uploading, press your browser's <B>STOP</B> button.
  </td></tr>
</table>
</body>

<%Else%> <!-- Netscape Navigator etc ... -->

<frameset rows="65%, 35%" cols="100%" border="0" framespacing="0" frameborder="no">
<frame src="edi_upload_bar.asp?Pid=<%= request("pid") & "&to=" & request("to") %>" noresize scrolling="no" frameborder="no" name="sp_body">
<frame src="edi_upload_note.htm" noresize scrolling="no" frameborder="no" name="sp_note">
</frameset>

<%End If%>

</html>
