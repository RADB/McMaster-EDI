<%@EnableSessionState=False%>

<%
  Response.Expires = -1
  PID = Request("PID")
  TimeO = Request("to")


  Set UploadProgress = Server.CreateObject("Persits.UploadProgress")

  format = "%TUploading files...%t%B3%T%R left (at %S/sec) %r%U/%V(%P)%l%t"

  bar_content = UploadProgress.FormatProgress(PID, TimeO, "#006600", format)
'00007F - old colour
  If "" = bar_content Then
%>
<html>
<head>
<title>upload finished</title>
<script language="Javascript">
function CloseMe()
{
	window.parent.close();
	return true;
}
</script>
</head>
<body OnLoad="CloseMe()">
</body>
</html>
<%
  Else    ' Not finished yet
%>
<html>
<head>

<!--%  If left(bar_content, 1) <> "." Then %-->

<meta HTTP-EQUIV="Refresh" CONTENT="1;URL=<%
 Response.Write Request.ServerVariables("URL")
 Response.Write "?to=" & TimeO & "&PID=" & PID %>">

<!--% End If %-->

<title>Uploading Files...</title>

<style type='text/css'>td {font-family:arial; font-size: 9pt } td.spread {font-size: 6pt; line-height:6pt } td.brick {font-size:6pt; height:12px}</style>

</head>
<body bgcolor="#c0c0b0" topmargin=0>
<% = bar_content %>
</body>
</html>

<% End If %>