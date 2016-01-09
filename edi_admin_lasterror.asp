<%
 dim objErr
 set objErr=Server.GetLastError()

 response.write("ASPCode=" & objErr.ASPCode)
 response.write("<br>")
 response.write("ASPDescription=" & objErr.ASPDescription)
 response.write("<br>")
 response.write("Category=" & objErr.Category)
 response.write("<br>")
 response.write("Column=" & objErr.Column)
 response.write("<br>")
 response.write("Description=" & objErr.Description)
 response.write("<br>")
 response.write("File=" & objErr.File)
 response.write("<br>")
 response.write("Line=" & objErr.Line)
 response.write("<br>")
 response.write("Number=" & objErr.Number)
 response.write("<br>")
 response.write("Source=" & objErr.Source)
 
 response.write "mpd = " & 799 mod 400
 %> 