<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' updated 6/3/98

Set db = Server.CreateObject("ADODB.Connection")
db.Open "bc"

if Request.Form("private") = "" then
	restrict = 0
else
	restrict = 1
end if

query = "INSERT INTO cal_context (name, private) VALUES ('" _
	& Request.Form("name") & "', '" _
	& restrict & "')"

Set rs = db.Execute(query)

' send user to calendar

response.redirect "cal.asp"
%>

<!-- used only to debug SQL -->

<html>
<body bgcolor="#FFFFFF">
<%= query %>
</body>
</html>