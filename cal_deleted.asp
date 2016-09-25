<%
' Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
' Last updated 6/3/98

if Session("user") = "guest" then response.redirect "/error.asp"

' delete event from database

Set db = Server.CreateObject("ADODB.Connection")
db.Open "bc"
query = "DELETE FROM cal_events" _
	& " WHERE (event_id)=" & Request.Form("event_id")
	
Set rs = db.Execute(query)

'uncomment this line to debug SQL:
'response.write query

' send user back to calendar

response.redirect "cal.asp" _
	& "?event_context=" & Request.Form("event_context") _
	& "&year=" & Request.Form("year") _
	& "&month=" & Request.Form("month")

rs.Close
db.Close
%>