<%
' Copyright 1999 Jason Abbott (jason@webott.com)
' Last updated 04/26/1999

' if the cancel button was hit then return the user
' to the event detail page

if Request.Form("cancel") = "No" then
	response.redirect Request.Form("url")
end if
%>
<!--#include file="data/webCal3_data.inc"-->
<%
' otherwise prepare to delete some event dates
' figure out which event dates need to be purged

dim query

Select Case Request.Form("scope")
	Case "one"

' if deleting only one occurrence then erase only
' a single day, leaving event info intact

		query = " AND event_date BETWEEN " & strDelim _
			& Request.Form("date") & " 12:00:00 AM" & strDelim _
			& " AND " & strDelim _
			& Request.Form("date") & " 11:59:59 PM" & strDelim
	Case "future"

' if deleting all future events then erase today
' and all after today, leaving event info intact

		query = " AND event_date >= " & strDelim _
			& Request.Form("date") & strDelim
	Case Else

' if erasing all occurrences then delete not only the dates
' but the event information itself

			db.Execute "DELETE FROM cal_events WHERE (event_id)=" _
				& Request.Form("event_id"),,&H0001
end Select

' put the query together

query = "DELETE FROM cal_dates" _
	& " WHERE event_id=" & Request.Form("event_id") _
	& query

' and run it

' &H0001 (hex 1), which is adCmdText, tells the
' connection object that we're sending a text command,
' which is speedier

'response.write query
db.Execute query,,&H0001
db.Close
Set db = nothing

' send the user back to the calendar

response.redirect "webCal3_" & Request.Form("view") _
	& ".asp?date=" & Request.Form("start_date")
%>