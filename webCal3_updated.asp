<%
' Copyright 1999 Jason Abbott (jason@webott.com)
' updated 06/04/1999

' if cancel is hit then send back to event detail page
' or calendar view

if Request.Form("cancel") = "Cancel" then
	if Request.Form("url") <> "" then
		response.redirect Request.Form("url")
	else
		response.redirect "webCal3_" & Request.Form("view") _
			& ".asp?date=" & Request.Form("start_date")
	end if
end if

' otherwise begin populating variables

dim startTime, endTime, eventDate, dateList(), count, eventHide
dim eventTitle, eventDescription, eventID, queryDate, queryEvent

eventDate = DateValue(Request.Form("start_date"))
if Request.Form("notime") <> "on" then
	startTime = "'" & TimeValue(Request.Form("start_hour") & ":" _
		& Request.Form("start_min")) & "'"
	endTime = "'" & TimeValue(Request.Form("end_hour") & ":" _
		& Request.Form("end_min")) & "'"
else
	startTime = "null"
	endTime = "null"
end if

' normalize some values

eventTitle = replace(Request.Form("title"), "'", "''")
eventDescription = replace(Request.Form("description"), "'", "''")

if Request.Form("private") = "on" then
	eventHide = 1
else
	eventHide = 0
end if

%>
<!--#include file="data/webCal3_data.inc"-->
<%
'------------------------------------
' clear old dates out of table in preparation for new ones
'------------------------------------

if Request.Form("edit_type") <> "new" then
	Select Case Request.Form("edit_type")
		Case "one"
		
' erase single date

			queryDate = " AND event_date BETWEEN " _
				& strDelim & eventDate & strDelim & " AND " _
				& strDelim & eventDate & strDelim
		Case "future"

' erase current and all future dates

			queryDate = " AND event_date >= " & strDelim _
				& eventDate & strDelim
		Case "all"

' erase all event dates without limitation

			queryDate = ""
	end Select
	queryDate = "DELETE FROM cal_dates" _
		& " WHERE event_id=" & Request.Form("event_id") _
		& queryDate	

' &H0001 (hex 1), which is adCmdText, tells the
' connection object that we're sending a text command,
' which is spedier

	db.Execute queryDate,,&H0001
end if

'------------------------------------
' update event information as needed
'------------------------------------
' only update the event if all occurrences of that event
' were selected for modification, otherwise create new
' events

if Request.Form("edit_type") = "all" then

' update existing event

	queryEvent = "UPDATE cal_events SET " _
		& "event_title = '" & eventTitle & "', " _
		& "event_description = '" & eventDescription & "', " _
		& "private = " & eventHide & ", " _
		& "event_recur = '" & Request.Form("event_recur") & "', " _
		& "event_color = '" & Request.Form("event_color") & "', " _
		& "time_start = " & startTime & ", " _
		& "time_end = " & endTime _
		& " WHERE (event_id)=" & Request.Form("event_id")

' 0001 is the hex value for adCmdText which tells the connection
' object that we're sending a text command, which is speedier
' the middle value, left blank, is the variable for the record
' count, which Access doesn't support

	db.Execute queryEvent,,&H0001
	eventID = Request.Form("event_id")
else

' add new event

	queryEvent = "INSERT INTO cal_events (" _
		& "event_title, event_description, " _
		& "event_recur, event_color, private, user_id"

' Access doesn't like getting blank dates, so just
' ignore the date fields if they're empty

	if startTime <> "" then
		queryEvent = queryEvent & ", time_start, time_end"
	end if
	
	queryEvent = queryEvent & ") VALUES ('" _
		& eventTitle & "', '" _
		& eventDescription & "', '" _
		& Request.Form("event_recur") & "', '" _
		& Request.Form("event_color") & "', " _
		& eventHide & ", '" _
		& Session(dataName & "User") & "'"

' again, ignore date fields if empty

	if startTime <> "null" then
		queryEvent = queryEvent & ", " & startTime _
			& ", " & endTime
	end if

	queryEvent = queryEvent & ")"
	db.Execute queryEvent,,&H0001
	
' event dates are keyed to event info by the event ID,
' so find out what auto-id was assigned
	
	queryEvent = "SELECT event_id, event_title FROM cal_events " _
		& "WHERE event_title='" & eventTitle _
		& "' ORDER BY event_id DESC"
	Set rs = db.Execute(queryEvent,,&H0001)
	eventID = rs("event_id")
	rs.Close
	Set rs = nothing
end if

' with event info updated, now update event date(s)--
' generate recurring dates if necessary, placing them
' in the dates array

count = 0
if Request.Form("event_recur") <> "none" then
	Select Case Request.Form("event_recur")
		Case "daily"
			addType = "d"
			addNum = 1
		Case "weekly"
			addType = "d"
			addNum = 7
		Case "2weeks"
			addType = "d"
			addNum = 14
		Case "monthly"
			addType = "m"
			addNum = 1
		Case "yearly"
			addType = "yyyy"
			addNum = 1
	end Select		

' populate the array with dates, according to the above
' addition, until the end date for the event

	While DateDiff("d", eventDate, Request.Form("end_date")) >= 0
		if Request.Form("skip") <> "on" _
			OR (WeekDay(eventDate) > 1 _
			AND WeekDay(eventDate) < 7) then

			ReDim Preserve dateList(count)
			dateList(count) = eventDate
			count = count + 1
		end if
		eventDate = DateAdd(addType, addNum, eventDate)
	Wend

' if there was no recurrence selected then put the single
' date into the array

else
	ReDim Preserve dateList(count)
	dateList(count) = eventDate
end if

' now go through everything inserted into the dates array
' and insert it into the event dates table

for each d in dateList
	queryDate = "INSERT INTO cal_dates (" _
		& "event_id, event_date) VALUES ('" _
		& eventID & "', '" & d & "')"
	db.Execute queryDate,,&H0001
next

db.Close
Set db = nothing

' with the data updated send user back to calendar
' or to the edit page again, if requested

if Request.Form("save") = "Save" then
	response.redirect "webCal3_" & Request.Form("view") _
		& ".asp?date=" & Request.Form("start_date")
elseif Request.Form("saveadd") = "Save & Add Another" then
	response.redirect "webCal3_edit.asp?date=" _
		& Request.Form("start_date") _
		& "&view=" & Request.Form("view")
end if
%>