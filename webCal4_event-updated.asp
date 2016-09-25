<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./data/webCal4_data.inc"-->
<!--#include file="./include/webCal4_constants.inc"-->
<!--#include file="./include/webCal4_settings.inc"-->
<!--#include file="./include/webCal4_functions.inc"-->
<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' updated 3/7/2001

dim m_strMessage		' feedback
dim m_strDate			' event date
dim m_arDates()			' recurrence dates
dim m_intEventID		' event id
dim m_strStartTime		' event start
dim m_strEndTime		' event end
dim m_strTitle			' event title
dim m_strDescription	' event description
dim m_strQuery			' query passed to database
dim m_oConn				' ADODB connection object
dim m_oRS				' ADODB recordset object
dim m_strTemp			' hold query for loop
dim m_strUserScopes		' list of group scopes passed in form
dim m_arUserScopes		' split into array
dim m_intCount			'
dim m_strAdd			' duration to add for recurrence
dim m_intAdd			' number of durations
dim x					' loop counter

Const m_GROUP_ID = 0
Const m_SCOPE_ID = 1

On Error Resume Next

m_strDate = DateValue(Request.Form("fldStartDate"))
m_intEventID = Request.Form("fldEventID")
m_strTitle = Replace(Request.Form("fldTitle"), "'", "''")
m_strDescription = Replace(Request.Form("fldDescription"), "'", "''")

' JavaScript generated scope string
' format (group id|scope id,[repeat])
m_strUserScopes = Request.Form("fldUserScopes")
if m_strUserScopes <> "" then m_arUserScopes = ListToArray(m_strUserScopes, ",", "|")

if Request.Form("fldNoTime") <> "on" then
	m_strStartTime = TimeValue(Request.Form("fldStartHour") & ":" _
		& Request.Form("fldStartMin"))
	m_strEndTime = TimeValue(Request.Form("fldEndHour") & ":" _
		& Request.Form("fldEndMin"))
else
	m_strStartTime = ""
	m_strEndTime = ""
end if

Set m_oRS = Server.CreateObject("ADODB.Recordset")
Set m_oConn	= Server.CreateObject("ADODB.Connection")
m_oConn.Open g_strDSN : m_oConn.BeginTrans

' clear old values out of tblEventDates in preparation for new ones
if Request.Form("fldEditType") <> "new" then
	Select Case Request.Form("fldEditType")
		Case "one"
			' erase single date
			m_strQuery = " AND event_date BETWEEN " & strDelim _
				& sqlDate(m_strDate) & strDelim & " AND " & strDelim _
				& sqlDate(DateAdd("d", 1, m_strDate)) & strDelim
		Case "future"
			' erase current and all future dates
			m_strQuery = " AND event_date >= " & strDelim _
				& sqlDate(m_strDate) & strDelim
		Case "all"
			' erase all event dates without limitation
			m_strQuery = ""
	end Select
	m_strQuery = "DELETE FROM tblEventDates" _
		& " WHERE event_id=" & m_intEventID & m_strQuery	
	m_oConn.Execute m_strQuery,,adCmdText + adExecuteNoRecords		
end if

' update tblEvents
' only update values if all occurrences of that event
' were selected for modification, otherwise create new entries
if Request.Form("fldEditType") = "all" then
	' update existing event
	m_strMessage = "updated"
	m_strQuery = "UPDATE tblEvents SET " _
		& "event_title = '" & m_strTitle & "', " _
		& "event_description = '" & m_strDescription & "', " _
		& "event_recur = '" & Request.Form("fldEventRecur") & "', " _
		& "event_color = '" & Request.Form("fldEventColor")

	' skip blank dates
	if m_strStartTime <> "" then
		m_strQuery = m_strQuery _
			& "', time_start = '" & m_strStartTime & "', " _
			& "time_end = '" & m_strEndTime
	end if
	
	m_strQuery = m_strQuery & "' WHERE (event_id)=" & m_intEventID
	m_oConn.Execute m_strQuery,,adCmdText + adExecuteNoRecords
	
	' remove old event-group relationships from tblEventGroupScopes
	m_strQuery = "DELETE FROM tblEventGroupScopes" _
		& " WHERE event_id=" & m_intEventID
	m_oConn.Execute m_strQuery,,adCmdText + adExecuteNoRecords
else
	' add new event
	m_strMessage = "added"
	' use ADO methods so we can immediately retrieve new event id
	m_oRS.Open "tblEvents", m_oConn, adOpenStatic, adLockOptimistic, adCmdTable
	m_oRS.AddNew
	m_oRS.Fields("event_title") = m_strTitle
	m_oRS.Fields("event_description") = m_strDescription
	m_oRS.Fields("user_id") = Session(g_unique & "UserID")
	m_oRS.Fields("event_recur") = Request.Form("event_recur")
	m_oRS.Fields("event_color") = Request.Form("event_color")
	if m_strStartTime <> "" then
		m_oRS.Fields("time_start") = m_strStartTime
		m_oRS.Fields("time_end") = m_strEndTime
	end if
	m_oRS.Update
	m_intEventID = m_oRS("event_id")
	m_oRS.Close
end if

m_oRS.CursorLocation = adUseClient	' allows batch updates
m_oRS.Open "tblEventGroupScopes", m_oConn, adOpenStatic, adLockBatchOptimistic, adCmdTable
for x = 0 to UBound(m_arUserScopes, 2)
	m_oRS.AddNew
	m_oRS.Fields("event_id") = m_intEventID
	m_oRS.Fields("group_id") = m_arUserScopes(m_GROUP_ID, x)
	m_oRS.Fields("scope_id") = m_arUserScopes(m_SCOPE_ID, x)
next
m_oRS.UpdateBatch
m_oRS.Close

' update tblEventDates as needed----------------------------------------------
' generate recurring dates if necessary
m_intCount = 0
if Request.Form("fldEventRecur") <> "none" then
	Select Case Request.Form("fldEventRecur")
		Case "daily"
			m_strAdd = "d"
			m_intAdd = 1
		Case "weekly"
			m_strAdd = "d"
			m_intAdd = 7
		Case "2weeks"
			m_strAdd = "d"
			m_intAdd = 14
		Case "monthly"
			m_strAdd = "m"
			m_intAdd = 1
		Case "yearly"
			m_strAdd = "yyyy"
			m_intAdd = 1
	end Select		

	' populate the array with dates, according to the above
	' addition, until the end date for the event
	While DateDiff("d", m_strDate, Request.Form("fldEndDate")) >= 0
		if Request.Form("fldSkipWE") <> "on" _
			OR (WeekDay(m_strDate) > 1 _
			AND WeekDay(m_strDate) < 7) then

			ReDim Preserve m_arDates(m_intCount)
			m_arDates(m_intCount) = m_strDate
			m_intCount = m_intCount + 1
		end if
		m_strDate = DateAdd(m_strAdd, m_intAdd, m_strDate)
	Wend
else
	' if there was no recurrence selected then put the single
	' date into the array
	ReDim Preserve m_arDates(m_intCount)
	m_arDates(m_intCount) = m_strDate
end if

' insert dates array into dates table
m_oRS.CursorLocation = adUseClient	' allows batch updates
m_oRS.Open "tblEventDates", m_oConn, adOpenStatic, adLockBatchOptimistic, adCmdTable
for each x in m_arDates
	m_oRS.AddNew
	m_oRS.Fields("event_id") = m_intEventID
	m_oRS.Fields("event_date") = sqlDate(x)
next
m_oRS.UpdateBatch
m_oRS.Close
Set m_oRS = nothing

updateCache(m_arDates)

Call HandleErrors(m_oConn, m_strMessage, "event-edit", "Your event was successfully", _
	"saving your event", Request.Form("fldView"), Request.Form("fldStartDate"))
%>