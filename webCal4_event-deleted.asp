<% Option Explicit %>
<% Response.Buffer = true %>
<!--#include file="./data/webCal4_data.inc"-->
<!--#include file="./include/webCal4_constants.inc"-->
<!--#include file="./include/webCal4_functions.inc"-->
<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' Last updated 2/23/01

dim m_strQuery		' query string passed to database
dim m_intEventID	' event id
dim m_oConn			' database connection object
dim m_strDate
dim m_strMessage	' status message
dim x				' loop counter for error collection

On Error Resume Next

m_strDate = sqlDate(Request.Form("fldDate"))
m_intEventID = Request.Form("fldEventID")

Set m_oConn = Server.CreateObject("ADODB.Connection")
m_oConn.Open g_strDSN
m_oConn.BeginTrans

Select Case Request.Form("fldTimeScope")
	Case "one"
		' if deleting only one occurrence then erase only
		' a single day, leaving event info intact
		m_strQuery = " AND event_date BETWEEN " & g_strDelim _
			& m_strDate & " 12:00:00 AM" & g_strDelim & " AND " & g_strDelim _
			& m_strDate & " 11:59:59 PM" & g_strDelim
	Case "future"
		' if deleting all future events then erase today
		' and all after today, leaving event info intact
		m_strQuery = " AND event_date >= " & g_strDelim _
			& m_strDate & g_strDelim
	Case Else
		' if erasing all occurrences then delete not only the dates
		' but the event information itself
		m_oConn.Execute "DELETE FROM tblEvents WHERE (event_id)=" _
			& m_intEventID,,adCmdText + adExecuteNoRecords
		m_oConn.Execute "DELETE FROM tblEventGroupScopes WHERE event_id=" _
			& m_intEventID,,adCmdText + adExecuteNoRecords
End Select

' put the query together
m_strQuery = "DELETE FROM tblEventDates WHERE event_id=" _
	& m_intEventID & m_strQuery
m_oConn.Execute m_strQuery,,adCmdText + adExecuteNoRecords

' Error handling----------------------------------------------------------
if m_oConn.Errors.Count = 0 AND Err.Number = 0 then
	m_oConn.CommitTrans
	m_oConn.Close
	Set m_oConn = nothing

	' send the user back to the calendar
	response.redirect "webCal4_" & Request.Form("fldView") _
		& ".asp?date=" & Request.Form("fldStartDate")
else
	m_oConn.RollbackTrans
	m_strMessage = "An error was encountered while deleting your event\n\n"
	if m_oConn.Errors.Count > 0 then
		for x = 0 to m_oConn.Errors.Count - 1
			m_strMessage = m_strMessage & m_oConn.Errors(x).Description & "\n"
		next
	end if
	if Err.Number <> 0 then
		' this will only return the most recent error
		m_strMessage = m_strMessage & Err.Source & " " & Err.Number _
			& "\n  " & Err.Description
	end if
	Set m_oConn = nothing
end if

' if we haven't been redirected then raise error
Call RaiseError(m_strMessage)
%>
