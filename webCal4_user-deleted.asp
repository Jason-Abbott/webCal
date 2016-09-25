<% Option Explicit %>
<!--#include file="./data/webCal4_data.inc"-->
<!--#include file="./include/webCal4_constants.inc"-->
<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' Last updated 3/24/2000

dim strQuery		' query passed to db
dim strMessage		' feedback message
dim intUserID
dim oConn			' connection object
dim x				' loop counter

On Error Resume Next

intUserID = Request.Form("user_id")
Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open g_strDSN
oConn.BeginTrans

' erase user
strQuery = "DELETE FROM tblUsers WHERE user_id=" & intUserID
oConn.Execute strQuery,,adCmdText

' erase user's permission settings
strQuery = "DELETE FROM tblPermissions WHERE user_id=" & intUserID
oConn.Execute strQuery,,adCmdText

' then deal with user's events, if any
if Request.Form("event_count") > 0 then
	select case Request.Form("action")
		case "delete"
			' delete all events created by this user
			strQuery = "DELETE FROM tblEvents " _
				& "WHERE user_id=" & intUserID
		case "some"
			' first delete all private events
			strQuery = "DELETE FROM tblEvents " _
				& "WHERE user_id=" & intUserID _
				& " AND private=" & True
			oConn.Execute strQuery,,adCmdText
			' then move the remaining public events to new user
			strQuery = "UPDATE tblEvents SET " _
				& "user_id = " & Request.Form("recipient") _
				& " WHERE user_id=" & intUserID
		case "move"
			' move all events to new user
			strQuery = "UPDATE tblEvents SET " _
				& "user_id = " & Request.Form("recipient") _
				& " WHERE user_id=" & intUserID
	end select
	oConn.Execute strQuery,,adCmdText
	strMessage = strMessage & " and their events erased"
end if

' Error handling----------------------------------------------------------
if oConn.Errors.Count = 0 AND Err.Number = 0 then
	oConn.CommitTrans
	strMessage = "User successfully deleted"
else
	oConn.RollbackTrans
	if oConn.Errors.Count > 0 then
		strMessage = "<font color='#bb0000'>Error: </font>"
		for x = 0 to oConn.Errors.Count - 1
			strMessage = strMessage & oConn.Errors(x).Description & "<br>"
		next
	end if
	if Err.Number > 0 then
		' this will only return the most recent error
		strMessage = strMessage & "<font color='#bb0000'>" _
			& Err.Source & " error " & Err.Number & "</font>: " _
			& Err.Description
	end if
	strMessage = strMessage & " <font color='#bb0000'>" _
		& "Use your browser's back button to return to the form.</font>"
end if

oConn.Close
Set oConn = nothing
response.redirect "webCal4_admin.asp?view=" _
	& Request.Form("view") & "&message=" _
	& Server.URLEncode(strMessage)
%>