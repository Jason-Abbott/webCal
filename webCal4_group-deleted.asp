<% Option Explicit %>
<!--#include file="./data/webCal4_data.inc"-->
<!--#include file="./include/webCal4_constants.inc"-->
<!--#include file="./include/webCal4_functions.inc"-->
<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' Last updated 2/27/2001

dim m_strQuery
dim m_intGroupID
dim m_arEvents
dim m_arUsers
dim m_strMessage	' feedback message
dim m_strEventList
dim m_strUserList
dim m_oConn			' connection object
dim x				' loop counter

On Error Resume Next

m_strEventList = Trim(Request.Form("fldEvents"))
m_strUserList = Trim(Request.Form("fldUsers"))

m_intGroupID = Request.Form("fldGroupID")
Set m_oConn = Server.CreateObject("ADODB.Connection") : m_oConn.Open g_strDSN
m_oConn.BeginTrans

' erase group's permission settings
m_strQuery = "DELETE FROM tblPermissions WHERE group_id=" & m_intGroupID
m_oConn.Execute m_strQuery,,adCmdText + adExecuteNoRecords

' erase group
m_strQuery = "DELETE FROM tblGroups WHERE group_id=" & m_intGroupID
m_oConn.Execute m_strQuery,,adCmdText + adExecuteNoRecords

' handle events assigned only to this group
if m_strEventList <> "" then
	m_arEvents = Split(Request.Form("fldEvents"), ",")
	if Request.Form("fldEventDo") = "delete" then
		m_strQuery = "DELETE FROM tblEventGroupScopes WHERE event_id IN (" & m_strEventList & ")"
		m_oConn.Execute m_strQuery,,adCmdText + adExecuteNoRecords
		
		m_strQuery = "DELETE FROM tblEventDates WHERE event_id IN (" & m_strEventList & ")"
		m_oConn.Execute m_strQuery,,adCmdText + adExecuteNoRecords

		m_strQuery = "DELETE FROM tblEvents WHERE (event_id) IN (" & m_strEventList & ")"
		m_oConn.Execute m_strQuery,,adCmdText + adExecuteNoRecords

	
		' delete events, event dates and event views
'		for x = 0 to UBound(m_arEvents)
'			m_strQuery = "DELETE FROM tblEventGroupScopes WHERE event_id=" & m_arEvents(x)
'			m_oConn.Execute m_strQuery,,adCmdText + adExecuteNoRecords
'
'			m_strQuery = "DELETE FROM tblEventDates WHERE event_id=" & m_arEvents(x)
'			m_oConn.Execute m_strQuery,,adCmdText + adExecuteNoRecords
'
'			m_strQuery = "DELETE FROM tblEvents WHERE (event_id)=" & m_arEvents(x)
'			m_oConn.Execute m_strQuery,,adCmdText + adExecuteNoRecords
'		next
	else
		' move events to a new group
		m_strQuery = "UPDATE tblEventGroupScopes SET " _
			& "group_id=" & Request.Form("fldEventToGroup") _
			& " WHERE group_id=" & m_intGroupID _
			& " AND event_id IN (" & m_strEventList & ")"
		m_oConn.Execute m_strQuery,,adCmdText + adExecuteNoRecords
		
'		for x = 0 to UBound(m_arEvents)
'			m_strQuery = "UPDATE tblEventGroupScopes SET " _
'				& "group_id=" & Request.Form("event_move") _
'				& " WHERE group_id=" & m_intGroupID _
'				& " AND event_id=" & m_arEvents(x)
'			m_oConn.Execute m_strQuery,,adCmdText
'		next
	end if
end if

' handle users belonging only to this group
if m_strUserList <> "" then
	m_arUsers = Split(Request.Form("users"), ",")
	if Request.Form("fldUserDo") = "delete" then
		' delete users and user permissions
		m_strQuery = "DELETE FROM tblPermissions WHERE " _
			& "user_id IN (" & m_strUserList & ")"
		m_oConn.Execute m_strQuery,,adCmdText + adExecuteNoRecords
		
		m_strQuery = "DELETE FROM tblUsers WHERE " _
			& "(user_id) IN (" & m_strUserList & ")"
		m_oConn.Execute m_strQuery,,adCmdText + adExecuteNoRecords
		
		' transfer users' events to other users
		m_strQuery = "UPDATE tblEvents SET " _
			& "user_id=" & Request.Form("fldEventToUser") _
			& " WHERE user_id IN (" & m_strUserList & ")"
		m_oConn.Execute m_strQuery,,adCmdText + adExecuteNoRecords
		
'		for x = 0 to UBound(m_arUsers)
'			m_strQuery = "DELETE FROM tblUsers WHERE " _
'				& "(user_id)=" & m_arUsers(x)
'			m_oConn.Execute m_strQuery,,adCmdText
'			
'			m_strQuery = "DELETE FROM tblPermissions WHERE " _
'				& "user_id=" & m_arUsers(x)
'			m_oConn.Execute m_strQuery,,adCmdText
'			
'			' transfer users' events to other users
'			m_strQuery = "UPDATE tblEvents SET " _
'				& "user_id=" & Request.Form("user_del") _
'				& " WHERE user_id=" & m_arUsers(x)
'			m_oConn.Execute m_strQuery,,adCmdText
'		next
	else
		' move users to a new group
		m_strQuery = "UPDATE tblPermissions SET " _
			& "group_id=" & Request.Form("fldToGroup") _
			& " WHERE group_id=" & m_intGroupID _
			& " AND user_id IN (" & m_strUserList & ")"
		m_oConn.Execute m_strQuery,,adCmdText + adExecuteNoRecords
'		for x = 0 to UBound(m_arUsers)
'			m_strQuery = "UPDATE tblPermissions SET " _
'				& "group_id=" & Request.Form("fldToGroup") _
'				& " WHERE group_id=" & m_intGroupID _
'				& " AND user_id IN (" & m_strUserList & ")"
'			m_oConn.Execute m_strQuery,,adCmdText + adExecuteNoRecords
'		next
	end if
end if

Call HandleErrors(m_oConn, "deleted", "group-delete", "The group [name]", _
	"deleting the group", Request.Form("fldView"), Request.Form("fldStartDate"))
%>