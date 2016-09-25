<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./data/webCal4_data.inc"-->
<!--#include file="./include/webCal4_verify.inc"-->
<!--#include file="./include/webCal4_constants.inc"-->
<!--#include file="./include/webCal4_settings.inc"-->
<!--#include file="./include/webCal4_functions.inc"-->
<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' Last updated 2/27/2001

dim m_strQuery				' query string passed to db
dim m_intGroupUserCount		' count of users in this group only
dim m_strGroupUsers			' list of member user ids
dim m_intGroupID			' group id
dim m_strGroupName			' group name
dim m_intGroupEventCount	' count of events in this group only
dim m_strGroupList			' drop-down list of other groups
dim m_strGroupEvents
dim m_arGroupPlural
dim m_strUserPlural			' say "user" or "users"
dim m_strUserList
dim m_strNonGroupUsers		' users to transfer events to
dim m_oConn					' connection object
dim m_oRS					' recordset object

m_intGroupID = Request.Form("fldGroupID")
Set m_oConn = Server.CreateObject("ADODB.Connection") : m_oConn.Open g_strDSN
Set m_oRS = Server.CreateObject("ADODB.RecordSet")

' get the group name
m_strQuery = "SELECT group_name FROM tblGroups WHERE (group_id)=" & m_intGroupID
m_oRS.Open m_strQuery, m_oConn, adOpenForwardOnly, adLockReadOnly, adCmdText
m_strGroupName = m_oRS("group_name")
m_oRS.Close

' get list of other groups
m_strQuery = "SELECT group_id, group_name FROM tblGroups WHERE (group_id)<>" & m_intGroupID
m_strGroupList = makeList(m_strQuery, "", m_oConn)

' get list of users who are members of ONLY this group
m_strQuery = "SELECT user_id FROM tblPermissions WHERE " _
	& "group_id=" & m_intGroupID _
	& " AND user_id NOT IN (SELECT user_id FROM tblPermissions WHERE " _
	& "group_id<>" & m_intGroupID & ")"
m_oRS.Open m_strQuery, m_oConn, adOpenForwardOnly, adLockReadOnly, adCmdText
if not m_oRS.EOF then
	m_strGroupUsers = m_oRS.GetString(2, , ",", ",")
	m_strGroupUsers = Left(m_strGroupUsers, Len(m_strGroupUsers) -1)
	m_intGroupUserCount = UBound(Split(m_strGroupUsers, ",")) + 1
else
	m_strGroupUsers = ""
	m_intGroupUserCount = 0
end if
m_oRS.Close

if m_intGroupUserCount > 0 then
	' create list of users who won't be erased as optional targets
	' for transferring members' events
	m_strQuery = "SELECT user_id, name_last + ', ' + name_first FROM tblUsers" _
		& " WHERE (user_id) NOT IN (" & m_strGroupUsers & ")" _
		& " ORDER BY name_last, name_first"
	m_strNonGroupUsers = makeList(m_strQuery, "", m_oConn)

	if m_intGroupUserCount > 1 then	m_strUserPlural = "s"
end if

' retrieve events that are scheduled ONLY in this group
' sub-query excludes events in other groups
m_strQuery = "SELECT event_id FROM tblEventGroupScopes WHERE " _
	& "group_id=" & m_intGroupID _
	& " AND event_id NOT IN (SELECT event_id FROM tblEventGroupScopes WHERE " _
	& "group_id<>" & m_intGroupID & ")"
m_oRS.Open m_strQuery, m_oConn, adOpenForwardOnly, adLockReadOnly, adCmdText
if not m_oRS.EOF then
	m_strGroupEvents = m_oRS.GetString(2, , ",", ",")
	m_strGroupEvents = Left(m_strGroupEvents, Len(m_strGroupEvents) -1)
	m_intGroupEventCount = UBound(Split(m_strGroupEvents, ",")) + 1
else
	m_strGroupEvents = ""
	m_intGroupEventCount = 0
end if
m_oRS.Close

if m_intGroupEventCount > 0 then
	' create group option list for any matching events
	m_strQuery = "SELECT group_id, group_name FROM tblGroups" _
		& " WHERE (group_id)<>" & m_intGroupID _
		& " ORDER BY group_name"
	m_strGroupList = makeList(m_strQuery, "", m_oConn)
	
	if m_intGroupEventCount > 1 then
		m_arGroupPlural = Array("s", "them")
	else
		m_arGroupPlural = Array("", "it")
	end if
end if

Set m_oRS = nothing
m_oConn.Close : Set m_oConn = nothing
%>
<html>
<head>
<script language="javascript" src="./script/webCal4_functions.js"></script>
</head>
<body onLoad="showMessage();" bgcolor="#<%=g_arColor(1)%>" link="#<%=g_arColor(7)%>" vlink="#<%=g_arColor(7)%>" alink="#<%=g_arColor(6)%>">
<center>

<!-- framing table -->
<table bgcolor="#<%=g_arColor(6)%>" border=0 cellpadding=2 cellspacing=0><tr><td>
<!-- end framing table -->

<table bgcolor="#<%=g_arColor(11)%>" border=0 cellpadding=3 cellspacing=0>
<form name="frmDelGroup" method="post">
<tr>
	<td bgcolor="#<%=g_arColor(4)%>" colspan=3>
	<font face="<%=g_arFont(0)%>" size=4>
	<b>Group Deletion</b></font></td>
<tr>
	<td colspan=3><font face="<%=g_arFont(0)%>" size=2>
	
<% if m_intGroupUserCount > 0 then %>
	
	<b>What should happen to the <%=m_intGroupUserCount & " user" & m_strUserPlural%> belonging only to the <%=m_strGroupName%> group?</b><br>
	<input type="radio" name="fldUserDo" value="delete">erase them and transfer ownership of their events to
	<select name="fldEventToUser"><%=m_strNonGroupUsers%></select><br>
	<input type="radio" name="fldUserDo" value="move">make them a member of the 
	<select name="fldUserToGroup"><%=m_strGroupList%></select> group
	<p>
	
<% end if %>

<% if m_intGroupEventCount > 0 then %>
	
	<b>What should happen to the <%=m_intGroupEventCount & " event" & m_arGroupPlural(0)%> scheduled only in the <%=m_strGroupName%> group?</b><br>
	<input type="radio" name="fldEventDo" value="delete">erase <%=m_arGroupPlural(1)%><br>
	<input type="radio" name="fldEventDo" value="move">transfer <%=m_arGroupPlural(1)%> to the 
	<select name="fldEventToGroup"><%=m_strGroupList%></select> group
	<p>
	
<% end if %>

Are you sure you want to erase the <%=m_strGroupName%> group?
	</font></td>
<tr>
	<td colspan=3 align="center" bgcolor="#<%=g_arColor(12)%>">
		<input type="button" value="Continue" onClick='goPage("webCal4_group-deleted.asp","frmDelGroup");'>
		<input type="button" value="Cancel" onClick='goPage("webCal4_admin.asp","frmDelGroup");'>
	</td>
<tr>
	<td align="center" colspan=3><font face="<%=g_arFont(0)%>" size=2>
	<b><font color="#cc0000">Caution</font>: erased groups cannot be restored</b></font></td>
</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->

<input type="hidden" name="fldGroupID" value="<%=m_intGroupID%>">
<input type="hidden" name="fldEventCount" value="<%=m_intGroupEventCount%>">
<input type="hidden" name="fldView" value="<%=Request.QueryString("view")%>">
<input type="hidden" name="fldUsers" value="<%=m_strGroupUsers%>">
<input type="hidden" name="fldEvents" value="<%=m_strGroupEvents%>">
</form>
</center>
</body>
</html>