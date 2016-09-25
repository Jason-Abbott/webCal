<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' Last updated 02/25/2000

' get the info on the group to be deleted

'dim strStatus		' text to display in status bar
dim intUsers		' count of users in this group
dim intEvents		' count of events in this group
dim strEvents		' list of events in this group
dim strGroups		' option list of groups
dim strUsersOnly	' users who are members of ONLY this group
dim strUsersNot		' users who are not members of this group
dim strUserSay		' plural or singular
dim strEventSay1	' ibid
dim strEventSay2	' ibid

strQuery = "SELECT * FROM tblGroups WHERE " _
	& "(group_id)=" & intGroupID
oRS.Open strQuery, oConn, adOpenForwardOnly, adLockReadOnly, adCmdText
strName = oRS("group_name")
oRS.Close

'function listGroups()
'	strStatus = "onMouseOver=""status='" & msg & "'; " _
'		& "return true;"" " & VbCrLf _
'		& "onMouseOut=""status=''; return true;"""
'	strQuery = "SELECT group_id, group_name FROM tblGroups" _
'		& " WHERE (group_id)<>" & intGroupID _
'		& " ORDER BY group_name"
'	oRS.Open strQuery, oConn, adOpenForwardOnly, adLockReadOnly, adCmdText
'	do while not oRS.EOF
'		strGroups = strGroups & "<option value='" & oRSs("group_id") _
'			& "'>" & oRSs("group_name") & VbCrLf
'		oRSs.MoveNext
'	loop
'	oRSs.Close
'	Set oRSs = nothing
'end function

' USERS-------------------------------------------------------------------
' count the users who are members of ONLY this group
' the sub-query provides a list of all users who
' are members of other groups--we don't want to
' count them

strQuery = "SELECT user_id, group_id FROM tblPermissions WHERE " _
	& "group_id=" & intGroupID _
	& " AND user_id NOT IN (SELECT user_id FROM tblPermissions WHERE " _
	& "group_id<>" & intGroupID & ")"
	
' cursor:	adOpenForwardOnly = 0
'			adOpenKeyset = 1
'			adOpenDynamic = 2
'			adOpenStatic = 3
' locking:	adLockReadOnly = 1
' command:	adCmdText = &H0001
	
oRS.Open strQuery, oConn, 3, 1, &H0001
intUsers = oRS.Recordcount
if intUsers > 0 then
	' create a string list of users that we can pass through form
	do while not oRS.EOF
		strUsersOnly = strUsers & "," & oRS("user_id")
		oRS.MoveNext
	loop
	oRS.Close
	strUsersOnly = Right(strUsersOnly, Len(strUsersOnly)-1)

	' create list of users who won't be erased as optional targets
	' for transferring members' events
	strQuery = "SELECT user_id, name_first, name_last FROM tblUsers" _
		& " WHERE (user_id) NOT IN (" & strUsersOnly & ")" _
		& " ORDER BY name_last, name_first"
	oRS.Open strQuery, oConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	do while not oRS.EOF
		strUsersNot = strUsersNot & "<option value='" & oRS("user_id") _
			& "'>" & oRS("name_first") & VbCrLf
		oRS.MoveNext
	loop

	' create a list of groups that members can be transferred to
	if intUsers > 1 then
		strUserSay = "users"
	else
		strUserSay = "user"
	end if
else
	' no users are members of just this group
	strUsersOnly = ""
end if
oRS.Close

' EVENTS------------------------------------------------------------------
' now count the events that are ONLY in this group
' the sub-query provides a list of all events that
' belong to other groups

strQuery = "SELECT group_id, event_id FROM tblEventGroupScopes WHERE " _
	& "group_id=" & intGroupID _
	& " AND event_id NOT IN (SELECT event_id FROM tblEventGroupScopes WHERE " _
	& "group_id<>" & intGroupID & ")"
oRS.Open strQuery, oConn, 3, 1, &H0001
intEvents = oRS.RecordCount
if intEvents > 0 then
	' create a string list of events that we can pass through form
	do while not oRS.EOF
		strEvents = strEvents & "," & oRS("event_id")
		oRS.MoveNext
	loop
	strEvents = Right(strEvents, Len(strEvents)-1)
	oRS.Close

	' create group option list for any matching events
	strQuery = "SELECT group_id, group_name FROM tblGroups" _
		& " WHERE (group_id)<>" & intGroupID _
		& " ORDER BY group_name"
	oRS.Open strQuery, oConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	do while not oRS.EOF
		strGroups = strGroups & "<option value='" & oRS("group_id") _
			& "'>" & oRS("group_name") & VbCrLf
		oRS.MoveNext
	loop
	
	if intEvents > 1 then
		strEventSay1 = "events"
		strEventSay2 = "them"
	else
		strEventSay1 = "event"
		strEventSay2 = "it"
	end if
else
	' no events in this group only
	strEvents = ""
end if
oRS.Close
%>

</head>
<body bgcolor="#<%=g_arColor(1)%>" link="#<%=g_arColor(7)%>" vlink="#<%=g_arColor(7)%>" alink="#<%=g_arColor(6)%>">
<center>

<!-- framing table -->
<table bgcolor="#<%=g_arColor(6)%>" border=0 cellpadding=2 cellspacing=0><tr><td>
<!-- end framing table -->

<table bgcolor="#<%=g_arColor(11)%>" border=0 cellpadding=3 cellspacing=0>
<form action="webCal4_group-deleted.asp" method="post">
<tr>
	<td bgcolor="#<%=g_arColor(4)%>" colspan=3>
	<font face="<%=g_arFont(0)%>" size=4>
	<b>Group Deletion</b></font></td>
<tr>
	<td colspan=3><font face="<%=g_arFont(0)%>" size=2>
	
<% if intUsers > 0 then %>
	
	<b>What should happen to the <%=intUsers & " " & strUserSay%> belonging only to the <%=strName%> group?</b><br>
	<input type="radio" name="user-do" value="delete">erase them and transfer events they've scheduled to
	<select name="user_del"><%=strUsersNot%></select><br>
	<input type="radio" name="user-do" value="move">transfer them to the 
	<select name="user_move"><%=strGroups%></select> group
	<p>
	
<% end if %>

<% if intEvents > 0 then %>
	
	<b>What should happen to the <%=intEvents & " " & strEventSay1%> scheduled only in the <%=strName%> group?</b><br>
	<input type="radio" name="event-do" value="delete">erase <%=strEventSay2%><br>
	<input type="radio" name="event-do" value="move">transfer <%=strEventSay2%> to the 
	<select name="event_move"><%=strGroups%></select> group
	<p>
	
<% end if %>

Are you sure you want to erase the <%=strName%> group?
	</font></td>
<tr>
	<td colspan=3 align="center" bgcolor="#<%=g_arColor(12)%>">
		<input type="submit" name="delete" value="Continue">
		<input type="submit" name="cancel" value="Cancel">
	</td>
<tr>
	<td align="center" colspan=3><font face="<%=g_arFont(0)%>" size=2>
	<b><font color="#cc0000">Caution</font>: erased groups cannot be restored</b></font></td>
</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->

<input type="hidden" name="group_id" value="<%=intGroupID%>">
<input type="hidden" name="event_count" value="<%=intEvents%>">
<input type="hidden" name="view" value="<%=Request.QueryString("view")%>">
<input type="hidden" name="users" value="<%=strUsers%>">
<input type="hidden" name="events" value="<%=strEvents%>">
</form>
</center>
</body>
</html>