<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./data/webCal4_data.inc"-->
<!--#include file="./include/webCal4_verify.inc"-->
<!--#include file="./include/webCal4_constants.inc"-->
<!--#include file="./include/webCal4_settings.inc"-->
<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' Last updated 2/27/2001

dim m_intGroupID	' group id
dim m_intUserID		' user id
dim m_intEventCount	' count of events in this group
dim m_strUser		' option selection for each user
dim m_strGroupName	' group name
dim m_bOverlap		' display overlap option?
dim m_strOverlap	' allow overlapping events
dim m_strGroupDesc	' group description
dim m_strType		' type of group modification
dim m_arAccess		' array of access levels
dim m_strAccessList	' selection list of permission levels
dim m_strQuery		' query passed to database
dim m_oConn			' connection object
dim m_oRS			' recordset object
dim m_strJSUsers	' javascript array of users
dim x				' loop counter

m_intGroupID = Request.Form("fldGroupID")
Set m_oRS = Server.CreateObject("ADODB.RecordSet")
Set m_oConn = Server.CreateObject("ADODB.Connection") : m_oConn.Open g_strDSN
m_arAccess = Array("Read only","Add to","Editor of","Manager of")
m_bOverlap = 1
m_strOverlap = " checked"

if Request.QueryString("action") = "edit" then
	' update an existing group	
	m_strType = "update"
	
	' count the number of events in this group
	m_strQuery = "SELECT COUNT(event_id) AS event_count " _
		& "FROM tblEventGroupScopes WHERE group_id=" & m_intGroupID
	m_oRS.Open m_strQuery, m_oConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	m_intEventCount = m_oRS("event_count")
	m_oRS.Close

	' retrieve information on selected group
	m_strQuery = "SELECT * FROM tblGroups WHERE (group_id)=" & m_intGroupID
	m_oRS.Open m_strQuery, m_oConn, adOpenForwardOnly, adLockReadOnly, adCmdText
		m_strGroupName = m_oRS("group_name")
		m_strGroupDesc = m_oRS("group_description")
		if m_oRS("allow_overlap") then
			if m_intEventCount > 1 then
				' if the group has two or more events and overlap
				' is already enabled then it cannot be disabled
				m_bOverlap = 0
			end if
		else
			m_strOverlap = ""
		end if
	m_oRS.Close : Set m_oRS = nothing

	' select users
	m_strQuery = "SELECT u.user_id, u.name_last + ', ' + u.name_first, " _
		& "u.default_access, p.access_level FROM (tblUsers AS u LEFT OUTER JOIN " _
		& "(SELECT user_id, group_id, access_level FROM tblPermissions WHERE group_id = " _
		& m_intGroupID & ") AS p ON u.user_id = p.user_id) "
else
	' select users
	m_strQuery = "SELECT u.user_id, u.name_last + ', ' + u.name_first, " _
		& "u.default_access, '' FROM tblUsers AS u "
	m_strType = "new"
	m_strGroupName = ""
	m_strGroupDesc = ""
end if

' only list users who aren't administrators
m_strQuery = m_strQuery & "WHERE u.default_access < " & g_ADMIN_ACCESS _
	& " ORDER BY u.name_last, u.name_first"
m_strJSUsers = getJSArray(m_strQuery, m_oConn)
m_oConn.Close : Set m_oConn = nothing

' generate permissions selection list
for x = 0 to UBound(m_arAccess)
	m_strAccessList = m_strAccessList & "<option value=" & x + 1 _
		& ">" & m_arAccess(x) & vbCrLf
next
%>
<html>
<head>
<script language="javascript" src="./script/webCal4_validate.js"></script>
<script language="javascript" src="./script/webCal4_functions.js"></script>
<script language="javascript" src="./script/webCal4_functions-group.js"></script>
<script language="javascript">
var m_aGroupUsers = [<%=m_strJSUsers%>];
</script>
</head>
<body onLoad="initPage();" bgcolor="#<%=g_arColor(1)%>" link="#<%=g_arColor(7)%>" vlink="#<%=g_arColor(7)%>" alink="#<%=g_arColor(6)%>">
<center>

<!-- framing table -->
<table bgcolor="#<%=g_arColor(6)%>" border=0 cellpadding=2 cellspacing=0><tr><td>
<!-- end framing table -->

<table bgcolor="#<%=g_arColor(11)%>" border=0 cellpadding=4 cellspacing=0>
<form name="frmEditGroup" method="post">
<tr>
	<td colspan=4 bgcolor="#<%=g_arColor(4)%>"><font face="<%=g_arFont(0)%>" size=4>
	<b>Group Details</b></font></td>
<tr>
	<td align="right" bgcolor="#<%=g_arColor(12)%>">
	<font face="<%=g_arFont(0)%>" size=2 color="#bb0000">Name</font></td>
	<td colspan=3><input type="text" name="fldGroupName" value="<%=m_strGroupName%>" size=15></td>

<% if m_bOverlap then %>
<tr>
	<td align="right" bgcolor="#<%=g_arColor(12)%>">
	<font face="<%=g_arFont(0)%>" size=2>Overlap</font></td>
	<td colspan=3><input type="checkbox" name="fldOverlap"<%=m_strOverlap%>>
	<font face="<%=g_arFont(0)%>" size=2>(allow events to overlap)</font></td>

<% else %>
	<input type="hidden" name="fldOverlap" value="on">
<% end if %>

<tr>
	<td align="right" bgcolor="#<%=g_arColor(12)%>">
	<font face="<%=g_arFont(0)%>" size=2>Description</font></td>
	<td colspan=3><input type="text" name="fldGroupDescription" value="<%=m_strGroupDesc%>" size=25></td>
<tr>
	<td valign="top" align="right" bgcolor="#<%=g_arColor(12)%>">
	<font face="<%=g_arFont(0)%>" size=2>Membership</font></td>
	
	<td valign="top" align="center" bgcolor="#<%=g_arColor(13)%>">
	<font face="<%=g_arFont(0)%>" size=1>Members</font><br>
	<select name="fldMembers" size="5" onChange="newUser(this);"></select></td>
	
	<td valign="center" align="center">
	<input type="button" name="add" value="&lt;-" onClick="moveUser('add');">
	<p>
	<input type="button" name="remove" value="-&gt;" onClick="moveUser('remove');">
	</td>
	
	<td width="30%" valign="top" align="center">
	<font face="<%=g_arFont(0)%>" size=1>Non-members</font><br>
	<select name="fldNonMembers" size="5"></select></td>
<tr>
	<td align="right" bgcolor="#<%=g_arColor(12)%>">
	<font face="<%=g_arFont(0)%>" size=2>Access</font></td>
	<td align="center" bgcolor="#<%=g_arColor(13)%>">
		<select name="fldAccessLevel" onChange="newAccess(this);"><%=m_strAccessList%></select></td>
	<td colspan='2' bgcolor="#<%=g_arColor(11)%>">&nbsp;</td>
<tr>
	<td colspan=4 align="center" bgcolor="#<%=g_arColor(12)%>">
	<input type="button" value="Save" onClick='saveGroup("frmEditGroup", false);'>
	<input type="button" value="Save & Add Another"	onClick='saveGroup("frmEditGroup", true);'>
	<input type="button" value="Cancel" onClick='goPage("webCal4_admin.asp","frmEditGroup");'>
	</td>

</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->

<input type="hidden" name="fldEditType" value="<%=m_strType%>">
<input type="hidden" name="fldGroupID" value="<%=Request.Form("fldGroupID")%>">
<input type="hidden" name="fldURL" value="<%=Request.Form("fldURL")%>">
<input type="hidden" name="fldView" value="<%=Request.QueryString("fldView")%>">
<input type="hidden" name="fldAccessList" value="">
</form>

<font face="<%=g_arFont(0)%>" size=2 color="#bb0000"><b>Red fields are required</b></font>

</center>
</body>
</html>