<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./data/webCal4_data.inc"-->
<!--#include file="./include/webCal4_verify.inc"-->
<!--#include file="./include/webCal4_constants.inc"-->
<!--#include file="./include/webCal4_settings.inc"-->
<!--#include file="./include/webCal4_functions.inc"-->
<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' Last updated 2/24/2001

dim m_strType			' type of user modification
dim m_strQuery			' query string passed to database
dim m_oGroups			' dictionary of group/access for this user
dim m_intGroupID		' group id
dim m_intUserID			' user id
dim m_strLCIDlist		' HTML option list of location ids
dim m_intLCID			' user's location id
dim m_strFirstName		' first name of this user
dim m_strLastName		' last name of user
dim m_strEmail			' email name of user
dim m_strLogin			' user login
dim m_strPassword		' user password
dim m_arAccess			' user's access to groups
dim m_strAccessList		' HTML selection list of access levels
dim m_intDefaultAccess	' default access level
dim m_strDefaultList	' HTML selection list of default access levels
dim m_strStartPage		' user's start page
dim m_strStartPageList	' HTML selection list of start pages
dim m_oConn				' connection object
dim m_oRS				' recordset object
dim m_oRSgroups			' recordset of groups
dim m_strGroupList		' HTML select list of groups
dim m_intPos			' position of available group among all groups
dim m_strJSGroups		' javascript array of groups
dim m_strJSOrigin		' copy of groups to compare with
dim m_strHelp			' help text for this page
dim m_strTemp			' temporarily hold javascript string
dim m_intPassLength		' minimum password length
dim m_strPassLength		' text describing minimum length
dim x					' loop counter

m_intUserID = Request.Form("fldUserID")
m_intPassLength = Application(g_unique & "PassLength")
if m_intPassLength = "" or m_intPassLength = 0 then
	m_strPassLength = ""
else
	m_strPassLength = " (at least " & m_intPassLength & " chars.)"
end if

m_arAccess = Array("No access to","Read-only in","Add to","Editor of","Manager of")
Set m_oRS = Server.CreateObject("ADODB.RecordSet")
Set m_oConn = Server.CreateObject("ADODB.Connection") : m_oConn.Open g_strDSN
Set m_oGroups = Server.CreateObject("Scripting.Dictionary")

if Request.QueryString("action") = "edit" then
	' get details on the selected user
	m_strType = "update"
	m_strQuery = "SELECT * FROM tblUsers WHERE " _
		& "(user_id)=" & m_intUserID
	m_oRS.Open m_strQuery, m_oConn, adOpenForwardOnly, adLockReadOnly, adCmdText
		m_strFirstName = m_oRS("name_first")
		m_strLastName = m_oRS("name_last")
		m_strEmail = m_oRS("user_email")
		m_strLogin = m_oRS("user_login")
		m_strPassword = m_oRS("user_password")
		m_intDefaultAccess = CInt(m_oRS("default_access"))
		m_strStartPage = m_oRS("start_page")
		m_intLCID = CInt(m_oRS("user_lcid"))
	m_oRS.Close

	' transfer allowed groups to dictionary
	m_strQuery = "SELECT * FROM tblPermissions WHERE user_id=" & m_intUserID
	m_oRS.Open m_strQuery, m_oConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	do while not m_oRS.EOF
		m_oGroups.Add CStr(m_oRS("group_id")), CInt(m_oRS("access_level"))
		m_oRS.MoveNext
	loop
	m_oRS.Close
else
	' otherwise create new user details
	m_strType = "new"
	m_strFirstName = ""
	m_strLastName = ""
	m_strEmail = ""
	m_strLogin = ""
	m_strPassword = ""
	m_intDefaultAccess = 0
	m_strStartPage = ""
	m_intLCID = Session(g_unique & "LCID")
end if

' get a list of all groups
m_strQuery = "SELECT * FROM tblGroups"
m_oRS.Open m_strQuery, m_oConn, adOpenForwardOnly, adLockReadOnly, adCmdText
do while not m_oRS.EOF
	m_intGroupID = m_oRS("group_id")
	m_strGroupList = m_strGroupList & "<option value='" & m_intGroupID _
		& "'>" & m_oRS("group_name") & vbCrLf

	' generate format for JavaScript collection
	m_strTemp = "['" & m_intGroupID & "'] = "
	if m_oGroups.Exists(CStr(m_intGroupID)) then
		m_strTemp = m_strTemp & m_oGroups.Item(CStr(m_intGroupID)) & "; "
	else
		m_strTemp = m_strTemp & m_intDefaultAccess & "; "
	end if

	m_strJSGroups = m_strJSGroups & "m_oNewGroup" & m_strTemp
	m_strJSOrigin = m_strJSOrigin & "m_oOldGroup" & m_strTemp
	m_oRS.MoveNext
loop
Set m_oGroups = nothing
m_oRS.Close : Set m_oRS = nothing

m_strLCIDlist = makeLCIDList(m_oConn)

m_oConn.Close : Set m_oConn = nothing

' generate permissions selection lists
for x = 0 to UBound(m_arAccess)
	m_strAccessList = m_strAccessList & "<option value=" & x _
		& ">" & m_arAccess(x) & vbCrLf
next
m_strDefaultList = makeSelected(m_strAccessList _
	& "<option value=5>Administrator", m_intDefaultAccess)
%>
<html>
<head>
<script language="javascript" SRC="./script/webCal4_functions.js"></script>
<script language="javascript" SRC="./script/webCal4_functions-user.js"></script>
<script language="javascript" SRC="./script/webCal4_validate.js"></script>
<script language="javascript">
// group ids and their access levels for this user
m_oNewGroup = new Object();
<%=m_strJSGroups%>

// make an identical copy to compare changes to
m_oOldGroup = new Object();
<%=m_strJSOrigin%>
</script>
</head>
<body onLoad="initPage('<%=m_strType%>');" bgcolor="#<%=g_arColor(1)%>" link="#<%=g_arColor(7)%>" vlink="#<%=g_arColor(7)%>" alink="#<%=g_arColor(6)%>">
<center>

<!-- framing table -->
<table bgcolor="#<%=g_arColor(6)%>" border=0 cellpadding=2 cellspacing=0 width="25%"><tr><td>
<!-- end framing table -->

<table bgcolor="#<%=g_arColor(11)%>" border=0 cellpadding=4 cellspacing=0>
<form name="frmEdit" method="post">
<tr bgcolor="#<%=g_arColor(4)%>" valign="bottom">
	<td colspan=2><font face="<%=g_arFont(0)%>" size=4>
	<b>User Details</b></font></td>
<tr>
	<td align="right" bgcolor="#<%=g_arColor(12)%>">
	<font face="<%=g_arFont(0)%>" size=2 color="#bb0000">First Name</font></td>
	<td><input type="text" name="fldNameFirst" value="<%=m_strFirstName%>" size=15></td>
<tr>
	<td align="right" bgcolor="#<%=g_arColor(12)%>">
	<font face="<%=g_arFont(0)%>" size=2>Last Name</font></td>
	<td><input type="text" name="fldNameLast" value="<%=m_strLastName%>" size=15></td>
<tr>
	<td align="right" bgcolor="#<%=g_arColor(12)%>">
	<font face="<%=g_arFont(0)%>" size=2>e-mail</td>
	<td><nobr><input type="text" name="fldEmail" value="<%=m_strEmail%>" size=15></td>
<tr>
	<td align="right" bgcolor="#<%=g_arColor(12)%>">
	<font face="<%=g_arFont(0)%>" size=2 color="#bb0000"><nobr>Login Name</nobr></font></td>
	<td><input type="text" name="fldLogin" value="<%=m_strLogin%>" size=15></td>
<tr>
	<td align="right" bgcolor="#<%=g_arColor(12)%>">
	<font face="<%=g_arFont(0)%>" size=2 color="#bb0000">Password</font></td>
	<td><nobr><input type="password" name="fldPassword" value="<%=m_strPassword%>" size=15><%=m_strPassLength%></nobr></td>
<tr>
	<td align="right" bgcolor="#<%=g_arColor(12)%>">
	<font face="<%=g_arFont(0)%>" size=2 color="#bb0000">confirm</font></td>
	<td><input type="password" name="fldConfirm" value="<%=m_strPassword%>" size=15></td>
<tr>
	<td align="right" bgcolor="#<%=g_arColor(12)%>">
	<font face="<%=g_arFont(0)%>" size=2>Language</font></td>
	<td><select name="fldLCID"><%=m_strLCIDlist%></select></td>
<tr>
	<td bgcolor="#<%=g_arColor(13)%>" height="2">
	<img src="./images/tiny_blank.gif" width="1" height="2"></td>
	<td align="center" bgcolor="#<%=g_arColor(12)%>" height="2">
	<font face="<%=g_arFont(1)%>" size=1>
	<b><a href="#" onClick="javascript:viewHelp(); return false;">p e r m i s s i o n s</a></b></font></td>
<tr>
	<td align="right" bgcolor="#<%=g_arColor(13)%>">
	<font face="<%=g_arFont(0)%>" size=2>Default</font></td>
	<td><font face="<%=g_arFont(0)%>" size=2>
	<select name="fldDefault" onChange="newDefault(this);"><%=m_strDefaultList%></select> (all groups)</font></td>
<tr>
	<td align="right" bgcolor="#<%=g_arColor(13)%>">
	<font face="<%=g_arFont(0)%>" size=2>Per group</font></td>
	<td><select name="fldAccess" onChange="newAccess(this);"><%=m_strAccessList%></select>
		<select name="fldGroup" onChange="newGroup(this);"><%=m_strGroupList%></select>
	</td>
<tr>
	<td colspan=2 align="center" bgcolor="#<%=g_arColor(12)%>">
	<nobr>
	<input type="button" value="Save" onClick="saveUser('frmEdit', false);">
	<input type="button" value="Save & Add Another" onClick="saveUser('frmEdit', true);">
	<input type="button" value="Cancel" onClick='goPage("webCal4_admin.asp","frmEdit");'>
	</nobr>
	</td>

</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->

<font face="<%=g_arFont(0)%>" size=2 color="#bb0000">
<b>Red fields are required</b></font>
</center>

<% response.flush %>

<input type="hidden" name="fldAccessList" value="0">
<input type="hidden" name="fldEditType" value="<%=m_strType%>">
<input type="hidden" name="fldUserID" value="<%=m_intUserID%>">
<input type="hidden" name="fldURL" value="<%=Request.Form("fldURL")%>">
<input type="hidden" name="fldView" value="<%=Request.Form("fldView")%>">
</form>
</body>
</html>
