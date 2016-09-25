<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./data/webCal4_data_inc.asp"-->
<!--#include file="./include/webCal4_verify_inc.asp"-->
<!--#include file="./include/webCal4_settings_inc.asp"-->
<!--#include file="./include/webCal4_constants_inc.asp"-->
<!--#include file="./include/webCal4_functions_inc.asp"-->
<!--#include file="./include/webCal4_message_inc.asp"-->
<%
' Copyright 1999 Jason Abbott (webcal@webott.com)
' Last updated 2/24/2001

dim m_strQuery		' query string passed to database
dim m_strLCIDlist	' selection list of location ids
dim m_oConn			' connection object
dim m_strUserList	' HTML select list of users
dim m_strGroupList	' HTML select list of groups

Set m_oConn = Server.CreateObject("ADODB.Connection")
m_oConn.Open g_strDSN

' create drop down list of users
m_strQuery = "SELECT user_id, name_last + ', ' + name_first FROM tblUsers " _
	& "WHERE user_id <> 1 ORDER BY name_last, name_first"
m_strUserList = makeList(m_strQuery, "", m_oConn)

' create drop down list of groups
m_strQuery = "SELECT group_id, group_name FROM tblGroups " _
	& "ORDER BY group_name"
m_strGroupList = makeList(m_strQuery, "", m_oConn)

' create location id selection list
m_strLCIDlist = makeLCIDList(m_oConn)

m_oConn.Close : Set m_oConn = nothing
%>
<html>
<head>
<script language="javascript" src="./script/webCal4_functions.js"></script>
<script language="javascript" SRC="./script/webCal4_functions-admin.js"></script>
</head>
<body onLoad="initPage();" bgcolor="#<%=g_arColor(1)%>" link="#<%=g_arColor(7)%>" vlink="#<%=g_arColor(7)%>" alink="#<%=g_arColor(6)%>">

<!-- layout table -->
<center><table><tr><td>
<!-- end layout table -->

<!-- framing table -->
<table bgcolor="#<%=g_arColor(6)%>" border=0 cellpadding=2 cellspacing=0><tr><td>
<!-- end framing table -->

<table bgcolor="#<%=g_arColor(11)%>" border=0 cellpadding=3 cellspacing=0>
<form name="frmUser" method="post">
<tr>
	<td bgcolor="#<%=g_arColor(4)%>" colspan=2>
	<a href="webCal4_<%=Request.QueryString("view")%>.asp"
	onMouseOver="iconOver('Month'); return true;" 
    onMouseOut="iconOut('Month'); return true;">
	<img name="Month" src="./images/icon_calprev_grey.gif" width=15 height=16 alt="" border=0></a>
	<font face="<%=g_arFont(0)%>" size=4>
	<b>Users</b></font></td>
<tr>
	<td align="right" valign="top" bgcolor="#<%=g_arColor(12)%>">
	<input type="button" value="Add" onClick='goPage("webCal4_user-edit.asp?action=add","frmUser");'>
	</td>
	<td><font face="<%=g_arFont(1)%>" size=2>
	a new user</font></td>

<% if m_strUserList <> "" then %>
	
<tr>
	<td align="right" valign="top" bgcolor="#<%=g_arColor(12)%>">
	<input type="button" value="Edit" onClick='goPage("webCal4_user-edit.asp?action=edit","frmUser");'>
	</td>
	<td><font face="<%=g_arFont(1)%>" size=2>
	the selected user</font></td>
<tr>
	<td align="right" valign="top" bgcolor="#<%=g_arColor(12)%>">
	<input type="button" value="Delete" onClick='goPage("webCal4_user-delete.asp","frmUser");'>
	</td>
	<td><font face="<%=g_arFont(1)%>" size=2>
	the selected user</font></td>
<tr>
	<td bgcolor="#<%=g_arColor(12)%>" align="right">
	<font face="<%=g_arFont(1)%>" size=2>user:</font></td>
	<td bgcolor="#<%=g_arColor(12)%>">
	<select name="fldUserID"><%=m_strUserList%></select>
	</td>

<% else %>

<tr>
	<td bgcolor="#<%=g_arColor(5)%>"><font size=1>&nbsp;</td>
	<td bgcolor="#<%=g_arColor(5)%>"><font size=1>&nbsp;</td>
	
<% end if %>
</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->

<input type="hidden" name="fldView" value="<%=Request.QueryString("view")%>">
</form>

<!-- layout table -->
</td><td>
<!-- end layout table -->

<!-- framing table -->
<table bgcolor="#<%=g_arColor(6)%>" border=0 cellpadding=2 cellspacing=0><tr><td>
<!-- end framing table -->

<table bgcolor="#<%=g_arColor(11)%>" border=0 cellpadding=3 cellspacing=0>
<form name="frmGroup" method="post">
<tr>
	<td bgcolor="#<%=g_arColor(4)%>" colspan=2>
	<a href="webCal4_<%=Request.QueryString("view")%>.asp"
	onMouseOver="iconOver('Month'); return true;" 
    onMouseOut="iconOut('Month'); return true;">
	<img name="Month" src="./images/icon_calprev_grey.gif" width=15 height=16 alt="" border=0></a>
	<font face="<%=g_arFont(0)%>" size=4>
	<b>Groups</b></font></td>
<tr>
	<td align="right" valign="top" bgcolor="#<%=g_arColor(12)%>">
	<input type="button" value="Add" onClick='goPage("webCal4_group-edit.asp","frmGroup");'>
	</td>
	<td><font face="<%=g_arFont(1)%>" size=2>
	a new group</font></td>

<% if m_strGroupList <> "" then %>
	
<tr>
	<td align="right" valign="top" bgcolor="#<%=g_arColor(12)%>">
	<input type="button" value="Edit" onClick='goPage("webCal4_group-edit.asp?action=edit","frmGroup");'>
	</td>
	<td><font face="<%=g_arFont(1)%>" size=2>
	the selected group</font></td>
<tr>
	<td align="right" valign="top" bgcolor="#<%=g_arColor(12)%>">
	<input type="button" value="Delete" onClick='goPage("webCal4_group-delete.asp","frmGroup");'>
	</td>
	<td><font face="<%=g_arFont(1)%>" size=2>
	the selected group</font></td>
<tr>
	<td bgcolor="#<%=g_arColor(12)%>" align="right">
	<font face="<%=g_arFont(1)%>" size=2>group:</font></td>
	<td bgcolor="#<%=g_arColor(12)%>">
	<select name="fldGroupID"><%=m_strGroupList%></select>
	</td>

<% end if %>
	
</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->

<input type="hidden" name="fldView" value="<%=Request.QueryString("view")%>">
</form>

<!-- layout table -->
</td><tr><td colspan=2>
<!-- end layout table -->

<form name="frmLCID">
<input type="submit" value="Set language to">
<select name="fldLCID"><%=m_strLCIDlist%></select>

<input type="button" onClick="newColor('back1');">

<script language="javascript">
var ie4 = document.all;
var ns4 = document.layers;

function newColor(el) {
	if (ie4) {
		elDiv = eval(el + ".style");
		elDiv.background = "ff0000";
	} else if (ns4) {
		elDiv = document.layers[el];
		elDiv.visibility = "hide";
	}
}
//background-color : Aqua;

</script>
<style>
.back { position:relative; }
</style>

<%
'dim x
'Set m_oRS = Server.CreateObject("ADODB.Recordset")

'm_strQuery = "SELECT * FROM cal_colm_oRS ORDER BY color_name"
'm_oRS.Open m_strQuery, g_strDSN, adOpenForwardOnly, adLockReadOnly, adCmdText

'x = 1
'do while not m_oRS.EOF
'	response.write "<table bgcolor='#" & m_oRS("hexBack") & "' width='100%'>" _
'		& "<tr><td><div id='back" & x & "' class='back'>" _
'		& "<table cellspacing=2 border=0 cellpadding=0 " _
'		& "bgcolor='#" & m_oRS("hex4") & "' align='right'><tr><td>" _
'		& "<table cellpadding=3 cellspacing=0 border=0><tr>" _
'		& "<td colspan=3 bgcolor='#" & m_oRS("hex3") & "'>Title</td>" _
'		& "<tr><td bgcolor='#" & m_oRS("hexWeekday") & "'>weekday</td>" _
'		& "<td bgcolor='#" & m_oRS("hexToday") & "'>today</td>" _
'		& "<td bgcolor='#" & m_oRS("hexWeekend") & "'>weekend</td>" _
'		& "</table></td></table><nobr>background</nobr></div></td></table><br>"
'	x = x + 1
'	m_oRS.MoveNext
'loop
'm_oRS.Close
'Set m_oRS = nothing
%>

</form>

<!-- layout table -->
</td></table></center>
<!-- end layout table -->

<%=showMsg(Request.QueryString("message"),g_arColor(6),"ffffff",g_arFont(0))%>

</body>
</html>