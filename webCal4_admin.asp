<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./data/webCal4_data.inc"-->
<!--#include file="./include/webCal4_verify.inc"-->
<!--#include file="./include/webCal4_settings.inc"-->
<!--#include file="./include/webCal4_constants.inc"-->
<!--#include file="./include/webCal4_functions.inc"-->
<!--#include file="./language/webCal4_language.inc"-->
<!--#include file="./include/webCal4_message.inc"-->
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
<link href="./style/webCal4_common.css" rel="stylesheet">
<link href="./style/webCal4_settings.css" rel="stylesheet">
<script language="javascript" src="./script/webCal4_functions.js"></script>
<script language="javascript" SRC="./script/webCal4_functions-admin.js"></script>
<script language="javascript" src="./script/webCal4_functions-<%=Session(g_unique & "Browser")%>.js"></script>
</head>
<body onLoad="initPage();" bgcolor="#<%=g_arColor(1)%>" link="#<%=g_arColor(7)%>" vlink="#<%=g_arColor(7)%>" alink="#<%=g_arColor(6)%>">

<!-- layout table -->
<center><table><tr><td>
<!-- end layout table -->

<div class='dialog'>
<table border='0' cellpadding='0' cellspacing='0' class='form'>
<form name="frmUser" method="post">
<tr>
	<td class='titleBar'><%=g_sTITLE_USERS%></td>
	<td align='right' class='titleBar'>
	<a href="webCal4_<%=Request.QueryString("view")%>.asp" class='SymbolButton'>
	<%=showSymbol(g_CHAR_CLOSE, 12)%></a>
	</td>
<tr>
	<td align="right" valign="top" class='formDark'>
	<%=makeButton(g_sBTN_ADD,"goPage('webCal4_user-edit.asp?action=add','frmUser');",17,60)%>
	</td>
	<td>&nbsp;<%=g_sFRM_NEW%>&nbsp;<%=g_sFRM_USER%></td>

<% if m_strUserList <> "" then %>
	
<tr>
	<td align="right" valign="top" class='formDark'>
	<%=makeButton(g_sBTN_EDIT,"goPage('webCal4_user-edit.asp?action=edit','frmUser');",17,60)%>
	</td>
	<td>&nbsp;<%=g_sFRM_SEL%>&nbsp;<%=g_sFRM_USER%>&nbsp;</td>
<tr>
	<td align="right" valign="top" class='formDark'>
	<%=makeButton(g_sBTN_DELETE,"goPage('webCal4_user-delete.asp','frmUser');",17,60)%>
	</td>
	<td>&nbsp;<%=g_sFRM_SEL%>&nbsp;<%=g_sFRM_USER%>&nbsp;</td>
<tr>
	<td align="right" class='formDark'><%=g_sFRM_USER%>:</td>
	<td class='formDark'><select name="fldUserID"><%=m_strUserList%></select></td>

<% else %>

<tr>
	<td><font size=1>&nbsp;</td>
	<td><font size=1>&nbsp;</td>
	
<% end if %>
</table>
</div>

<input type="hidden" name="fldView" value="<%=Request.QueryString("view")%>">
</form>

<!-- layout table -->
</td><td>
<!-- end layout table -->

<div class='dialog'>
<table border='0' cellpadding='0' cellspacing='0' class='form'>
<form name="frmGroup" method="post">
<tr>
	<td class='titleBar'><%=g_sTITLE_GROUPS%></td>
	<td align='right' class='titleBar'>
	<a href="webCal4_<%=Request.QueryString("view")%>.asp" class='SymbolButton'>
	<%=showSymbol(g_CHAR_CLOSE, 12)%></a>
	</td>
<tr>
	<td align="right" valign="top" class='formDark'>
	<%=makeButton(g_sBTN_ADD,"goPage('webCal4_group-edit.asp','frmGroup');",17,60)%>
	</td>
	<td>&nbsp;<%=g_sFRM_NEW%>&nbsp;<%=g_sFRM_GROUP%></td>

<% if m_strGroupList <> "" then %>
	
<tr>
	<td align="right" valign="top" class='formDark'>
	<%=makeButton(g_sBTN_EDIT,"goPage('webCal4_group-edit.asp?action=edit','frmGroup');",17,60)%>
	</td>
	<td>&nbsp;<%=g_sFRM_SEL%>&nbsp;<%=g_sFRM_GROUP%>&nbsp;</td>
<tr>
	<td align="right" valign="top" class='formDark'>
	<%=makeButton(g_sBTN_DELETE,"goPage('webCal4_group-delete.asp','frmGroup');",17,60)%>
	</td>
	<td>&nbsp;<%=g_sFRM_SEL%>&nbsp;<%=g_sFRM_GROUP%>&nbsp;</td>
<tr>
	<td align="right" class='formDark'><%=g_sFRM_GROUP%>:</td>
	<td class='formDark'><select name="fldGroupID"><%=m_strGroupList%></select></td>

<% end if %>
	
</table>
</div>

<input type="hidden" name="fldView" value="<%=Request.QueryString("view")%>">
</form>

<!-- layout table -->
</td><tr><td colspan=2>
<!-- end layout table -->

<form name="frmLCID">
<%=makeButton(g_sBTN_LANGUAGE,"document.frmLCID.submit();",17,120)%>
<select name="fldLCID"><%=m_strLCIDlist%></select>
</form>

<!-- layout table -->
</td></table></center>
<!-- end layout table -->

<%=showMsg(Request.QueryString("message"),g_arColor(6),"ffffff",g_arFont(0))%>

</body>
</html>