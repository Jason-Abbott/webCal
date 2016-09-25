<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./data/webCal4_data.inc"-->
<!--#include file="./include/webCal4_settings.inc"-->
<!--#include file="./include/webCal4_constants.inc"-->
<!--#include file="./include/webCal4_functions.inc"-->
<!--#include file="./language/webCal4_language.inc"-->
<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' Last updated 3/4/2001

' this presents and validates the login form

dim m_strError			' error to display if login fails
dim m_strStatus			' message returned to user after login attempt
dim m_oConn				' connection object
dim m_oRS				' recordset of user information
dim m_arUser			' user fields
dim m_arGroups			' holds rsAccess
dim m_arScopes			' holds scopes visible to each user
dim m_strQuery			' query string passed to db
dim m_intCount			' total groups visible to user
dim x					' loop counter

Const m_USER_ID = 0
Const m_PASSWORD = 1
Const m_LCID = 2
Const m_SHOW_WEEKEND = 3
Const m_WEEK_SEG_MINS = 4
Const m_WEEK_SEG_START = 5
Const m_WEEK_SEG_END = 6
Const m_DAY_SEG_MINS = 7
Const m_DAY_SEG_START = 8
Const m_DAY_SEG_END = 9
Const m_START_PAGE = 10

m_strStatus = g_sMSG_NEED_LOGIN
m_strError = g_sMSG_BAD_LOGIN
	
if Request.Form("fldUserName") <> "" then
	' login was attempted--validate
	m_strQuery = "SELECT user_id, user_password, user_lcid, show_weekend, " _
		& "week_seg_mins, week_seg_start, week_seg_end, " _
		& "day_seg_mins, day_seg_start, day_seg_end, start_page " _
		& "FROM tblUsers WHERE user_login = '" & Request.Form("fldUserName") & "'"
	Set m_oConn = Server.CreateObject("ADODB.Connection") : m_oConn.Open g_strDSN
	m_arUser = getRowArray(m_strQuery, m_oConn)
	
	if Not IsArray(m_arUser) then
		' login not found
		m_strStatus = m_strError
	else
		if m_arUser(m_PASSWORD) = Request.Form("fldPassword") then
			Session(g_unique & "UserID") = m_arUser(m_USER_ID)
			Session(g_unique & "LCID") = m_arUser(m_LCID)
			Session(g_unique & "Weekends") = (m_arUser(m_SHOW_WEEKEND) <> 0)
			Session(g_unique & "StartPage") = m_arUser(m_START_PAGE)
			' Array of arrays--dummy array of 0s for month values
			Session(g_unique & "Segments") = Array( _
				Array(m_arUser(m_WEEK_SEG_MINS), _
					m_arUser(m_WEEK_SEG_START), m_arUser(m_WEEK_SEG_END)), _
				Array(m_arUser(m_DAY_SEG_MINS), _
					m_arUser(m_DAY_SEG_START), m_arUser(m_DAY_SEG_END)), _
				Array(0,0,0))
			
			Call initDataAccess(m_oConn) : Call cleanUp()
			response.clear : response.redirect Request.Form("url")
		else
			' password doesn't match
			m_strStatus = m_strError
		end if
	end if
	Call cleanUp()
end if

' cleanup objects
Sub cleanUp()
	m_oConn.Close : Set m_oConn = nothing
End Sub
%>
<html>
<head>
<style>
<!--#include file="./style/webCal4_common.css"-->
<!--#include file="./style/webCal4_settings.css"-->
</style>
<script language="javascript" src="./script/webCal4_validate.js"></script>
<script language="javascript" src="./script/webCal4_functions-<%=Session(g_unique & "Browser")%>.js"></script>
<script language="javascript">
var m_oFields = { fldUserName:{desc:"Username",type:"String",req:1} };
</script>
</head>
<body bgcolor="#<%=g_arColor(1)%>" link="#<%=g_arColor(7)%>" vlink="#<%=g_arColor(7)%>" alink="#<%=g_arColor(6)%>">
<center>

<!-- framing table -->
<table bgcolor="#<%=g_arColor(6)%>" width="60%" border=0 cellpadding=2 cellspacing=0><tr><td>
<!-- end framing table -->

<table bgcolor="#<%=g_arColor(11)%>" border=0 cellpadding=3 cellspacing=0 width="100%">
<form name="frmLogin" action="webCal4_login.asp" method="post" onSubmit="return isValid('frmLogin', m_oFields);">
<tr bgcolor="#<%=g_arColor(4)%>" valign="bottom">
	<td colspan=4><font face="<%=g_arFont(0)%>" size=4>
	<b><%=g_sTITLE_LOGIN%></b></font></td>
<tr>
	<td colspan=4 align="center"><font face="<%=g_arFont(2)%>" size=2>
	<%=m_strStatus%><br></font></td>
<tr>
	<td>&nbsp;</td>
	<td bgcolor="#<%=g_arColor(12)%>" align="right"><font face="<%=g_arFont(2)%>"><%=g_sFRM_USERNAME%>:&nbsp;</td>
	<td bgcolor="#<%=g_arColor(12)%>"><input type="text" name="fldUserName" size=10 value="<%=Request.Form("login")%>"></td>
	<td>&nbsp;</td>
<tr>
	<td>&nbsp;</td>
	<td bgcolor="#<%=g_arColor(12)%>" align="right"><font face="<%=g_arFont(2)%>"><%=g_sFRM_PASSWORD%>:&nbsp;</td>
	<td bgcolor="#<%=g_arColor(12)%>"><input type="password" name="fldPassword" size=10></td>
	<td>&nbsp;</td>
<tr>
	<td colspan=4 align="center"><%=makeButton(g_sBTN_CONTINUE,"javascript:document.frmLogin.submit();","submit",15,80)%></td>
</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->

<input type="hidden" name="url" value="<%=Request("url")%>">
</form>

</center>
</body>
</html>