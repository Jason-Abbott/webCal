<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./include/wc_settings_inc.asp"-->
<!--#include file="./include/wc_constants_inc.asp"-->
<!--#include file="./include/wc_functions_inc.asp"-->
<!--#include file="./include/wc_data_cls.asp"-->
<!--#include file="./include/wc_session_cls.asp"-->
<!--#include file="./include/wc_user_cls.asp"-->
<!--#include file="./include/wc_layout_cls.asp"-->
<!--#include file="./language/wc_language.inc"-->
<%
' Copyright 1996-2004 Jason Abbott (webcal@webott.com)
' this presents and validates the login form

dim m_sError
dim m_sStatus
dim m_oLayout
dim m_oUser

m_sStatus = g_sMSG_NEED_LOGIN

with Request
	If .Form("fldUserName") <> "" Then
		Set m_oUser = New wcUser
		m_sStatus = m_oUser.Login(.Form("fldUserName"), .Form("fldPassword"))
		Set m_oUser = Nothing
		If IsVoid(m_sStatus) Then
			response.redirect .Form("url")
		End If
	End If
end with

'Set m_oLayout = New wcLayout
%>
<html>
<head>
<link href="./style/wc_skin.css" rel="stylesheet">
<link href="./style/webCal4_settings.css" rel="stylesheet">
<script language="javascript" src="./script/webCal4_validate.js"></script>
<!-- <script language="javascript" src="./script/webCal4_functions-<%'Session(g_sDB_NAME & "Browser")%>.js"></script> -->
<script language="javascript">
var m_oFields = { fldUserName:{desc:"Username",type:"String",req:1} };
</script>
</head>
<body>
<center>

<table border='0' cellpadding='3' cellspacing='0' class='DialogBox'>
<form name="frmLogin" action="<%=g_sFILE_PREFIX%>login.asp" method="post" onSubmit="return isValid('frmLogin', m_oFields);">
<tr>
	<td colspan='4' class='BoxTitle'><%=g_sTITLE_LOGIN%></td>
<tr>
	<td colspan='4' align="center"><%=m_sStatus%></td>
<tr>
	<td>&nbsp;</td>
	<td class='BoxLabel'><%=g_sFRM_USERNAME%>:</td>
	<td class='BoxField'><input type="text" name="fldUserName" size='10' value="<%=Request.Form("fldUserName")%>"></td>
	<td>&nbsp;</td>
<tr>
	<td>&nbsp;</td>
	<td class='BoxLabel'><%=g_sFRM_PASSWORD%>:</td>
	<td class='BoxField'><input type="password" name="fldPassword" size='10'></td>
	<td>&nbsp;</td>
<tr>
	<td colspan='4' align='center'><input type='submit' value='Login'></td>
</table>

<input type="hidden" name="url" value="<%=Request("url")%>">
</form>

</center>
</body>
</html>