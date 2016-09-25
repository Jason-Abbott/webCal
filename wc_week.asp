<% Option Explicit %>
<% Response.Buffer = true %>
<!--#include file="./include/wc_settings_inc.asp"-->
<!--#include file="./include/wc_common_inc.asp"-->
<!--#include file="./include/wc_week_cls.asp"-->
<!--#include file="./language/wc_language.inc"-->
<%
' Copyright 1996-2004 Jason Abbott (webcal@webott.com)
dim m_oWeek
'Request.Form("year"), Request.Form("month"), 
%>
<html>
<head>
<link href="./style/wc_skin.css" rel="stylesheet">
<link href="./style/webCal4_settings.css" rel="stylesheet">
<script language="javascript" src="./script/webCal4_buttons.js"></script>
<script language="javascript" src="./script/<%=g_sFILE_PREFIX%>functions.js"></script>
</head>
<body>
<%
Set m_oWeek = New wcWeek
Call m_oWeek.writeHTML(Request.QueryString("date"))
Set m_oWeek = Nothing
%>
</body>
</html>