<% Option Explicit %>
<% Response.Buffer = true %>
<!--#include file="./include/wc_settings_inc.asp"-->
<!--#include file="./include/wc_common_inc.asp"-->
<!--#include file="./include/wc_day_cls.asp"-->
<!--#include file="./language/wc_language.inc"-->
<%
' Copyright 1996-2004 Jason Abbott (webcal@webott.com)

dim m_oDay
%>
<html>
<head>
<link href="./style/wc_skin.css" rel="stylesheet">
<link href="./style/webCal4_settings.css" rel="stylesheet">
<script language="javascript" src="./script/webCal4_buttons.js"></script>
<script language="javascript" src="./script/<%=g_sFILE_PREFIX%>functions.js"></script>
<script language="javascript" src="./script/webCal4_functions-<%'g_strBrowser%>.js"></script>
</head>
<body>

<% Set m_oDay = New wcDay %>
<% Call m_oDay.writeHTML(Request.QueryString("date")) %>
<% Set m_oDay = Nothing %>
</body>
</html>