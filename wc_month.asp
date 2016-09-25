<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./include/wc_settings_inc.asp"-->
<!--#include file="./include/wc_common_inc.asp"-->
<!--#include file="./include/wc_month_cls.asp"-->
<!--#include file="./language/wc_language.inc"-->
<%
' Copyright 1996-2004 Jason Abbott (webcal@webott.com)
Const m_sGRID = "0-1"
Const m_sVIEW = "month"

dim m_oMonth
%>
<html>
<head>
<link href="./style/wc_skin.css" rel="stylesheet">
<script language="javascript" src="./script/webCal4_help.js"></script>
<script language="javascript" src="./script/webCal4_buttons.js"></script>
<script language="javascript" src="./script/<%=g_sFILE_PREFIX%>functions.js"></script>
</head>
<body> <!-- onLoad="showMessage();"> -->

<%
Set m_oMonth = New wcMonth
Call m_oMonth.writeHTML(Request.Form("year"), Request.Form("month"), Request.QueryString("date"))
Set m_oMonth = Nothing
%>

<!--include file="./include/webCal4_month-options.inc"-->

</body>
</html>