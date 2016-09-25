<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./data/webCal4_data.inc"-->
<!--#include file="./include/webCal4_settings.inc"-->
<!--#include file="./data/webCal4_cache.inc"-->
<!--#include file="./include/webCal4_constants.inc"-->
<!--#include file="./include/webCal4_functions.inc"-->
<!--#include file="./include/webCal4_month-functions.inc"-->
<!--#include file="./language/webCal4_language.inc"-->
<%
' Copyright 2001 Jason Abbott (webcal@webott.com)

dim m_strTitle		' text displayed at top of calendar
dim m_strQuery		' query string passed to db
dim m_strHTML		' page HTML
dim m_strLoadFrom	' text to indicate page was cached
dim m_intStartTime	' time the query
dim m_arDates		' date parameters for this month
dim m_strView		' calendar view
dim m_strGrid

m_strGrid = "0-1"
m_strView = "month"
m_arDates = getMonthDates()
m_intStartTime = milliTime()
m_strQuery = makeQuery(m_arDates(g_FIRST_DATE), m_arDates(g_LAST_DATE))
'm_strHTML = readCache(m_strQuery, g_MONTH)

if m_strHTML = "" then
	' query is not cached--retrieve from database
	m_strHTML = makeMonthHTML(m_strQuery)
	Call saveCache(m_strQuery, m_strHTML, m_arDates(g_FIRST_DATE), m_arDates(g_LAST_DATE), g_MONTH)
	m_strLoadFrom = g_MSG_DATABASE
Else
	m_strLoadFrom = g_MSG_CACHE
End if

m_strTitle = "<div class='viewName'>" & MonthName(Month(m_arDates(g_FIRST_DATE))) & " " _
	& Year(m_arDates(g_FIRST_DATE)) & "</div>"
%>
<html>
<head>
<link href="./style/webCal4_common.css" rel="stylesheet">
<link href="./style/webCal4_settings.css" rel="stylesheet">
<script language="javascript" src="./script/webCal4_help.js"></script>
<script language="javascript" src="./script/webCal4_buttons.js"></script>
<script language="javascript" src="./script/webCal4_functions.js"></script>
<script language="javascript" src="./script/webCal4_functions-<%=g_strBrowser%>.js"></script>
</head>

<body onLoad="showMessage();">
<!-- <%=m_strQuery%> -->

<table width="100%" border='0' cellspacing='0' cellpadding='1'>
<!--#include file="./include/webCal4_buttons.inc"-->
<tr>
	<td bgcolor="#<%=g_arColor(6)%>" align="center" colspan=2><%=m_strHTML%></td>
<tr>
	<td valign="top"><div class='footnote'>
	<%=showLoadTime(m_strQuery, m_strLoadFrom)%>
	<a href="http://webott.com/jason/webCal.html" target="_top">
	webCal 4.0</a>
	</div>
	</td>
	<td align='right'><form>
	<%=makeButton(g_BTN_LOGOUT,"logout();",12,60)%>&nbsp;
	<%=makeButton(g_BTN_DISPLAY,"showSettings();",12,160)%>
	</form></td>
</table>

<!--#include file="./include/webCal4_month-options.inc"-->

</body>
</html>