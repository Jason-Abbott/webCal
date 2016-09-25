<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./data/webCal4_data.inc"-->
<!--#include file="./data/webCal4_cache.inc"-->
<!--#include file="./include/webCal4_constants.inc"-->
<!--#include file="./language/webCal4_language.inc"-->
<!--#include file="./include/webCal4_settings.inc"-->
<!--#include file="./include/webCal4_functions.inc"-->
<!--#include file="./include/webCal4_grid-functions.inc"-->
<!--#include file="./include/webCal4_week-functions.inc"-->
<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' Last updated 3/8/01

dim m_strQuery				' query passed to database
dim m_strHTML				' generated HTML
dim m_strTitle				' text displayed at top of calendar
dim m_strLoadFrom			' text to indicate page was cached
dim m_intStartTime			' time the query
dim m_strView				' calendar view type (month/week/day)
dim m_arDates				' date parameters for this week
dim m_arSegments			' array of segment properties

m_strView = "week"
m_arDates = getWeekDates()
m_intStartTime = milliTime()
m_arSegments = getGridProperties(Session(g_unique & "Segments")(g_WEEK))
m_strQuery = makeQuery(m_arDates(g_FIRST_DATE), m_arDates(g_LAST_DATE))
m_strHTML = readCache(m_strQuery, g_WEEK)

if m_strHTML = "" then
	' page not found in cache--retrieve from database
	m_strHTML = makeWeekHTML(m_strQuery)
	Call saveCache(m_strQuery, m_strHTML, m_arDates(g_FIRST_DATE), m_arDates(g_LAST_DATE), g_WEEK)
	m_strLoadFrom = "database"
Else
	m_strLoadFrom = "cache"
End if
	
' generate the title to display at the top of the calendar
m_strTitle = "<font face='" & g_arFont(1) & "' size=5 color='#" _
	& g_arColor(4) & "'><b>Week " & DatePart("ww", m_arDates(g_FIRST_DATE)) _
	& " in " & Year(m_arDates(g_FIRST_DATE)) & "</b></font>"
%>
<html>
<head>
<style>
<!--#include file="./style/webCal4_common.css"-->
<!--#include file="./style/webCal4_settings.css"-->
</style>
<script language="javascript" src="./script/webCal4_buttons.js"></script>
<script language="javascript" src="./script/webCal4_functions.js"></script>
<script language="javascript" src="./script/webCal4_functions-<%=g_strBrowser%>.js"></script>
</head>
<body onLoad="showMessage();" bgcolor="#<%=g_arColor(1)%>" link="#<%=g_arColor(7)%>" vlink="#<%=g_arColor(7)%>" alink="#<%=g_arColor(6)%>">

<!-- <%=m_strQuery%> -->

<table width="100%" border=0 cellspacing=0 cellpadding=1>
<!--#include file="./include/webCal4_buttons.inc"-->
<tr>
	<td bgcolor="#<%=g_arColor(6)%>" align="center" colspan=2><%=m_strHTML%></td>
<tr>
	<td valign="top">
	<font face="<%=g_arFont(1)%>" size=1>
	<%=showLoadTime(m_strQuery, m_strLoadFrom)%>
	<a href="http://webott.com/jason/webCal.html" target="_top">
	webCal 4.0</a> 
	</font>
	</td>
	<td align='right'><form>
	<%=makeButton(g_sBTN_LOGOUT,"javascript:logout();","logout",15,60)%>&nbsp;
	<%=makeButton(g_sBTN_DISPLAY,"javascript:showSettings();","show",15,160)%>
	</form></td>
</table>
<% response.flush %>

<!--#include file="./include/webCal4_grid-options.inc"-->

</body>
</html>