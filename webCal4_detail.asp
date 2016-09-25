<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./data/webCal4_data.inc"-->
<!--#include file="./include/webCal4_constants.inc"-->
<!--#include file="./include/webCal4_settings.inc"-->
<!--#include file="./include/webCal4_functions.inc"-->
<!--#include file="./include/webCal4_recur.inc"-->
<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' Last updated 2/24/2001

dim m_strRecur			' event recurrence type
dim m_strRecurTable		' tables of recurring dates
dim m_strDescription	' event description
dim m_strTitle			' event title
dim m_intEventID		' event id
dim m_strDate			' selected date
dim m_strShowDate		' formatted date
dim m_strStartTime		' event start time
dim m_strEndTime		' event end time
dim m_intCount			' number of event recurrences
dim m_strBody			' body text of detail page
dim m_strLeft			' text in left cell
dim m_strSay			' describe recurrence
dim m_arDates			' other dates on which this event occurs
dim m_intIndex			' index into array
dim m_intUserID			' user id of event creator
dim m_strHTML			' build HTML in functions
dim m_strURL			' store url in case we need to return
dim m_strQuery			' query passed to database
dim m_oConn				' connection to database object
dim m_oRS				' recordset object
dim m_strView
dim x					' loop counter

' remember URL as we may return directly to this page
' extra code is necessary to accomodate virtual hosts
m_strURL = Request.ServerVariables("PATH_TRANSLATED")
m_strURL = Right(m_strURL, Len(m_strURL) - InStrRev(m_strURL,Chr(92))) & "?" _
	& Server.URLEncode(Request.ServerVariables("QUERY_STRING"))

m_strView = Request.QueryString("view")
m_strDate = Request.QueryString("date")

m_strQuery = "SELECT * FROM tblEvents " _
	& "WHERE (event_id)=" & Request.QueryString("event_id")

' retrieve information from database
Set m_oConn	= Server.CreateObject("ADODB.Connection")
Set m_oRS = Server.CreateObject("ADODB.Recordset")
m_oConn.Open g_strDSN
m_oRS.Open m_strQuery, m_oConn, adOpenForwardOnly, adLockReadOnly, adCmdText
m_strDescription	= m_oRS("event_description")
m_strRecur		= m_oRS("event_recur")
m_strTitle		= m_oRS("event_title")
m_intEventID	= m_oRS("event_id")
m_strStartTime	= m_oRS("time_start")
m_strEndTime	= m_oRS("time_end")
m_intUserID		= m_oRS("user_id")

m_oRS.Close : Set m_oRS = nothing

' display formatting for recurring events
if m_strRecur <> "none" then
	' retrieve other dates from table
	m_strQuery = "SELECT event_date FROM tblEventDates" _
		& " WHERE event_id=" & m_intEventID _
		& " ORDER BY event_date"
	m_arDates = getRowArray(m_strQuery, m_oConn)
	m_strRecurTable = "<table bgcolor='" & g_arColor(2) _
		& "' width='100%'><form><tr><td>" _
		& showRecur(m_arDates, m_strDate, m_intEventID, m_strRecur, Request.Form("fldView")) _
		& "</td></form></table>"
	m_intCount = UBound(m_arDates) + 1
else
	m_strRecurTable = ""
	m_intCount = 1
end if

m_oConn.Close : Set m_oConn = nothing

' format date
m_strShowDate = "<font face='" & g_arFont(0) & "' color='" _
	& g_arColor(4) & "'><b><font size=2>" _
	& WeekdayName(WeekDay(m_strDate)) & "</font><br>" & vbCrLf _
	& "<font size=7>" & Day(m_strDate) & "</font><br>" & vbCrLf _
	& "<font size=5>" & MonthName(Month(m_strDate),1) & "</font></b><br>" _
	& "<font size=4>" & Year(m_strDate) & "</font></font>" & vbCrLf

' choose where in the table things should be placed
if m_strDescription <> "" then
	m_strBody = "<font size=2>" _
		& Replace(m_strDescription, VbCrLf, "<br>") _
		& "</font>"
	m_strLeft = "<td rowspan=2 align='center' valign='top'>" _
		& m_strShowDate & "</td>"
elseif m_strRecurTable <> "" then
	m_strBody = ""
	m_strLeft = "<td rowspan=2 align='center' valign='top'>" _
		& m_strShowDate & "</td>"
else
	m_strBody = "<br><center>" & m_strShowDate & "</center>"
	m_strLeft = ""
end if

' display time range if one was entered for this event (updated 2/23/01)
' returns string ---------------------------------------------------------
Function timeRange(ByVal v_strStartTime, ByVal v_strEndTime)
	dim intStartHour	' military start hour
	dim intEndHour		' military end hour
	dim strHour			' hour text
	dim strHourClr		' color of hour background
	dim strTextClr		' color of text
	dim strHTML

	if v_strStartTime <> "" then
		' build table of hours
		intStartHour = Hour(v_strStartTime)
		intEndHour = Hour(v_strEndTime)
		strHTML = strHTML & "<td rowspan=2 valign='top' width='1%'>" _
			& "<table cellspacing=1 cellpadding=0 border=0>"

		for x = 0 to 23
			if x = intStartHour then
				strHour = "<b>" & Replace(m_strStartTime, ":00 ", " ")  _
					& vbCrLf & "<center>" _
					& "<img src='./images/arrow_down_black.gif' width='7' height='4'>" _
					& "</center>" & vbCrLf _
					& Replace(v_strEndTime, ":00 ", " ") & "</b>"
				strHourClr = "ffffff"
				strTextClr = "000000"
				x = x + intEndHour - intStartHour
			else
				' show normal, unbolded time
				if x = 0 then
					strHour = "<b>midnight</b>"
				elseif x < 12 then
					strHour = x & ":00 AM"
				elseif x = 12 then
					strHour = "<b>noon</b>"
				else
					strHour = x - 12 & ":00 PM"
				end if
				strHourClr = g_arColor(2)
				strTextClr = g_arColor(5)
			end if

			strHTML = strHTML & "<tr><td bgcolor=""#" & strHourClr _
				& """ align=""right"" nowrap><font face='" & g_arFont(0) & "'" _
				& "size=1 color=""#" & strTextClr & """>" _
				& strHour & "</font></td>"
		next
		strHTML = strHTML & "</td></table>"
	else
		strHTML = ""
	end if
	timeRange = strHTML
End Function

' determine whether login is needed to edit event (updated 2/23/01)
' returns string ---------------------------------------------------------
Function needLogin(ByVal v_intCurrentUserID, ByVal v_intEventUserID, ByVal v_strURL)
	dim strHTML			' build HTML

	if v_intCurrentUserID <> v_intEventUserID then
		' display login key if current event belongs to different user
		strHTML = "<a href='webCal4_login.asp?url=" & v_strURL & "' " _
		 	& switchIcon("Key","","") & "><img name='Key' " _
			& "src='./images/icon_key_grey.gif' width=16 height=15 border=0></a>"
	else
		strHTML = ""
	end if
	needLogin = strHTML
End Function

%>
<html>
<head>
<style><!--#include file="./style/webCal4_common.css"--></style>
<script language="javascript" src="./script/webCal4_functions.js"></script>
<script language="javascript" src="./script/webCal4_buttons.js"></script>
</head>
<body bgcolor="#<%=g_arColor(1)%>" link="#<%=g_arColor(7)%>" vlink="#<%=g_arColor(7)%>" alink="#<%=g_arColor(6)%>">
<br>
<center>
<table border=0 cellspacing=0 cellpadding=3 width="450">
<tr>
	<%=m_strLeft%>
	<td valign="top">
		<table cellspacing=0 cellpadding=3 border=0 width="100%">
		<tr>
			<td bgcolor="#<%=g_arColor(4)%>">
				<font face="<%=g_arFont(1)%>" size=4><b><nobr>
				<a href="webCal4_<%=Request.QueryString("view")%>.asp?date=<%=m_strDate%>"
				<%=switchIcon("Prev","","Return to " & m_strView)%>><img name="Prev" src="./images/icon_calprev_grey.gif" width=15 height=16 alt="" border=0></a>
				<%=m_strTitle%></b></nobr></font>&nbsp;
				<%=needLogin(Session(g_unique & "UserID"), m_intUserID, m_strURL)%>
			</td>
		<tr>
			<td><%=m_strBody%></td>
		</table>

		<%=m_strRecurTable%>
	</td>
	<%=timeRange(m_strStartTime, m_strEndTime)%>

<!-- buttons -->

<% if Session(g_unique & "UserID") = m_intUserID then %>

<tr>
	<td valign="bottom">

	<!-- framing table -->
	<table bgcolor="#<%=g_arColor(5)%>" width="100%" cellspacing=0 cellpadding=2 border=0><tr><td>
	<!-- end framing table -->

	<table cellspacing=0 cellpadding=2 border=0 width="100%">
	<tr>
		<td align="right" bgcolor="#<%=g_arColor(12)%>">
		<form name="frmEdit" method="post">
		<input type="button" value="Edit" onClick='goPage("webCal4_event-edit.asp","frmEdit");'><br>
		<input type="button" value="Delete" onClick='goPage("webCal4_event-delete.asp","frmEdit");'>
		</td>

<%
	' NOTE that this excludes events which might be listed
	' as recurring but now occur on just one date (m_intCount=1)
	if m_strRecur <> "none" AND m_intCount > 1 then
		if m_intCount = 2 then
			m_strSay = "both"
		else
			m_strSay = "all " & m_intCount
		end if
%>
		<td bgcolor="#<%=g_arColor(12)%>"><font face="<%=g_arFont(0)%>" size=2>
		<input type="radio" name="fldTimeScope" value="one">only this occurence<br>
		<input type="radio" name="fldTimeScope" value="future">this and all future occurences<br>
		<input type="radio" name="fldTimeScope" value="all" checked><%=m_strSay%> occurences
		</font>

<%	end if %>

		</td>

	</table>

	<!-- framing table -->
	</td></table>
	<!-- end framing table -->

	<input type="hidden" name="fldEventID" value="<%=m_intEventID%>">
	<input type="hidden" name="fldDate" value="<%=m_strDate%>">
	<input type="hidden" name="fldCount" value="<%=m_intCount%>">
	<input type="hidden" name="fldURL" value="<%=m_strURL%>">
	<input type="hidden" name="fldEdit" value=1>
	<input type="hidden" name="fldView" value="<%=m_strView%>">
	</form>

<% end if %>

	</td>
</table>

</center>
</body>
</html>