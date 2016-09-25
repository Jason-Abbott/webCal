<!-- 
Copyright 1999 Jason Abbott (jason@webott.com)
Last updated 06/04/1999
-->

<!--#include file="data/webCal3_data.inc"-->
<!--#include file="webCal3_themes.inc"-->
<!--#include file="show_status.inc"-->
<%
dim rs, url, eventDescription, eventRecur, eventTitle, eventID
dim eventDate, eventStart, eventEnd

' these are the variables used by the included webCal3_showrecur:

dim monthList(12), eventYear, x, years

' store the url since we may need it to return to this
' page after authentication

' accomodate virtual host (jea:3/8/00)
strPath = Request.ServerVariables("PATH_TRANSLATED")
strPath = Right(strPath, Len(strPath) - InStrRev(strPath,"\"))
url = strPath & "?" _
	& Request.ServerVariables("QUERY_STRING")

' pull the event information and dates from Access

query = "SELECT * FROM cal_events " _
	& "WHERE (event_id)=" & Request.QueryString("event_id")

' &H0001 (hex 1), which is adCmdText, tells the
' connection object that we're sending a text command,
' which is speedier

Set rs = db.Execute(query,,&H0001)
	eventDescription = rs("event_description")
	eventRecur = rs("event_recur")
	eventTitle = rs("event_title")
	eventID = rs("event_id")
	eventDate = Request.QueryString("date")
	startTime = rs("time_start")
	endTime = rs("time_end")
	userID = rs("user_id")
rs.Close
Set rs = nothing
%>

<html>

<script language="javascript"><!--
//preload images and text for faster operation

if (document.images) {
// back to calendar icon
	var iconMonth = new Image();
	iconMonth.src = "images/icon_calprev_grey.gif";
	var iconMonthOn = new Image();
	iconMonthOn.src = "images/icon_calprev.gif"
	statusMonth = "View in Calendar";

// login icon
	var iconKey = new Image();
	iconKey.src = "images/icon_key_grey.gif";
	var iconKeyOn = new Image();
	iconKeyOn.src = "images/icon_key.gif"
	statusKey = "Login";
}

function iconOver(name){
	if (document.images) {
  		document.images[name].src=eval("icon"+name+"On.src");
		status=eval("status"+name);
	}
}

function iconOut(name){
	if (document.images) {
  		document.images[name].src=eval("icon"+name+".src");
		status="";
	}
}
//-->
</script>

<body bgcolor="#<%=color(1)%>" link="#<%=color(7)%>" vlink="#<%=color(7)%>" alink="#<%=color(6)%>">
<br>
<center>
<table border=0 cellspacing=0 cellpadding=3>
<tr>
	<td rowspan=2 align="center" valign="top">
		<font face="Tahoma, Arial, Helvatica" color="#<%=color(4)%>">
		<b><font size=2><%=WeekdayName(WeekDay(eventDate))%></font><br>
		<font size=7><%=Day(eventDate)%></font><br>
		<font size=5><%=MonthName(Month(eventDate),1)%></font></b><br>
		<font size=4><%=Year(eventDate)%></font>
		</font>
	</td>

	<td valign="top">
	<table cellspacing=0 cellpadding=3 border=0 width="100%">
	<tr>
		<td bgcolor="#<%=color(4)%>">
			<font face="Verdana, Arial, Helvetica" size=4><b>
			<a href="webCal3_<%=Request("view")%>.asp?date=<%=eventDate%>"
			<%=switchIcon("Month")%>><img name="Month" src="./images/icon_calprev_grey.gif" width=15 height=16 alt="" border=0></a>

			<%=eventTitle%></b></font>&nbsp;

<%
if Session(dataName & "User") <> userID then
	response.write "<a href=""webCal3_login.asp?url=" & Server.URLEncode(url) & """ "
%>
		<%=switchIcon("Key")%>><img name="Key" src="./images/icon_key_grey.gif" width=16 height=15 alt="" border=0></a>
<% end if %>
		</td>
<%
if eventDescription <> "" then
	response.write "<tr><td><font size=2>" _
		& Replace(eventDescription, VbCrLf, "<br>") _
		& "</font></td>"
end if
%>

</table>

<%
if eventRecur <> "none" then
	dim count, dateList()

	query = "SELECT * FROM cal_dates" _
		& " WHERE event_id=" & eventID _
		& " ORDER BY event_date"
	Set rs = db.Execute(query,,&H0001)

' generate array of event dates

	count = 0
	do while not rs.EOF
		ReDim preserve dateList(count)
		dateList(count) = rs("event_date")
		count = count + 1
		rs.MoveNext
	loop
	rs.Close
	Set rs = nothing

' if the event recurs then display the dates on
' which it occurs, invoking the include file
' that does the special formatting
%>
	<table bgcolor="#<%=color(2)%>" width="100%"><form><tr><td>
	<!--#include file="webCal3_showrecur.inc"-->
	</td></form></table>
<%
end if
response.write "</td>" & VbCrLf

'-----------------------------------
' display time range if one was entered for this event
'-----------------------------------

if startTime <> "" then
	dim hrStart, hrEnd, span, hrCurrent, textColor

' the Hour function formats to military time

	hrStart = Hour(startTime)
	hrEnd = Hour(endTime)

	response.write "<td rowspan=2 valign=""top"">" _
		& "<table cellspacing=1 cellpadding=0 border=0>"

' calculate the hours spanned by the event

	span = (hrEnd - hrStart) + 1

	for h = 0 to 23
		if h = hrStart then
			hrCurrent = "<b>" & Replace(startTime, ":00 ", " ") & "</b>"
		elseif h = hrEnd then
			hrCurrent = "<b>" & Replace(endTime, ":00 ", " ") & "</b>"
		else

' otherwise insert the array value with regular clock notation
' appended, changing 12PM to noon for the temporally challenged

			if h = 0 then
				hrCurrent = "<b>midnight</b>"
			elseif h < 12 then
				hrCurrent = h & ":00 AM"
			elseif h = 12 then
				hrCurrent = "<b>noon</b>"
			else
				hrCurrent = h - 12 & ":00 PM"
			end if
		end if

' make the hours covered by the event a different color

		if h >= hrStart AND h <= hrEnd then
			hrColor = "ffffff"
			textColor = "000000"
		else
			hrColor = color(2)
			textColor = color(5)
		end if

		response.write "<tr><td bgcolor=""#" & hrColor _
			& """ align=""right"" nowrap><font face=""Tahoma, Arial, Helvetica""" _
			& "size=1 color=""#" & textColor & """>" _
			& hrCurrent & "</font></td>"
	next
	response.write "</td></table>"
end if

' from here display the management buttons corresponding to the
' existing level of user access
%>

<!-- buttons -->

<% if CInt(Session(dataName & "User")) = CInt(userID) OR Session(dataName & "Access") = "admin" then %>

<tr>
	<td valign="bottom">

	<!-- framing table -->
	<table bgcolor="#<%=color(5)%>" width="100%" cellspacing=0 cellpadding=2 border=0><tr><td>
	<!-- end framing table -->

	<table cellspacing=0 cellpadding=2 border=0 width="100%">
	<tr>
		<td align="right" bgcolor="#<%=color(12)%>">
		<form action="webCal3_edit.asp" method="post">
		<input type="submit" name="edit" value="Edit"><br>
		<input type="submit" name="delete" value="Delete"></td>

<%
' NOTE that this excludes events which might be listed
' as recurring but now occur on just one date (count=1)

	if eventRecur <> "none" AND count > 1 then
%>

		<td bgcolor="#<%=color(12)%>"><font face="Tahoma, Arial, Helvetica" size=2>
		<input type="radio" name="scope" value="one">only this occurence<br>
		<input type="radio" name="scope" value="future">this and all future occurences<br>
		<input type="radio" name="scope" value="all" checked>all <%=count%> occurences
		</font>

<%		end if %>

		</td>

	</table>

	<!-- framing table -->
	</td></table>
	<!-- end framing table -->

	<input type="hidden" name="event_id" value="<%=eventID%>">
	<input type="hidden" name="date" value="<%=eventDate%>">
	<input type="hidden" name="count" value="<%=count%>">
	<input type="hidden" name="url" value="<%=url%>">
	<input type="hidden" name="view" value="<%=Request("view")%>">
	</form>
<%
end if
db.Close
Set db = nothing
%>
	</td>
</table>

</center>
</body>
</html>