<!--#include file="data/webCal3_data.inc"-->
<!--#include file="webCal3_themes.inc"-->

<%
' Copyright 1999 Jason Abbott (jason@webott.com)
' Last updated 06/29/1999

dim dayFirst, dayLast, d, cal, col, description, row
dim events(31), rowCurrent, rowTotal, m, y, mLoop, yLoop
dim datePrev, dateNext, userID, index

if Request.Form("logout") = "Logout" OR _
	Session(dataName & "User") = "" then
	Session(dataName & "User") = 0
	Session(dataName & "Public") = 1
	Session(dataName & "Access") = "user"
end if

if Request.Form("public") = "Hide public events" then
	Session(dataName & "Public") = 0
elseif Request.Form("public") = "Show public events" then
	Session(dataName & "Public") = 1
elseif Session(dataName & "Public") = "" then
	Session(dataName & "Public") = 1
end if

' ---------------------------------------------------------
' setup values
' ---------------------------------------------------------

' determine how this page was called and assign values
' for month and year accordingly

if Request.Form("month") <> "" then
	m = CDbl(Request.Form("month"))
	y = CDbl(Request.Form("year"))
elseif Request.QueryString("date") <> "" then
	m = Month(Request.QueryString("date"))
	y = Year(Request.QueryString("date"))
else
	m = Month(Date)
	y = Year(Date)
end if

dateNext = DateAdd("m", 1, DateSerial(y, m, 1))
datePrev = DateAdd("m", -1, DateSerial(y, m, 1))

' ---------------------------------------------------------
' build an array of event data for selected month
' ---------------------------------------------------------
' find the numeric value of the first day of the month
' ie Sunday = 1, Wednesday = 4

dayFirst = WeekDay(Dateserial(y, m, 1))

' find the last day by subtracting 1 day from the first day
' of the next month (no need for yNext here)

dayLast = Day(Dateserial(y, Month(dateNext), 1) - 1)

' now get the total for last month to write the few
' days of last month that show up on this calendar

dayLastMonth = Day(Dateserial(y, m, 1) - 1)

' find all events occuring between the first and
' last second of the selected month

query = "SELECT * FROM cal_events E INNER JOIN cal_dates D" _
 	& " ON (E.event_id = D.event_id) " _
 	& " WHERE event_date BETWEEN " & strDelim _
 	& m & "/1/" & y & strDelim _
 	& " AND " & strDelim & m & "/" & dayLast & "/" & y _
 	& strDelim & " AND (user_id = " & Session(dataName & "User")

if Session(dataName & "Public") then
	query = query & " OR private = 0)"
else
	query = query & " AND private = 1)"
end if

query = query & " ORDER BY event_date, time_start"
	
' put all matching events in an array indexed by day number

%>
<!--#include file="show_status.inc"-->
<%
' &H0001 (hex 1), which is adCmdText, tells the
' connection object that we're sending a text command,
' which is speedier

Set rs = db.Execute(query,,&H0001)
'response.write query
do while not rs.EOF
	index = Day(rs("event_date"))
	events(index) = events(index) _
		& "<img src=""./images/arrow_right_" & rs("event_color") _
		& ".gif"" width=4 height=7>" & VbCrLf _
		& "<a href=""webCal3_detail.asp?event_id=" & rs("E.event_id") _
		& "&date=" & rs("event_date") & "&view=month"" " & VbCrLf

' display time value in status bar, getting rid of seconds,
' only if time assigned to event

	if rs("time_start") <> "" then
		description = Replace(TimeValue(rs("time_start")), ":00 ", " ") _
			& " to " & Replace(TimeValue(rs("time_end")), ":00 ", " ")
	else
		description = "Click for more details"
	end if
	events(index) = events(index) _
		& showStatus(description) & ">" _
		& rs("event_title") & "</a><br>" & VbCrLf
	rs.movenext
loop

rs.Close
db.Close
set rs = nothing
set db = nothing
%>

<html>
<head>
<script language="javascript"><!--
//preload images and text for faster operation

if (document.images) {
// back icon
	var iconPrev = new Image();
	iconPrev.src = "images/icon_calprev_grey.gif";
	var iconPrevOn = new Image();
	iconPrevOn.src = "images/icon_calprev.gif"
	statusPrev = "Back to <%=MonthName(Month(datePrev))%>";

// forward icon
	var iconNext = new Image();
	iconNext.src = "images/icon_calnext_grey.gif";
	var iconNextOn = new Image();
	iconNextOn.src = "images/icon_calnext.gif"
	statusNext = "Go to <%=MonthName(Month(dateNext))%>";

// login icon	
	var iconKey = new Image();
	iconKey.src = "images/icon_key_grey.gif";
	var iconKeyOn = new Image();
	iconKeyOn.src = "images/icon_key.gif"
	statusKey = "Login";

// users icon	
	var iconUsers = new Image();
	iconUsers.src = "images/icon_users_grey.gif";
	var iconUsersOn = new Image();
	iconUsersOn.src = "images/icon_users.gif"
	statusUsers = "Manage users";
	
// print icon
	var iconPrint = new Image();
	iconPrint.src = "images/icon_print_grey.gif";
	var iconPrintOn = new Image();
	iconPrintOn.src = "images/icon_print.gif"
	statusPrint = "Make printable";
	
// search icon
	var iconSearch = new Image();
	iconSearch.src = "images/icon_search_grey.gif";
	var iconSearchOn = new Image();
	iconSearchOn.src = "images/icon_search.gif"
	statusSearch = "Find a scheduled event";

// week view icon
	var iconWeek = new Image();
	iconWeek.src = "images/week_grey.gif";
	var iconWeekOn = new Image();
	iconWeekOn.src = "images/week.gif"
	statusWeek = "View last week";

// goto icon
	var iconGoto = new Image();
	iconGoto.src = "images/icon_goto_grey.gif";
	var iconGotoOn = new Image();
	iconGotoOn.src = "images/icon_goto.gif"
	statusGoto = "Goto the selected date";
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

</head>
<body bgcolor="#<%=color(1)%>" link="#<%=color(7)%>" vlink="#<%=color(7)%>" alink="#<%=color(6)%>">

<!-- calendar table -->

<table width="100%" border=0 cellspacing=0 cellpadding=1>
<tr>
	<td><font face="Verdana, Arial, Helvatica" size=5 color="#<%=color(4)%>">
		<b><nobr><%=MonthName(m) & " " & y%></nobr></b></font>
	</td>
<form method="post" action="webCal3_month.asp">
	<td align="right" valign="bottom"><nobr>

		<a href="webCal3_month.asp?date=<%=datePrev%>"
		<%=switchIcon("Prev")%>><img name="Prev" src="./images/icon_calprev_grey.gif" width=15 height=16 alt="Previous month" border=0></a>
		&nbsp;
		<a href="webCal3_find.asp?view=month"
		<%=switchIcon("Search")%>><img name="Search" src="./images/icon_search_grey.gif"
		 width=17 height=16 alt="Find events" border=0></a>
		&nbsp;
		<a href="webCal3_month.asp?date=<%=dateNext%>"
		<%=switchIcon("Next")%>><img name="Next" src="./images/icon_calnext_grey.gif"
		 width=15 height=16 alt="Next month" border=0></a>
		&nbsp;
<%
if Session(dataName & "User") = 0 then
	' accomodate virtual host (jea:3/8/00)
	strPath = Request.ServerVariables("PATH_TRANSLATED")
	strPath = Right(strPath, Len(strPath) - InStrRev(strPath,"\"))

	response.write "<a href=""webCal3_login.asp?url=" _
		& strPath & "?" _
		& Server.URLEncode(Request.ServerVariables("QUERY_STRING")) & """ "
%>
		<%=switchIcon("Key")%>><img name="Key" src="./images/icon_key_grey.gif"
		 width=16 height=15 alt="Login" border=0></a>
		&nbsp;
<% elseif Session(dataName & "Access") = "admin" then %>
		<a href="webCal3_user-admin.asp?view=month"
		<%=switchIcon("Users")%>><img name="Users" src="./images/icon_users_grey.gif"
		 width=12 height=15 alt="User Manager" border=0></a>
		&nbsp;
<% end if%>
		<a href="webCal3_print-month.asp?date=<%=DateSerial(y, m, 1)%>" target="_top"
		<%=switchIcon("Print")%>><img name="Print" src="./images/icon_print_grey.gif"
		width=16 height=14 border=0 alt="Make printable"></a>
		&nbsp;
		<a href="javascript:document.forms[0].submit();" 
		<%=switchIcon("Goto")%>><img name="Goto" src="./images/icon_goto_grey.gif"
		 width=18 height=15 alt="Goto the selected date" border=0></a>
		
		<select name="month">
<%
' this creates the form list of month names

for mLoop = 1 to 12
	response.write "<option value='" & mLoop & "'"
	if mLoop = Month(Date) then response.write " selected"
	response.write ">" & MonthName(mLoop,1) & VbCrLf
next
%>
		</select>
		<select name="year">
<%
' this creates the form list of 20 years

for yLoop = Year(Date) - 10 to Year(Date) + 10
	response.write "<option"
	if yLoop = Year(Now) then response.write " selected"
	response.write ">" & yLoop & VbCrLf
next
%>
		</select>
	</nobr>
	</td>
	</form>
<tr>
	<td bgcolor="#<%=color(6)%>" align="center" colspan=2>
	<table width="100%" border=0 cellspacing=1 cellpadding=1>
	<tr>

<%
' print all the day names as headings

for col = 1 to 7
	response.write "   <td width=""14%"" align=""center"" bgcolor=""#" & color(2) & """>" _
		& VbCrLf & "<font face=""Verdana, Arial, Helvetica"" size=1>" _
		& WeekDayName(col,0) & "</font></td>" & VbCrLf
next

response.write "<td></td><tr>"

' ---------------------------------------------------------
' now generate calendar body
' ---------------------------------------------------------

' the column variable keeps constant track of the
' current calendar column

column = 0

font = "<font face=""Tahoma, Arial, Helvetica"" size=2><b>"
nondayFormat = "<td valign=""top"" bgcolor=""#" & color(11) _
	& """>" & font & "<font color=""#" & color(13) & """>" & VbCrLf

' cycle through all the days previous to the first
' day of the active month

for d = 1 to dayFirst - 1
	response.write nondayFormat & dayLastMonth - dayFirst + d + 1 _
		& "</b></font></font></td>" & VbCrLf
	column = column + 1
next

' now cycle through all the days of the current month

row = 1
for d = 1 to dayLast
	column = column + 1
	response.write "<td height=45 valign=""top"""
	if y & m & d = Year(now) & Month(now) & Day(now) then
		response.write " bgcolor=""#" & color(8) & """"
	elseif column = 1 or column = 7 then
		response.write " bgcolor=""" & color(10) & """"
	else
		response.write " bgcolor=""" & color(9) & """"
	end if
	response.write ">" & font & VbCrLf _
		& VbCrLf & "<a href=""webCal3_edit.asp?" _
		& "date=" & Dateserial(y, m, d) & "&view=month"" " _
		& VbCrLf _
		& showStatus("Add a new event to " & DateSerial(y, m, d)) _
		& ">" & d & "</a></b></font><br>" _
		& "<font face=""Arial, Helvetica"" size=1>" _
		& events(d) & "</font></td>" & VbCrLf
	if column = 7 AND d < dayLast then
		response.write "<td valign=""center""><a href=""webCal3_week.asp?date=" _
			& DateSerial(y, m, d) & """ " _
			& "onMouseOver=""document.images['Week" & row _
			& "'].src='images/week.gif'; " _
			& "status='View week " & row & "'; return true;"" " _
			& "onMouseOut=""document.images['Week" & row _
			& "'].src='images/week_grey.gif'; " _
			& "status=''; return true;"">" _
			& "<img name=""Week" & row _
			& """ src=""./images/week_grey.gif"" border=0></a></td><tr>" & VbCrLf
		column = 0
		row = row + 1
	end if
next

' finally, cycle through as many days of the next month as
' necessary to fill the calendar grid through column 7

if column > 0 then
	d = 1
	do while column < 7
		response.write nondayFormat & d & "</font></b></font></td>" & VbCrLf
		d = d + 1
		column = column + 1
	loop
	response.write "<td valign=""center""><a href=""webCal3_week.asp?date=" _
		& dateNext & """ " & switchIcon("Week") _
		& "><img name=""Week"" src=""./images/week_grey.gif"" border=0></a></td>"
end if
%>
	</table>
	</td>
<% if Session(dataName & "User") <> 0 then %>
<form action="webCal3_month.asp?date=<%=Dateserial(y, m, 1)%>" method="post">
<tr>
	<td colspan=2 align="right">
<%		if Session(dataName & "Public") then %>
	<input type="submit" name="public" value="Hide public events">
<%		else %>
	<input type="submit" name="public" value="Show public events">
<%		end if %>
	<input type="submit" name="logout" value="Logout">
	</td>
<% end if %>
</form>
</table>

<font face="Verdana, Arial, Helvetica" size=1>
<a href="http://webott.com/jason/webCal.html" target="_top">
webCal 3.5</a>
</font>

</body>
</html>