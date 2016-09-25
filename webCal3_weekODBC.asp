<!--#include file="data/webCal3_data.inc"-->
<!--#include file="webCal3_themes.inc"-->
<%
' Copyright 1999 Jason Abbott (jason@webott.com)
' Last updated 06/04/1999

dim dayFirst, dayLast, daySelect, dayNames, d, col
dim events(7), dayShow, prevFirst, nextFirst

if Request.Form("logout") = "Logout" OR _
	Session(dataName & "User") = "" then
	Session(dataName & "User") = 0
	Session(dataName & "Public") = 1
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

if Request.Form("month") <> "" then
	daySelect = Dateserial(Request.Form("year"), Request.Form("month"), 1)
	dayFirst = DateAdd("d", 1 - WeekDay(daySelect), daySelect)
	dayLast = DateAdd("d", 6, dayFirst)
elseif Request.QueryString("date") <> "" then
	daySelect = Request.QueryString("date")
	dayFirst = DateAdd("d", 1 - WeekDay(daySelect), daySelect)
	dayLast = DateAdd("d", 6, dayFirst)
else
	dayFirst = DateAdd("d", 1 - WeekDay(Date), Date)
	dayLast = DateAdd("d", 7 - WeekDay(Date), Date)
end if

prevFirst = DateAdd("d", -7, dayFirst)
nextFirst = DateAdd("d", 1, dayLast)

' ---------------------------------------------------------
' build an array of event data for selected week
' ---------------------------------------------------------

query = "SELECT * FROM cal_events E INNER JOIN cal_dates D" _
	& " ON (E.event_id = D.event_id) " _
	& " WHERE event_date BETWEEN " & strDelim _
	& Month(dayFirst) & "/" & Day(dayFirst) & "/" & Year(dayFirst) _
	& strDelim & " AND " & strDelim _
	& Month(dayLast) & "/" & Day(dayLast) & "/" & Year(dayLast) _
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
	index = WeekDay(rs("event_date"))
	events(index) = events(index) _
		& "<img src=""./images/arrow_right_" & rs("event_color") _
		& ".gif"" width=4 height=7>" & VbCrLf _
		& "<a href=""webCal3_detail.asp?event_id=" & rs("event_id") _
		& "&date=" & rs("event_date") & "&view=week"" " & VbCrLf

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
	statusPrev = "Last week";

// forward icon
	var iconNext = new Image();
	iconNext.src = "images/icon_calnext_grey.gif";
	var iconNextOn = new Image();
	iconNextOn.src = "images/icon_calnext.gif"
	statusNext = "Next week";

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
	
// goto icon
	var iconGoto = new Image();
	iconGoto.src = "images/icon_goto_grey.gif";
	var iconGotoOn = new Image();
	iconGotoOn.src = "images/icon_goto.gif"
	statusGoto = "Goto the first week of the selected month";
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

<!-- heading table -->

<table width="100%" border=0 cellspacing=0 cellpadding=1>
<tr>
	<td>
	<font face="Verdana, Arial, Helvatica" size=5 color="#<%=color(4)%>">
	<b><nobr>Week View</nobr></b></font>
	</td>
<form method="post" action="webCal3_week.asp">
	<td align="right" valign="bottom"><nobr>

		<a href="webCal3_week.asp?date=<%=prevFirst%>"
		<%=switchIcon("Prev")%>><img name="Prev" src="./images/icon_calprev_grey.gif" width=15 height=16 alt="" border=0></a>
		&nbsp;
		<a href="webCal3_find.asp?view=week"
		<%=switchIcon("Search")%>><img name="Search" src="./images/icon_search_grey.gif"
		 width=17 height=16 alt="" border=0></a>
		&nbsp;
		<a href="webCal3_week.asp?date=<%=nextFirst%>"
		<%=switchIcon("Next")%>><img name="Next" src="./images/icon_calnext_grey.gif"
		 width=15 height=16 alt="" border=0></a>
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
		 width=16 height=15 alt="" border=0></a>
		&nbsp;
<% elseif Session(dataName & "Access") = "admin" then %>
		<a href="webCal3_user-admin.asp?view=week"
		<%=switchIcon("Users")%>><img name="Users" src="./images/icon_users_grey.gif"
		 width=12 height=15 alt="" border=0></a>
		&nbsp;
<% end if%>
		<a href="webCal3_print-week.asp?date=<%=dayFirst%>" target="_top"
		<%=switchIcon("Print")%>><img name="Print" src="./images/icon_print_grey.gif"
		 width=16 height=14 border=0 alt="Make printable"></a>
		&nbsp;
		<a href="javascript:document.forms[0].submit();" 
		<%=switchIcon("Goto")%>><img name="Goto" src="./images/icon_goto_grey.gif"
		 width=18 height=15 alt="Goto the first week of the selected month" border=0></a>		
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
<%
' generate the heading

dayShow = dayFirst
dim spanArray(1)
m = 0
for col = 1 to 7
	dayNames = dayNames & "<td width=""14.3%"" align=""center"" bgcolor=""#" _
 		& color(2) & """><font face=""Verdana, Arial, Helvetica"" size=1>" _
		& WeekDayName(col,0) & "</font></td>"

	spanArray(m) = spanArray(m) + 1
	dayNext = DateAdd("d", 1, dayShow)
	if Month(dayNext) <> Month(dayShow) then
		m = 1
	end if
	dayShow = dayNext
next
%>
<tr>
	<td bgcolor="#<%=color(2)%>" colspan=<%=spanArray(0)%> align="center">
	<font face="Tahoma, Arial, Helvetica" size=2 color="#<%=color(6)%>">
	<a href="webCal3_month.asp?date=<%=dayFirst%>" <%=showStatus("View all of " & MonthName(Month(dayFirst)))%>>
	<b><%=MonthName(Month(dayFirst)) & " " & Year(dayFirst)%></b></a></font></td>
<% if spanArray(0) < 7 then %>
	<td bgcolor="#<%=color(2)%>" colspan=<%=spanArray(1)%> align="center">
	<font face="Tahoma, Arial, Helvetica" size=2 color="#<%=color(6)%>">
	<a href="webCal3_month.asp?date=<%=dayLast%>" <%=showStatus("View all of " & MonthName(Month(dayLast)))%>>
	<b><%=MonthName(Month(dayLast)) & " " & Year(dayLast)%></b></a></font></td>
<% end if %>	
	
<tr><%=dayNames%>
<tr>

<%
dayShow = dayFirst
for col = 1 to 7
%>
	<td height=200 valign="top"
<%	if dayShow = Date then %>
	bgcolor="#<%=color(8)%>"
<%	elseif col = 1 or col = 7 then %>
	bgcolor="#<%=color(10)%>"
<%	else %>
	bgcolor="#<%=color(9)%>"
<%	end if %>
	><font face="Tahoma, Arial, Helvetica" size=2><b>
	<a href="webCal3_edit.asp?date=<%=dayShow%>&view=week"
	<%=showStatus("Add a new event to " & dayShow)%>>
	<%=Day(dayShow)%></a></b></font>
	<br>
	<font face="Arial, Helvetica" size=1>
	<%=events(col)%></font></td>
<%
	dayShow = DateAdd("d", 1, dayShow)
next
%>

	</table>
	</td>
<% if Session(dataName & "User") <> 0 then %>
<form action="webCal3_week.asp?date=<%=dayFirst%>" method="post">
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
</table>

<font face="Verdana, Arial, Helvetica" size=1>
<a href="http://webott.com/jason/webCal.html" target="_top">
webCal 3.5</a>
</font>

</body>
</html>