<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' Last updated 3/4/2000

' defines values for the month calendar

dim strFirstDate	' first date in month
dim strLastDate		' last date in month
dim strSelect		' selected date to display
dim strView			' text displayed at top of calendar
dim strPrev			' the date of the last day before this month
dim strNext			' the date of the first day after this month

strView = "month"

' retrieve selected date
if Request.Form("month") <> "" then
	strSelect = DateSerial(Request.Form("year"), Request.Form("month"), 1)
elseif Request.QueryString("date") <> "" then
	strSelect = Request.QueryString("date")
else
	strSelect = Date
end if

strFirstDate	= Dateserial(Year(strSelect), Month(strSelect), 1)
strNext			= DateAdd("m", 1, strFirstDate)
strPrev			= DateAdd("m", -1, strFirstDate)
strLastDate		= DateAdd("d", -1, strNext)
%>