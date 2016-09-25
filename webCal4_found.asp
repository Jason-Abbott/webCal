<% Option Explicit %>
<% Response.Buffer = True %>
<html>
<head>
<!--#include file="./data/webCal4_data.inc"-->
<!--#include file="./include/webCal4_constants.inc"-->
<!--#include file="./include/webCal4_settings.inc"-->
<!--#include file="./include/webCal4_rollovers.inc"-->
<SCRIPT TYPE="text/javascript" LANGUAGE="javascript" SRC="./include/webCal4_buttons.js"></SCRIPT>
<!--#include file="./include/webCal4_showrecur.inc"-->
<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' Last updated 3/8/2000

dim strQuery			' query passed to database
dim oRS					' recordset object
dim oDates				' dictionary object of dates
dim intEventID			' event id
dim strTitle			' event title
dim strRecur			' event recurrence
dim strCombine			' parameter combination
dim intCount			' count of matching events
dim arDates				' dates for each event
dim strDates			' temporarily hold date list
dim arTemp				' temporary array of event details
dim x					' loop counter
dim intSrchStart		' start of search
dim intDispStart		' start of data display
dim intSrchTime			' total time of search
dim intDispTime			' total time of display

intSrchStart = milliTime()

' this subroutine converts the field information into
' proper SQL
Function parsefield (strField, s)
	dim strWhere		' WHERE part of SELECT SQL
	dim strCombine		' type of parameter combination
	dim arWords			' array of words used in search
	dim x				' loop counter

	strWhere = strWhere & strCombine & "("

	' if it begins and ends with a quote
	' use Mid to cut out what is between the quotes
	if Left(s,1) = """" and Right(s,1) = """" then
		strWhere = strWhere & strField & " LIKE '%" _
			& Mid(s, 2, Len(s)-2) & "%'"

	' or if it has a plus (+) in it
	elseif InStr(s,"+") > 0 then
		arWords = Split(s, "+")
		for x = 0 to UBound(arWords)
			strWhere = strWhere & strField & " LIKE '%" & arWords(x) & "%'"
			if x < UBound(arWords) then
				strWhere = strWhere & " AND "
			end if
		next

	' or if it starts with a minus (-)
	elseif Left(s,1) = "-" then
		strWhere = strWhere & strField & " NOT LIKE '%" _
			& Right(s, Len(s)-1) & "%'"

	' otherwise split on spaces
	else
		arWords = Split(s)
		for x = 0 to UBound(arWords)
			strWhere = strWhere & strField & " LIKE '%" & arWords(x) & "%'"
			if x < UBound(arWords) then
				strWhere = strWhere & " OR "
			end if
		next
	end if

	parsefield = strWhere & ")"
End Function

' go through all the form elements to figure out which
' ones have values to generate concise SQL

if Request.Form("title") <> "" then
	strQuery = strQuery & strCombine _
		& parsefield("event_title", Request.Form("title"))
	strCombine = " AND "
end if

if Request.Form("description") <> "" then
	strQuery = strQuery & strCombine _
		& parsefield("event_description", Request.Form("description"))
	strCombine = " AND "
end if

' if a start or end date was entered, but not both, then
' use the current date as the missing date

if Request.Form("date_start") <> "" OR Request.Form("date_end") <> "" then
	strQuery = strQuery & strCombine & "(event_date BETWEEN #"
	if Request.Form("date_start") <> "" then
		strQuery = strQuery & Request.Form("date_start")
	else
		strQuery = strQuery & Date
	end if
	strQuery = strQuery & "# AND #"
	if Request.Form("date_end") <> "" then
		strQuery = strQuery & Request.Form("date_end")
	else
		strQuery = strQuery & Date
	end if
	strQuery = strQuery & "#)"
	strCombine = " AND "
end if

' looks like a hack
if Session(unique & "User") = "" then
	Session(unique & "User") = 0
end if

' *** add permissions check ***
' now build the full query and run it
strQuery = "SELECT * FROM tblEvents E INNER JOIN tblEventDates D" _
	& " ON (E.event_id = D.event_id) " _
	& " WHERE " & strQuery _
	& " ORDER BY event_date"

response.write strQuery
response.flush

Set oRS		= Server.CreateObject("ADODB.RecordSet")
Set oDates	= Server.CreateObject("Scripting.Dictionary")

oRS.Open strQuery, g_strDSN, adOpenForwardOnly, adLockReadOnly, adCmdText

intSrchTime = milliTime() - intSrchStart
intDispStart = milliTime()

' use associative array (dictionary object) to generate
' a unique list of events
intCount = 0
do while not oRS.EOF
	intCount = intCount + 1
	intEventID = oRS("E.event_id")
	if oDates.Exists(intEventID) then
		' if the event is already in the array then we need to
		' temporarily save the already entered event oDates,
		' erase the key and then add in the old oDates plus the
		' new one

		strDates = oDates.Item(intEventID)
		oDates.Remove(intEventID)
		oDates.Add intEventID, strDates & " " & oRS("event_date")

		' event info and oDates are delimited so that we can later
		' split them out into other arrays
	else
		' otherwise we can just add a new key with the single date

 		oDates.Add intEventID, oRS("event_title") & "|" _
			& oRS("event_recur") & "|" & oRS("event_date")
'		response.write oDates.Item(intEventID) & "<br>"
	end if
	oRS.movenext
loop
oRS.Close
Set oRS = nothing

' now, after all that trouble, let's see if we ended up with
' any matches; if no matches then send back to search form
' with notice

if intCount = 0 then
	response.redirect "webCal4_find.asp?retry=1" _
		& "&view=" & Request.Form("view")
end if

' if we're still here then there must have been matches
' let's display them:
%>

</head>
<body bgcolor="#<%=g_arColor(1)%>" link="#<%=g_arColor(7)%>" vlink="#<%=g_arColor(7)%>" alink="#<%=g_arColor(6)%>">
<center>

<table border=0 cellpadding=1 cellspacing=0 width="60%">
<tr bgcolor="#<%=g_arColor(4)%>">
	<td><nobr><font face="<%=g_arFont(0)%>" size=4><b>
<%
response.write "Found " & oDates.Count & " match"
if oDates.Count > 1 or oDates.Count < 1 then response.write "es"
response.write "</b> covering " & intCount & " date"
if intCount > 1 then response.write "s"
response.write "</font></nobr></td>"

' now go through every matched event and display
' its information
for each x in oDates.Keys
	arTemp = split(oDates.Item(x), "|")
	strTitle = arTemp(0)
	strRecur = arTemp(1)
	arDates = split(arTemp(2))
	Erase arTemp

	' this creates a zero based array, the first two items being
	' the title and recur type, the rest being the event dates

	response.write "<tr><td bgcolor='#" & g_arColor(2) & "'>" _
	& "<font face='" & g_arFont(1) & "' size=2><b>" _
	& strTitle & "</b></font></td>" & vbCrLf _
	& "<tr><form><td>" & showrecur(arDates, Date, intEventID, strRecur) _
	& "</td></form>" & VbCrLf & "<tr><td>&nbsp;</td>" & VbCrLf
next
oDates.RemoveAll

response.flush
intDispTime = milliTime() - intDispStart
%>

	</td>
</table>

<table><tr><td bgcolor="#<%=g_arColor(0)%>" align="center">
<font face="<%=g_arFont(2)%>" size=1 color="#<%=g_arColor(6)%>">
Your search took <%=timer(intSrchTime)%> and the<br>
results were rendered in <%=timer(intDispTime)%>
</font>
</td></table>

</center>
</body>
</html>