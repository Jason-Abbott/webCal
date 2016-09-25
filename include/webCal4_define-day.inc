<%
' Copyright 1999 Jason Abbott (webcal@webott.com)
' Last updated 10/27/1999

strCommon = "BETWEEN " & strDelim _
	& strThisDay & " 12:00:00 AM" & strDelim & " AND " _
	& strDelim & strThisDay & " 11:59:59 PM" & strDelim

' ---------------------------------------------------------
' build an array of event data for the selected day
' ---------------------------------------------------------
' first select the untimed program events

strQuery = "SELECT * FROM tblPrograms WHERE (StartDate " _
	& strCommon & " OR EndDate " & strCommon & ")" _
	& strProgQuery

Set rsPrograms = Server.CreateObject("ADODB.RecordSet")

' adOpenStatic = 3
' adLockReadOnly = 1
' adCmdText = &H0001

rsPrograms.Open strQuery, g_strDSN, 3, 1, &H0001
do while not rsPrograms.EOF

' determine if the program ends or begins in this week

	if DateValue(rsPrograms("StartDate")) = DateValue(strThisDay) then
		strName = "Start of " & rsPrograms("ProgramName")
	else
		strName = "End of " & rsPrograms("ProgramName")
	end if

	strPrograms = strPrograms _
		& "<img src=""./images/arrow_right_black" _
		& ".gif"" width=4 height=7> " & VbCrLf _
		& "<a href=""webCal4_detail.asp?program_id=" & rsPrograms("ProgramID") _
		& "&view=week"" " _
		& showStatus(rsPrograms("Participants")) & ">" _
		& strName & "</a><br>" & VbCrLf
	rsPrograms.MoveNext
loop
rsPrograms.Close
Set rsProgams = nothing

' now build the grid for timed events

strQuery = "SELECT * FROM tblEvents WHERE (Date " _
	& strCommon & ")" & strEventQuery & "ORDER BY Start"

Set rsEvents = Server.CreateObject("ADODB.RecordSet")
rsEvents.Open strQuery, g_strDSN, 3, 1, &H0001
intCount = CInt(rsEvents.RecordCount) - 1

ReDim arTemp(intCount,4)

' 0 = description
' 1 = start time
' 2 = column
' 3 = time segments spanned
' 4 = color

dim arLast()
ReDim arLast(0)
arLast(0) = 0

' this keeps track of the last used time in each column

x = 0
do while not rsEvents.EOF

	strDescription = "From " _
		& FormatDateTime(rsEvents("Start"),3) & " to " _
		& FormatDateTime(rsEvents("End"),3)
		
	arTemp(x,0) = "<a href=""webCal4_detail.asp?event_id=" _
		& rsEvents("EventID") & "&date=" & rsEvents("Date") _
		& "&view=day"" " & showStatus(strDescription) & ">" _
		& rsEvents("Title") & "</a>" & VbCrLf

	arTemp(x,1) = segments(rsEvents("Start"))

	arTemp(x,2) = ""

	intEnd = segments(rsEvents("End"))
	
	arTemp(x,3) = intEnd - arTemp(x,1)

	arTemp(x,4) = g_arColor(0)
	
' go through each column to find the first fit
' update arLast to new end segment for that column
	
	for y = 0 to UBound(arLast)
		if arTemp(x,1) > arLast(y) then
			arTemp(x,2) = y
			arLast(y) = intEnd
			exit for
		end if
	next
	
	if arTemp(x,2) = "" then
	
' there was no room in any existing columns
' so add a new column and set end segment
	
		y = UBound(arLast) + 1
		arTemp(x,2) = y
		ReDim Preserve arLast(y)
		arLast(y) = intEnd
	end if

	x = x + 1
	rsEvents.MoveNext
loop
rsEvents.Close
set rsEvents = nothing

' get the total number (ZERO-BASED) of columns

intCols = UBound(arLast)

dim arEvents()
ReDim arEvents(intCols,intTotal,2)

' DIMENSIONS
' 0 = column
' 1 = time segment of day
' 2 = event information

for y = 0 to intCols

' go through each column

	for x = 0 to intCount
	
' go through each event
	
		if arTemp(x,2) = y then
		
' the event occurs in this column

			arEvents(y,arTemp(x,1),0) = arTemp(x,0)
			arEvents(y,arTemp(x,1),1) = arTemp(x,3)
			arEvents(y,arTemp(x,1),2) = arTemp(x,4)
		end  if
	next
next
%>