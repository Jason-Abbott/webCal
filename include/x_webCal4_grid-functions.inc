<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' Last updated 3/3/2001

' grid segments
Const m_SEG_PER_HOUR = 3	' number of segments (cells) per hour
Const m_SEG_MAX = 4			' max segments per day (mins./day/intErval)
Const m_SEG_TOTAL = 5		' segments between start and end
Const m_SEG_HEIGHT = 6		' height in pixels

' get settings for week or day grid (updated 2/28/01)
' returns array ----------------------------------------------------------
Function getGridProperties(ByVal v_arSegDefault)
	dim arSegments(6)	' array of segments properties
	
	response.write v_arSegDefault(g_SEG_MINS) : response.flush
	arSegments(g_SEG_MINS) = v_arSegDefault(g_SEG_MINS)
	arSegments(g_SEG_START) = v_arSegDefault(g_SEG_START)
	arSegments(g_SEG_END) = v_arSegDefault(g_SEG_END)
	Erase v_arSegDefault
	arSegments(m_SEG_PER_HOUR) = 60 / arSegments(g_SEG_MINS)

	if arSegments(g_SEG_START) = 0 Or arSegments(g_SEG_END) = 0 then
		' set defaults--begin at 6:00 AM (6:00 hours)
		arSegments(g_SEG_START) = 6 * arSegments(m_SEG_PER_HOUR)
		' end at 9:00 PM (22:00 hours)
		arSegments(g_SEG_END) = 22 * arSegments(m_SEG_PER_HOUR) - 1
	end if
	
	' total minutes in a day (1440) / minutes in segment
	arSegments(m_SEG_MAX) = (1440 / arSegments(g_SEG_MINS)) - 1
	arSegments(m_SEG_TOTAL) = arSegments(g_SEG_END) - arSegments(g_SEG_START)
	' pixel height for each segment
	arSegments(m_SEG_HEIGHT) = 24 / arSegments(m_SEG_PER_HOUR) - 1
	
	getGridProperties = arSegments
End Function

' Convert time to segment based on specified interval (updated 2/28/01)
' returns integer --------------------------------------------------------
Function getSegment(ByVal v_strTime, ByVal v_arSegments)
	dim intMin		' minutes past the hour of start time
	dim intAdd		'
	dim intSegment
	dim x			' loop counter

	intMin = Minute(v_strTime)
	intAdd = v_arSegments(m_SEG_PER_HOUR)/2
	for x = 0 to 60/v_arSegments(m_SEG_PER_HOUR) - 1
		if intMin < intAdd then
			intSegment = x
			exit for
		end if
		intAdd = intAdd + v_arSegments(m_SEG_PER_HOUR)
	next
	getSegment = intSegment + (Hour(v_strTime) * v_arSegments(m_SEG_PER_HOUR)) - v_arSegments(g_SEG_START)
End Function

' Match color to color name for grid display
' returns string----------------------------------------------------------
Function hexColor(ByVal v_strNameClr)
	dim strColor
	select case v_strNameClr
		case "blue"
			strColor = "aaaaff"
		case "purple"
			strColor = "aa88aa"
		case "red"
			strColor = "ff8888"
		case "green"
			strColor = "99dd99"
		case "orange"
			strColor = "ffcc88"
		case else
			strColor = "ffffff"
	end select
	hexColor = strColor
End Function

' create the hour labels for grid views (updated 2/28/01)
' updates row array ------------------------------------------------------
Sub makeHourLabels(ByRef r_arRows, ByVal v_arSegments)
	dim intHour		' hour of day corresponding to time segment
	dim strHour		' hour text
	dim intRow		' current row or segment

	for intRow = 0 to v_arSegments(m_SEG_TOTAL) step v_arSegments(m_SEG_PER_HOUR)
		' step through segments by hour
		intHour = (intRow + v_arSegments(g_SEG_START)) / v_arSegments(m_SEG_PER_HOUR)
		if intHour = 0 then
			strHour = "<b>midnight</b>"
		elseif intHour < 12 then
			strHour = intHour & ":00 AM"
		elseif intHour = 12 then
			strHour = "<b>noon</b>"
		else
			strHour = intHour - 12 & ":00 PM"
		end if
	
		r_arRows(intRow) = "<td rowspan='" & v_arSegments(m_SEG_PER_HOUR) _
			& "' bgcolor='#" & g_arColor(2) _
			& "' align='right' width='10%'><font face='" & g_arFont(0) _
			& "' size='2'><nobr>" & strHour & "</nobr></font></td>"
		'x = x + 1
	next
End Sub

' Create final column to highlight business hours (updated 2/28/01)
' updates row array ------------------------------------------------------
Sub makeEndColumn(ByRef r_arRows, ByVal v_arSegments)
	dim intHour
	dim intRow
	dim strColor
	
	for intRow = 0 to v_arSegments(m_SEG_TOTAL) step v_arSegments(m_SEG_PER_HOUR)
		intHour = (intRow + v_arSegments(g_SEG_START)) / v_arSegments(m_SEG_PER_HOUR)
		strColor = "f5f5f5"
		if intHour > 12 then intHour = intHour - 12
		if intHour Mod 2 = 0 then strColor = "dfdfdf"
		r_arRows(intRow) = r_arRows(intRow) & "<td bgcolor='#" & strColor _
			& "' rowspan='" & v_arSegments(m_SEG_PER_HOUR) & "' width='1%' align='center'>" _
			& "<font face='" & g_arFont(0) & "' size=1>" & intHour & "</font></td>"
	next
		
	for intRow = 0 to v_arSegments(m_SEG_TOTAL)
		intHour = Int((intRow + v_arSegments(g_SEG_START)) / v_arSegments(m_SEG_PER_HOUR))
		if intHour < 8 OR intHour > 16 OR (intHour < 13 AND intHour > 11) then
			strColor = g_arColor(1)
		else
			strColor = g_arColor(6)
		end if
		r_arRows(intRow) = r_arRows(intRow) & "<td bgcolor='#" & strColor & "' height=" _
			& v_arSegments(m_SEG_HEIGHT) & " width='1%'>" _
			& "<img src='./images/tiny_blank.gif' height=1></td>" & vbCrLf
	next
End Sub
%>