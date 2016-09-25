<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' Last updated 2/28/2001

Const m_EVENT_TITLE = 0
Const m_TIME_START = 1
Const m_TIME_END = 2
Const m_EVENT_COLOR = 3
Const m_ALIGNMENT = 4

' variables used only by functions on this page
dim m_arRows()			' rows of HTML table

' generates complete week table (updated 2/28/01)
' returns string ---------------------------------------------------------
Function makeWeekHTML(ByVal v_strQuery)
	
	dim strHTML				' build HTML to return
	dim arWeekDay			' array of daily arrays of event arrays
	dim arUntimed			' array of daily untimed events
	dim arEvents			' array of events for this week
	dim intDaysInWeek		' five or seven day week
	dim strPrevDate			' date from previous week
	dim strNextDate			' date in next week
	dim bUntimed			' are there any untimed events
	dim x
	
	intDaysInWeek = m_arDates(g_LAST_DAY) - m_arDates(g_FIRST_DAY)
	ReDim m_arRows(m_arSegments(m_SEG_TOTAL))
	Call makeHourLabels(m_arRows, m_arSegments)
	arEvents = getWeekArray(v_strQuery)
	if IsArray(arEvents) then
		' assign events to weekdays or untimed
		bUntimed = slotEvents(arEvents, intDaysInWeek, arUntimed, arWeekDay)
		Call makeDayColumns(intDaysInWeek, arWeekDay)
		Erase arEvents
	else
		' create empty column for each day of week
		for x = 0 to intDaysInWeek
			Call makeBlankSegments(0, m_arSegments(m_SEG_TOTAL))
		next
		bUntimed = false
	end if
	Call makeEndColumn(m_arRows, m_arSegments)
	
	strHTML = "<table width='100%' border='0' cellspacing='1' cellpadding='0'>" _
		& "<tr><td rowspan=2 width='10%'>" _
		& "<img src='./images/tiny_blank.gif' height=1></td>"
		
	strHTML = strHTML & makeWeekDayHeaders(intDaysInWeek, arWeekDay, arUntimed)
	
	' collapse row array into string
	strHTML = strHTML & Join(m_arRows, "<tr>") : Erase m_arRows
	
	if bUntimed then
		strHTML = strHTML & makeUntimedRow(arUntimed, intDaysInWeek)
		Erase arUntimed : Erase arWeekDay
	end if
	
	makeWeekHTML = strHTML & "</table>"
End Function

' generate dates that define this week (updated 2/28/01)
' returns array of properties --------------------------------------------
Function getWeekDates()
	dim arDates(6)
	
	if Request.QueryString("date") <> "" then
		arDates(g_THIS_DATE) = Request.QueryString("date")
	else
		arDates(g_THIS_DATE) = Date
	end if
	
	' define first and last days to display
	if Session(g_unique & "Weekends") then
		arDates(g_FIRST_DAY) = 1 : arDates(g_LAST_DAY) = 7
	else
		arDates(g_FIRST_DAY) = 2 : arDates(g_LAST_DAY) = 6
	end if
	
	arDates(g_FIRST_DATE) = DateAdd("d", arDates(g_FIRST_DAY) - WeekDay(arDates(g_THIS_DATE)), arDates(g_THIS_DATE))
	arDates(g_LAST_DATE) = DateAdd("d", arDates(g_LAST_DAY) - 1, arDates(g_FIRST_DATE))
	arDates(g_PREV_DATE) = DateAdd("d", -arDates(g_LAST_DAY), arDates(g_FIRST_DATE))
	arDates(g_NEXT_DATE) = DateAdd("d", arDates(g_FIRST_DAY), arDates(g_FIRST_DATE))
	
	'r_intDaysInWeek = m_arDates(g_LAST_DAY) - m_arDates(g_FIRST_DAY)
	getWeekDates = arDates
End Function

' retrieves week's events from database (updated 2/28/01)
' returns array ----------------------------------------------------------
Function getWeekArray(ByVal v_strQuery)
	dim strDescription	' description displayed in status bar
	dim intRowStart		' event start segment (row)
	dim intRowEnd		' event end segment
	dim arEvents 		' array of events
	dim x				' loop counter

	if v_strQuery <> g_sNO_EVENTS then arEvents = getArray(v_strQuery, "")
	if IsArray(arEvents) then
		' events match the query
		for x = 0 to UBound(arEvents, 2)
			if Trim(arEvents(g_TIME_START, x)) <> "" then
				intRowStart = getSegment(arEvents(g_TIME_START, x), m_arSegments)
				intRowEnd = getSegment(arEvents(g_TIME_END, x), m_arSegments)
				strDescription =  Replace(TimeValue(arEvents(g_TIME_START, x)), ":00 ", " ") _
					& " to " & Replace(TimeValue(arEvents(g_TIME_END, x)), ":00 ", " ")
			else
				' untimed event
				intRowStart = -1
				intRowEnd = -1
				strDescription = "Click for more details"
			end if
		
			arEvents(g_EVENT_TITLE, x) = "<a href='webCal4_detail.asp?event_id=" _
				& arEvents(g_EVENT_ID, x) & "&date=" & arEvents(g_EVENT_DATE, x) _
				& "&view=week' " & showStatus(strDescription) & ">" _
				& arEvents(g_EVENT_TITLE, x) & "</a>"
			arEvents(g_TIME_START, x) = intRowStart
			arEvents(g_TIME_END, x) = intRowEnd
			' Weekday is 0 based so subtract start day
			arEvents(g_EVENT_DATE, x) = WeekDay(arEvents(g_EVENT_DATE, x)) - m_arDates(g_FIRST_DAY)
		next
	end if
	getWeekArray = arEvents
End Function

' build HTML for row of untimed events (updated 2/28/01)
' returns string ---------------------------------------------------------
Function makeUntimedRow(ByVal v_arUntimed, ByVal v_intDaysInWeek)
	dim strHTML
	dim strColor
	dim intDayOfWeek
	
	strHTML = strHTML & "<tr><td align='center' bgcolor='#" & g_arColor(4) _
		& "'><font face='" & g_arFont(0) & "' size='2'>untimed</font></td>"
	for intDayOfWeek = 0 to v_intDaysInWeek
		if v_arUntimed(intDayOfWeek) <> "" then
			strColor = g_arColor(9)
		else
			strColor = g_arColor(6)
		end if
		strHTML = strHTML & "<td ' bgcolor='#" & strColor _
			& "'><font face='" & g_arFont(2) & "' size='1'>" _
			& v_arUntimed(intDayOfWeek) & "</font></td>"
	next
	makeUntimedRow = strHTML & "<td colspan='1'></td>"
End Function

' add event HTML to days of week (updated 2/28/01)
' updates row array by reference -----------------------------------------
Sub makeDayColumns(ByVal v_intDaysInWeek, ByVal v_arWeekDay)
	
	dim intDayOfWeek
	dim intRow				' current row
	dim intRowStart			' first event segment
	dim intRowEnd			' last event segment
	dim intRowSpan			' getSegment spanned by event
	dim x					' loop counter
	
	for intDayOfWeek = 0 to v_intDaysInWeek
		intRow = 0
		if IsArray(v_arWeekDay(intDayOfWeek)(0)) then
			' there are events on this day
			for x = 0 to UBound(v_arWeekDay(intDayOfWeek))
				' go through each event array
				intRowStart = v_arWeekDay(intDayOfWeek)(x)(m_TIME_START)
				intRowEnd = v_arWeekDay(intDayOfWeek)(x)(m_TIME_END)
				intRowSpan = intRowEnd - intRowStart
				Call makeBlankSegments(intRow, intRowStart - 1)
				m_arRows(intRowStart) = m_arRows(intRowStart) & "<td rowspan='" _
					& intRowSpan & "' bgcolor='#" _
					& v_arWeekDay(intDayOfWeek)(x)(m_EVENT_COLOR) _
					& "' align='center'><font face='" & g_arFont(2) & "' size=1>" _
					& v_arWeekDay(intDayOfWeek)(x)(m_EVENT_TITLE) & "</font></td>"
				' keep track of the current row (segment) of the day
				intRow = intRowEnd
			next
		end if
		' make blank getSegment to fill out rest of day
		Call makeBlankSegments(intRow, m_arSegments(m_SEG_TOTAL))
	next
End Sub

' create day headings (updated 2/28/01)
' returns string ---------------------------------------------------------
Function makeWeekDayHeaders(ByVal v_intDaysInWeek, ByVal v_arWeekDay, ByVal v_arUntimed)
	dim intDayWidthPercent
	dim intDayOfWeek
	dim strThisDay
	dim strNextDay
	dim strColor
	dim strDayHeading
	dim strHTML
	dim intMonth
	dim arSpan()			' days spanned by month; cells spanned by event
	
	intDayWidthPercent = Round(90/(v_intDaysInWeek + 1), 0)
	strThisDay = m_arDates(g_FIRST_DATE)
	ReDim arSpan(1)
	intMonth = 0
	
	for intDayOfWeek = m_arDates(g_FIRST_DAY) to m_arDates(g_LAST_DAY)
		if strThisDay = Date then
			' highlight current day
			strColor = g_arColor(8)
		elseif intDayOfWeek = 1 or intDayOfWeek = 7 then
			' shade weekend
			strColor = g_arColor(10)
		else
			strColor = g_arColor(9)
		end if
		strDayHeading = strDayHeading & "<td width='" & intDayWidthPercent _
			& "%' bgcolor='#" & strColor & "'>" _
			& "<table cellspacing='0' cellpadding='1' border='0' width='100%'><tr>" _
			& "<td align='center' bgcolor='#000000' width='20%' valign='top'>" _
			& "<font face='" & g_arFont(0) & "' color='#ffffff' size='2'><b>" _
			& Day(strThisDay) & "</b></font></td>" _
			& "<td align='center' rowspan='2'>" _
			& "<font face='" & g_arFont(0) & "' size='1'>" _
			& "<a href='webCal4_event-edit.asp?action=new&view=week&date=" & strThisDay & "' " _
			& showStatus("Add a new event to " & strThisDay) & ">" _
			& WeekDayName(intDayOfWeek, 0) & "</a></td>" _
			& "<tr>"
	
		if hasEvents(intDayOfWeek - m_arDates(g_FIRST_DAY), v_arWeekDay, v_arUntimed) then
			' this day has events--display the link to day detail
			strDayHeading = strDayHeading & "<td>" _
				& "<a href='webCal4_day.asp?date=" & strThisDay & "' " _
				& switchIcon("Day" & intDayOfWeek, "Day", "View " & strThisDay & " detail") _
				& "><img name='Day" & intDayOfWeek & "' src='./images/day_grey.gif' width='18' " _
				& "height='7' border='0'>"
		else
			' insert equivalent spacer
			strDayHeading = strDayHeading & "<td bgcolor='#000000'>" _
				& "<img src='./images/tiny_blank.gif' height='7'>"
		end if
		
		strDayHeading = strDayHeading & "</td></table></td>"
		
		' increment the day span value for current month
		arSpan(intMonth) = arSpan(intMonth) + 1
		' get the next date to show
		strNextDay = DateAdd("d", 1, strThisDay)
		' if the next day begins a new month then increment month
		if Month(strNextDay) <> Month(strThisDay) then intMonth = 1
		strThisDay = strNextDay
	next
	
	strHTML = strHTML & linkMonth(m_arDates(g_FIRST_DATE), arSpan(0))
	
	if arSpan(0) < v_intDaysInWeek + 1 then
		' if week spans two months, display second month
		strHTML = strHTML & linkMonth(m_arDates(g_LAST_DATE), arSpan(1))
	end if
	
	strHTML = strHTML & "<td rowspan='2'><img src='./images/tiny_blank.gif' height=1>" _
		& "</td><tr>" & vbCrLf & vbCrLf & strDayHeading _
		& "<td colspan='2'><img src='./images/tiny_blank.gif' height=1></td><tr>"
	
	makeWeekDayHeaders = strHTML
End Function

' does given day of week have events (updated 3/1/01)
' returns boolean --------------------------------------------------------
Function hasEvents(ByVal v_intDayOfWeek, ByVal v_arWeekDay, ByVal v_arUntimed)
	dim bEvents
	if IsArray(v_arWeekDay) then
		if IsArray(v_arWeekDay(v_intDayOfWeek)(0)) or v_arUntimed(v_intDayOfWeek) <> "" then
			bEvents = true
		else
			bEvents = false
		end if
	else
		bEvents = false
	end if
	hasEvents = bEvents
End Function

' slot events by day of week (updated 2/28/01)
' returns boolean and updates values by reference ------------------------
Function slotEvents(ByVal v_arEvents, ByVal v_intDaysInWeek, ByRef r_arUntimed, _
	ByRef r_arWeekDay)
	
	dim intDayOfWeek
	dim intRowStart		' first event segment
	dim intRowEnd		' last event segment
	dim arTemp			' array of daily event arrays
	dim bUntimed		' are there any untimed events?
	dim strColor
	dim strDescription
	dim x, y			' loop counters
	
	bUntimed = false
	ReDim r_arWeekDay(v_intDaysInWeek)
	ReDim r_arUntimed(v_intDaysInWeek)
	
	for intDayOfWeek = 0 to v_intDaysInWeek
		Redim arTemp(0) : y = 0 : intRowEnd = 0
		for x = 0 to UBound(v_arEvents, 2)
			if v_arEvents(g_EVENT_DATE, x) = intDayOfWeek then
				' event occurs on this day
				if v_arEvents(g_TIME_START, x) = -1 then
					' untimed event--add to untimed array
					r_arUntimed(intDayOfWeek) = r_arUntimed(intDayOfWeek) _
						& "<img src='./images/arrow_right_" _
						& v_arEvents(g_EVENT_COLOR, x) & ".gif' width=4 height=7>" _
						& v_arEvents(g_EVENT_TITLE, x) & "</a><br>"
					bUntimed = true
				elseif v_arEvents(g_TIME_START, x) < intRowEnd then
					' this event overlaps the previous one
					' this can only be true if there WAS a previous event
					intRowStart = arTemp(y-1)(m_TIME_START)
					if v_arEvents(g_TIME_END, x) > arTemp(y-1)(m_TIME_END) then
						' the new event lasts longer
						intRowEnd = v_arEvents(g_TIME_END, x)
					else
						intRowEnd = arTemp(y-1)(m_TIME_END)
					end if
					strColor = g_arColor(6)
					if arTemp(y-1)(m_ALIGNMENT) = "center" then
						' this is the first overlap--initialize table
						strDescription = "<font face='" & g_arFont(0) & "' " _
							& "color='#ffffff' size=2><b>" _
							& "<a href='webCal4_day.asp?date=" _
							& DateAdd("d", intDayOfWeek, m_arDates(g_FIRST_DATE)) & "' " _
							& showStatus("Zoom in on overlapping events") & ">" _
							& "overlap</a></b></font>" _
							& "<table width='100%' height='100%' border='0' " _
							& "cellspacing='2' cellpadding='1'><tr>" _
							& "<td bgcolor='#" & hexColor(arTemp(y-1)(m_EVENT_COLOR)) _
							& "' align='center'><font face='" & g_arFont(2) _
							& "' size='1'>"	& arTemp(y-1)(m_EVENT_TITLE) & "</font></td>"
					else
						' re-open the table for more additions
						strDescription = Replace(arTemp(y-1)(m_EVENT_TITLE),"</table>","")
					end if
					strDescription = strDescription & "<tr><td bgcolor='#" _
						& hexColor(v_arEvents(g_EVENT_COLOR, x)) & "' align='center'><font face='" _
						& g_arFont(2) & "' size='1'>" & v_arEvents(g_EVENT_TITLE, x) _
						& "</font></td></table>"
						
					arTemp(y-1) = Array(strDescription, intRowStart, intRowEnd, strColor,"left")
					strDescription = ""
				else						
					' add to normal day array (timed event with no overlap)
					intRowEnd = v_arEvents(g_TIME_END, x)
					ReDim Preserve arTemp(y)
					arTemp(y) = Array(v_arEvents(g_EVENT_TITLE, x), v_arEvents(g_TIME_START, x), _
						intRowEnd, hexColor(v_arEvents(g_EVENT_COLOR, x)),"center")
					y = y + 1
				end if
			end if
		next
		r_arWeekDay(intDayOfWeek) = arTemp
	next
	Erase arTemp
	slotEvents = bUntimed
End Function

' create link to months at top of week view (updated 2/28/01)
' returns string ---------------------------------------------------------
Private Function linkMonth(ByVal v_strDate, ByVal v_intSpan)
	linkMonth = vbCrLf & "<td bgcolor='#" & g_arColor(2) & "' colspan=" _
		& v_intSpan & " align='center'><font face='" & g_arFont(0) _
		& "' size=2 color='#" & g_arColor(6) & "'>" _
		& "<a href='webCal4_month.asp?date=" & v_strDate & "' " _
		& showStatus("View all of " & MonthName(Month(v_strDate))) _
		& "><b>" & MonthName(Month(v_strDate)) & " " & Year(v_strDate) _
		& "</b></a></font></td>"
End Function

' create blank column segment (updated 2/28/01)
' updates array ----------------------------------------------------------
Function makeBlankSegments(ByVal v_intFirst, ByVal v_intLast)
	
	dim strColor	' hour color
	dim x			' loop counter

	for x = v_intFirst to v_intLast
		' alternate hour color
		if Int(x / m_arSegments(m_SEG_PER_HOUR)) Mod 2 = 0 then
			strColor = "dfdfdf"	
		else
			strColor = "f5f5f5"
		end if
		m_arRows(x) = m_arRows(x) & "<td bgcolor='#" & strColor & "'>" _
			& "<img src='./images/tiny_blank.gif' height='" _
			& m_arSegments(m_SEG_HEIGHT) & "'></td>"	
	next
End Function
%>