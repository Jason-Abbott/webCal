<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' Last updated 3/3/2001

' variables used only by functions on this page
dim m_bCanUserAdd		' can user add events

' generate dates that define this month (updated 3/1/01)
' returns array of properties --------------------------------------------
Function getMonthDates()
	dim arDates(6)
		
	' retrieve selected date
	if Request.Form("month") <> "" then
		arDates(g_THIS_DATE) = DateSerial(Request.Form("year"), Request.Form("month"), 1)
	elseif Request.QueryString("date") <> "" then
		arDates(g_THIS_DATE) = Request.QueryString("date")
	else
		arDates(g_THIS_DATE) = Date
	end if

	arDates(g_FIRST_DATE) = Dateserial(Year(arDates(g_THIS_DATE)), Month(arDates(g_THIS_DATE)), 1)
	arDates(g_NEXT_DATE) = DateAdd("m", 1, arDates(g_FIRST_DATE))
	arDates(g_PREV_DATE) = DateAdd("m", -1, arDates(g_FIRST_DATE))
	arDates(g_LAST_DATE) = DateAdd("d", -1, arDates(g_NEXT_DATE))
	arDates(g_FIRST_DAY) = WeekDay(arDates(g_FIRST_DATE))
	arDates(g_LAST_DAY) = Day(arDates(g_LAST_DATE))
	
	getMonthDates = arDates
End Function

' put all matching events in an array indexed by day number (updated 3/2/01)
' returns array ----------------------------------------------------------
Function getMonthArray(ByVal v_strQuery)
	dim arDay(31)			' array of events for each day of month
	dim arEvents			' array of events from query
	dim intIndex			' day of month index into array
	dim strDescription		' event description
	dim x					' loop counter
	
	if v_strQuery <> g_NO_EVENTS then arEvents = getArray(v_strQuery, "")
	if IsArray(arEvents) then
		for x = 0 to UBound(arEvents, 2)
			intIndex = Day(arEvents(g_EVENT_DATE, x))
			arDay(intIndex) = arDay(intIndex) _
				& "<img src='./images/arrow_right_" & arEvents(g_EVENT_COLOR, x) _
				& ".gif' width=4 height=7> " _
				& "<a href='webCal4_detail.asp?event_id=" & arEvents(g_EVENT_ID, x) _
				& "&date=" & arEvents(g_EVENT_DATE, x) & "&view=month' "
				
			' display time value in status bar, getting rid of seconds,
			' only if time assigned to event
			if arEvents(g_TIME_START, x) <> "" then
				strDescription = Replace(TimeValue(arEvents(g_TIME_START, x)), ":00 ", " ") _
					& " to " & Replace(TimeValue(arEvents(g_TIME_END, x)), ":00 ", " ")
			else
				strDescription = "Click for more details"
			end if
			arDay(intIndex) = arDay(intIndex) _
				& showStatus(strDescription) & ">" _
				& arEvents(g_EVENT_TITLE, x) & "</a><br>"
		next
		Erase arEvents
	end if

	getMonthArray = arDay
End Function

' this generates the body of the month calendar (updated 3/14/01)
' returns string ---------------------------------------------------------
Function makeMonthHTML(ByVal v_strQuery) 
	dim y, m, d			' year, month, day
	dim intCol			' current calendar column
	dim intRow			' current calendar row
	dim strLastPrevDay	' last day of previous month
	dim strHTML			' hold all this stuff
	dim arEvents
	
	strLastPrevDay	= Day(m_arDates(g_FIRST_DATE) - 1)
	arEvents = getMonthArray(v_strQuery)
	y = Year(m_arDates(g_FIRST_DATE))
	m = Month(m_arDates(g_FIRST_DATE))
	strHTML = "<table width='100%' border='0' cellspacing='1' cellpadding='0'><tr>" & vbCrLf

	' print all the day names as headings
	for intCol = 1 to 7
		strHTML = strHTML & "<td width='14%' class='dayName'>" _
			& WeekDayName(intCol,0) & "</td>"
	next
	
	strHTML = strHTML & "<td></td><tr>"
	
	' last days of previous month
	intCol = 0
	for d = 1 to m_arDates(g_FIRST_DAY) - 1
		strHTML = strHTML & "<td valign='top' class='dayOther'>" _
			& strLastPrevDay - m_arDates(g_FIRST_DAY) + d + 1 & "</td>"
		intCol = intCol + 1
	next
	
	' cycle through all the days of the current month
	intRow = 1
	for d = 1 to m_arDates(g_LAST_DAY)
		intCol = intCol + 1
		strHTML = strHTML & "<td height='45' valign='top' class='"
		if y & m & d = Year(now) & Month(now) & Day(now) then
			strHTML = strHTML & "dayThis'>"
		elseif intCol = 1 or intCol = 7 then
			strHTML = strHTML & "dayWeekend'>"
		else
			strHTML = strHTML & "dayCommon'>"
		end if

		' make day number a link if user can add events
		strHTML = strHTML & "<table width='100%' cellspacing='0' " _
			& "cellpadding='0' border='0'><tr><td class='dayAdd'>"
		if canUserAdd() then
			strHTML = strHTML & "<a href='webCal4_event-edit.asp?" _
				& "date=" & Dateserial(y, m, d) & "&view=month' " _
				& showStatus("Add a new event to " & DateSerial(y, m, d)) _
				& ">" & d & "</a>"
		else
			strHTML = strHTML & d
		end if
		strHTML = strHTML & "</td>"
		
		' if the day contains events then generate link to day detail
		if arEvents(d) <> "" then
			'strHTML = strHTML & " <a href='webCal4_day.asp?date=" _
			'	& DateSerial(y, m, d) & "' " & switchIcon("Day" & d, "Day", "View " & DateSerial(y, m, d) & " detail") _
			'	& "><img name='Day" & d & "' src='./images/day_grey.gif' border=0></a>"
			strHTML = strHTML & "<td class='dayLink'><a href='webCal4_day.asp?date=" _
				& DateSerial(y, m, d) & "'>Show Day</a></td>"
		end if

		strHTML = strHTML & "</table><div class='eventTitle'>" _
			& arEvents(d) & "</div></td>"
		if intCol = 7 AND d <= m_arDates(g_LAST_DAY) then
		
			' if we're at the last column then generate link to week view
			strHTML = strHTML & "<td valign='center'><a href='webCal4_week.asp?date=" _
				& DateSerial(y, m, d) & "' " & switchIcon("Week" & intRow, "Week", "View week " & intRow) _
				& "><img name='Week" & intRow _
				& "' src='./images/week_grey.gif' border=0></a></td>"
			
			' only start a new row if days of the month remain
			if d < m_arDates(g_LAST_DAY) then strHTML = strHTML & "<tr>"
			strHTML = strHTML
			intCol = 0 : intRow = intRow + 1
		end if
	next
	
	' first days of next month
	if intCol > 0 then
		d = 1
		do while intCol < 7
			strHTML = strHTML & "<td valign='top' class='dayOther'>" & d & "</td>"
			d = d + 1 : intCol = intCol + 1
		loop
		strHTML = strHTML & "<td valign='center'><a href='webCal4_week.asp?date=" _
			& m_arDates(g_NEXT_DATE) & "' " & switchIcon("Week", "Week", "View week " & intRow) _
			& "><img name='Week' src='./images/week_grey.gif' border=0></a></td>"
	end if
	
	strHTML = strHTML & "</table>"
	
	makeMonthHTML = strHTML
End Function

' determine whether user has sufficient permissions to add events (updated 3/3/01)
' returns boolean --------------------------------------------------------
Function canUserAdd()
	dim bCanAdd
	dim arGroups
	dim x
	
	if IsEmpty(m_bCanUserAdd) then
		' hasn't been checked yet
		arGroups = Session(g_unique & "Groups")
		for x = 0 to UBound(arGroups, 2)
			if arGroups(g_GROUP_ACCESS, x) > g_READ_ACCESS then
				bCanAdd = true : exit for
			end if
		next
		m_bCanUserAdd = bCanAdd
	else
		' use already determined value
		bCanAdd = m_bCanUserAdd
	end if
	
	canUserAdd = bCanAdd
End Function
%>