<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' Last updated 3/4/2001

' variables used only by functions on this page
dim m_arRows()			' array of strings for HTML table
dim m_arGrid()			' array of event coordinates in rows and columns
dim m_arList()			' array of dictionary objects with event info
dim m_strUntimed		' untimed events

' generate table to display day's events (updated 3/4/01)
'   creates as many columns as necessary to display overlapping events
' returns string----------------------------------------------------------
Function makeDayHTML(ByVal v_strQuery)
	dim strHTML			' returned HTML string
	dim intEventCount	' zero-based count of events
	dim intEvent		' index to current event
	dim intCol			' current column
	dim intColCount		' total number of columns
	dim intRow			' current row
	dim arEvents		' 2D array of events
	
	arEvents = getDayArray(v_strQuery)
	intEventCount = UBound(arEvents)
	ReDim m_arGrid(m_arSegments(m_SEG_MAX), 0)
	ReDim m_arRows(m_arSegments(m_SEG_TOTAL))
	ReDim m_arList(0)
	Set m_arList(0) = Server.CreateObject("Scripting.Dictionary")
	
	intColCount = slotEvents(arEvents)
	Call spanEventColumns(arEvents, intColCount) : Erase arEvents
	Call spanBlankColumns(intColCount) : Erase m_arGrid
	
	strHTML = "<table width='100%' border='0' cellspacing='1' cellpadding='1'><tr>"
	
	if m_strUntimed <> "" then
		' add row to display untimed events
		strHTML = strHTML & "<td align='center' bgcolor='#" & g_arColor(4) _
			& "'><font face='" & g_arFont(0) & "' size='2'>untimed</font></td>" _
			& "<td colspan='" & intColCount + 3 & "' bgcolor='#" & g_arColor(9) _
			& "'><font face='" & g_arFont(0) & "' size='2'>" _
			& m_strUntimed & "</font></td><tr>"
	end if
	
	Call makeHourLabels(m_arRows, m_arSegments)
	
	' group into rows
	for intCol = 0 to intColCount
		for each intRow in m_arList(intCol).Keys
			m_arRows(intRow) = m_arRows(intRow) & m_arList(intCol).Item(intRow)
		next
		' release dictionary object
		Set m_arList(intCol) = nothing
	next
	Erase m_arList
	
	' final column highlights business hours
	Call makeEndColumn(m_arRows, m_arSegments) : Erase m_arSegments

	makeDayHTML = strHTML & Join(m_arRows, "<tr>") & "</table>"
End Function

' retrieve day's events from database (updated 3/2/01)
' returns array ----------------------------------------------------------
Function getDayArray(ByVal v_strQuery)
	dim strDescription	' description displayed in status bar
	dim intRowStart		' event start segment (row)
	dim intRowEnd		' event end segment
	dim arEvents		' array of events indexed by column
	dim x				' loop counters

	if v_strQuery <> g_NO_EVENTS then arEvents = getArray(v_strQuery, "")
	if IsArray(arEvents) then
		for x = 0 to UBound(arEvents, 2)
			if Trim(arEvents(g_TIME_START, x)) <> "" then
				intRowStart = getSegment(arEvents(g_TIME_START, x), m_arSegments)
				intRowEnd = getSegment(arEvents(g_TIME_END, x), m_arSegments)
				strDescription = "From " _
					& FormatDateTime(arEvents(g_TIME_START, x),3) & " to " _
					& FormatDateTime(arEvents(g_TIME_END, x),3)
			else
				' untimed event
				intRowStart = -1
				intRowEnd = -1
				strDescription = "Click for more details"
			end if
			arEvents(g_EVENT_TITLE, x) = "<a href='webCal4_detail.asp?event_id=" _
				& arEvents(g_EVENT_ID, x) & "&date=" & arEvents(g_EVENT_DATE, x) _
				& "&view=day' " & showStatus(strDescription) & ">" _
				& arEvents(g_EVENT_TITLE, x) & "</a>" & vbCrLf
			arEvents(g_TIME_START, x) = intRowStart
			arEvents(g_TIME_END, x) = intRowEnd
		next
	end if
	getDayArray = arEvents
End Function

' generate dates that define this day (updated 3/3/01)
' returns array of properties --------------------------------------------
Function getDayDates()
	dim arDates(6)
	
	if Request.QueryString("date") <> "" then
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
	
	getDayDates = arDates
End Function

' group events into columns to avoid overlap (updated 3/3/01)
' returns ----------------------------------------------------------------
Function slotEvents(ByVal v_arEvents)
	dim intCol			' current column
	dim intRow			' current row
	dim intColCount		' total columns
	dim bFit			' does event fit in this column
	dim x				' loop counter

	intColCount = 0

	for x = 0 to UBound(v_arEvents, 2)
		if v_arEvents(g_TIME_START, x) = -1 then
			' this is an untimed event
			m_strUntimed = m_strUntimed & "<img src='./images/arrow_right_" _
				& v_arEvents(g_EVENT_COLOR, x) & ".gif' width='4' height='7'>" _
				& v_arEvents(g_EVENT_TITLE, x) & "</a><br>"
		else
			intCol = 0
			do until intCol > intColCount
				bFit = true
				for intRow = v_arEvents(g_TIME_START, x) to v_arEvents(g_TIME_END, x)
					if Trim(m_arGrid(intRow, intCol)) <> "" then
						' there's already an event here--try next column
						bFit = false : exit for
					end if
				next
				if bFit then
					' the event fits--add to dictionary
					m_arList(intCol).Add v_arEvents(g_TIME_START, x), v_arEvents(g_EVENT_TITLE, x)
					' update grid--put event index in start position
					m_arGrid(v_arEvents(g_TIME_START, x), intCol) = x
					for intRow = v_arEvents(g_TIME_START, x) + 1 to v_arEvents(g_TIME_END, x) - 1
						' mark other grid positions as full
						m_arGrid(intRow, intCol) = "x"
					next 
					exit do		' move on to next event
				elseif intCol = intColCount then
					' no fit and we're out of columns--make new column
					intColCount = intColCount + 1
					ReDim Preserve m_arGrid(m_arSegments(m_SEG_MAX), intColCount)
					ReDim Preserve m_arList(intColCount)
					Set m_arList(intColCount) = Server.CreateObject("Scripting.Dictionary")
				end if
				intCol = intCol + 1
			loop
		end if
	next
	slotEvents = intColCount
End Function

' populate grid array with column span information (updated 3/2/01)
' updates page scope variables -------------------------------------------
Sub spanEventColumns(ByVal v_arEvents, ByVal v_intColCount)
	dim intCol			' current column
	dim intRow			' current row
	dim intEvent		' index into event array
	dim intRowStart		' event start time
	dim intRowEnd		' event end time
	dim intColSpan		' number of columns spanned
	dim intRowSpan		' number of rows spanned
	dim intColRemain	' number of columns after this event
	dim intRowRemain	' number of rows after this event
	dim intRatio		' columns percent of total table width
	dim bFit			' room to span?
	
	intRatio = 90 / (v_intColCount + 1)

	for intCol = 0 to v_intColCount
		for intRow = 0 to m_arSegments(m_SEG_TOTAL)
			if IsNumeric(Trim(m_arGrid(intRow,intCol))) then
			'if Trim(m_arGrid(intRow,intCol)) <> "" then
				' an event starts here--test for potential colspan
				intEvent = m_arGrid(intRow, intCol)
				intRowStart = v_arEvents(g_TIME_START, intEvent)
				intRowEnd = v_arEvents(g_TIME_END, intEvent)
				intColSpan = 1
				if intCol < v_intColCount then
					' columns remain for possible spanning
					for intColRemain = intCol + 1 to v_intColCount
						' cycle through remaining columns
						bFit = true
						for intRowRemain = intRow to intRowEnd
							' every segment covered by this event
							if m_arGrid(intRowRemain, intColRemain) = "x" then
								' already something here
								bFit = false : exit for
							end if
						next
						if bFit then
							' span this column
							intColSpan = intColSpan + 1
							for intRowRemain = intRow to intRowEnd - 1
								' mark spanned column segments as used
								m_arGrid(intRowRemain, intColRemain) = "x"
							next
						end if
					next
				end if
				intRowSpan = intRowEnd - intRowStart
				' update grid and dictionary
				m_arGrid(intRow, intCol) = "x"
				m_arList(intCol).Item(intRow) = "<td colspan='" & intColSpan & "' " _
					& "rowspan='" & intRowSpan & "' bgcolor='#" & hexColor(v_arEvents(g_EVENT_COLOR, intEvent)) _
					& "' width='" & intRatio * intColSpan & "%' align='center'>" _
					& "<font face='" & g_arFont(0) & "' size=2>" _
					& m_arList(intCol).Item(intRow) & "</font></td>"
				intRow = intRow + intRowSpan
			end if
		next
	next
End Sub

' add spanning information for blank segments (updated 3/2/01)
' updates page scope variables -------------------------------------------
Sub spanBlankColumns(ByVal v_intColCount)
	dim intRow			' current row
	dim intCol			' current column
	dim strColor
	dim intColRemain	' number of columns after this
	dim intColSpan		' number of columns spanned

	for intRow = 0 to m_arSegments(m_SEG_TOTAL)
		strColor = "f5f5f5"
		for intCol = 0 to v_intColCount
			if Trim(m_arGrid(intRow, intCol)) = "" then
				' unused segment
				intColSpan = 1
				if intCol < v_intColCount then
					' columns remain for potential spanning	
					for intColRemain = intCol + 1 to v_intColCount
						' cycle through remaining columns to find blanks
						if Trim(m_arGrid(intRow,intColRemain)) = "" then
							' the segment is unused--may be spanned
							intColSpan = intColSpan + 1
							m_arGrid(intRow, intColRemain) = "x"
						else
							exit for
						end if
					next
				end if
				' alernate color every hour
				if Int(intRow / m_arSegments(m_SEG_PER_HOUR)) Mod 2 = 0 then strColor = "dfdfdf"
				m_arGrid(intRow, intCol) = "x"
				m_arList(intCol).Add intRow, "<td colspan='" & intColSpan & "' bgcolor='#" & strColor & "'>" _
					& "<img src='./images/tiny_blank.gif'></td>"
			end if
		next
	next
End Sub

' generates small month view to select other days (updated 3/3/01)
' returns string ---------------------------------------------------------
Function makeNavMonthHTML(ByVal v_strThisDay)
	dim y, m, d			' year, month, day
	dim strHTML			' return string
	dim arDays(31)		' array of day information
	dim strQuery
	dim strLastPrevDay	' last day of previous month
	dim intCol			' running track of calendar column
	dim strNonDay		' formatting for other-month days
	dim arEvents		' array of events
	dim x				' loop counter

	intCol = 0
	strNonDay = "<td align='right'><font face='" & g_arFont(0) & "' size=1 color='#" _
			& g_arColor(11) & "'>"

	y = Year(m_arDates(g_FIRST_DATE))
	m = Month(m_arDates(g_FIRST_DATE))
	strLastPrevDay	= Day(m_arDates(g_FIRST_DATE) - 1)
	
	strQuery = makeQuery(m_arDates(g_FIRST_DATE), m_arDates(g_LAST_DATE))
	arEvents = getArray(strQuery, "")
	if IsArray(arEvents) then
		for x = 0 to UBound(arEvents, 2)
			arDays(Day(arEvents(g_EVENT_DATE, x))) = True
		next
		Erase arEvents
	end if
  
	' framing table
	strHTML = "<table bgcolor='#" & g_arColor(6) & "' cellspacing='0' cellpadding='2' border='0'>" _
		& "<tr><td>" & vbCrLf & "<table border=0 cellspacing=1 cellpadding=1 bgcolor='#" _
		& g_arColor(4) & "'><tr height=14>"

	' day of week headings
	for x = 1 to 7
		strHTML = strHTML & "<td width='14.3%' align='center'>" _
			& "<font face='" & g_arFont(0) & "' size=1 color='#ffffff'>" _
			& "<b>" & Left(WeekDayName(x),1) & "</b></font></td>" & vbCrLf
	next
	strHTML = strHTML & "</tr>"
	
	' last days of previous month
	for d = 1 to m_arDates(g_FIRST_DAY) - 1
		strHTML = strHTML & strNonDay & strLastPrevDay - m_arDates(g_FIRST_DAY) + d + 1 & "</font></td>"
		intCol = intCol + 1
	next
	
	' all days of current month
	for d = 1 to m_arDates(g_LAST_DAY)
		intCol = intCol + 1
		strHTML = strHTML & "<td align='right' bgcolor='#"
		if y & m & d = Year(now) & Month(now) & Day(now) then
			strHTML = strHTML & g_arColor(8)
		elseif arDays(d) then
			strHTML = strHTML & g_arColor(0)
		elseif intCol = 1 OR intCol = 7 then
			strHTML = strHTML & g_arColor(3)
		else
			strHTML = strHTML & g_arColor(2)
		end if
		strHTML = strHTML & "'><font face='" & g_arFont(0) & "' size=1>"
		
		if arDays(d) then
			' events are present on this day
			strHTML = strHTML & "<a href='webCal4_day.asp?date=" & DateSerial(y,m,d) _
				& "' " & showStatus("View " & DateSerial(y,m,d) & " details") & "><b>" _
				& d & "</b></a>"
		else
			strHTML = strHTML & d
		end if
		
		strHTML = strHTML & "</font></td>"
		
		if intCol = 7 AND d < m_arDates(g_LAST_DAY) then
			strHTML = strHTML & "<tr>"
			intCol = 0
		end if
	next
	
	' first days of next month
	if intCol > 0 then
		d = 1
		do while intCol < 7
			strHTML = strHTML & strNonDay & d & "</font></td>"
			d = d + 1
			intCol = intCol + 1
		loop
	end if

	makeNavMonthHTML = strHTML & "</table></td></table>"
End Function
%>
