<%
' Copyright 1996-2004 Jason Abbott (webcal@webott.com)
' methods common to grid display in various views

Class wcGrid
	
	Private m_aColumnsPerDay()			' number of event columns within each day column, zero-based
	Private m_aGrid						' segment properties in grids
	Private m_aEvents					' array of events for this view
	Private m_bAltStyle					' handles div widths differently in style sheets
	Private m_oLayout
	
	Public Property Get columnsPerDay()
		columnsPerDay = m_aColumnsPerDay
	End Property
	
	Public Property Let columnsPerDay(ByVal v_lElements)
		ReDim m_aColumnsPerDay(v_lElements)
	End Property
	
	Private Sub Class_Initialize()
		m_bAltStyle = MatchesOne(Session(g_sDB_NAME & "Browser")(g_BROWSER_ID), Array(g_BROSWER_NS, g_BROWSER_MOZILLA), true)
		Set m_oLayout = New wcLayout
	End Sub
	
	Private Sub Class_Terminate()
		Set m_oLayout = nothing
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		initialize()
	'	Purpose: 	create grid property, day columns and event arrays
	' Modifications:
	'	Date:		Name:	Description:
	'	10/14/03	JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub initialize(ByVal v_aDates, ByVal v_lView)
		m_aGrid = getGridProperties(v_lView)
		m_aEvents = getEvents(v_aDates)
		Call normalizeColumnsPerDay()
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		writeEvents()
	'	Purpose: 	write event grid within any view
	' Modifications:
	'	Date:		Name:	Description:
	'	10/9/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub writeEvents(ByVal v_lDaysInView, ByVal v_lDayWidthPercent)
		dim aSegment			' array of events in column in day
		dim lDayWidthPercent	' table percentage width per day
		dim lEventWidthPercent	' table percentage width per column
		dim lDayColumns			' event columns in current day
		dim lDay				' current day in view
		dim lColumn				' current column within day
		dim lSegment			' current segment within column and day
		dim lHour				' hour in loop
		dim sClass				' style sheet class
		dim sRootClass			' style sheet stuff
		dim bEvents				' are there events this week
		dim bNewRow				' starting new row
		dim bEndRow				' end of event grid
		dim x
		
		bEvents = IsArray(m_aEvents)
		v_lDaysInView = v_lDaysInView + 1	' convert to 1-based count
		
		for x = 0 to ((m_aGrid(g_SEG_TOTAL) + 1) * v_lDaysInView) - 1
			' every segment across every day
			lDay = x Mod v_lDaysInView
			lSegment = Int(x / v_lDaysInView)
			lHour = Int((lSegment - 1) / m_aGrid(g_SEG_PER_HOUR)) + g_GRID_START_HOUR
			sRootClass = IIf(CBool(lHour Mod 2 = 0), "Even", "Odd")
			sClass = sRootClass & IIf(isBusinessHour(lHour, lDay, v_lDaysInView), "Biz", "")
			bNewRow = CBool(x Mod v_lDaysInView = 0)
			bEndRow = CBool(x > 0 And ((x + 1) Mod v_lDaysInView = 0))
			
			if lSegment = 0 then
				' untimed event row
				with response
					if x = 0 then .write "<tr><td class='untimed'></td>"
					.write "<td class='untimed' colspan='"
					.write m_aColumnsPerDay(lDay) + 1
					.write "'></td>"
					if bEndRow then .write "<td class='untimed'></td>"
				end with
			else
			
				if bNewRow then Call writeHourTD(lHour, m_aGrid(g_SEG_PER_HOUR), lSegment)
				
				if bEvents then
					lColumn = 0
					aSegment = m_aEvents(lDay, lSegment)
					lDayColumns = m_aColumnsPerDay(lDay)
					
					if IsArray(aSegment) then
						' column array exists in this segment
						lEventWidthPercent = Int(v_lDayWidthPercent/(UBound(aSegment) + 1))
						for lColumn = 0 to UBound(aSegment)
							if IsArray(aSegment(lColumn)) then
								' event exists at this segment-column position
								Call writeEventTD(aSegment(lColumn), lEventWidthPercent)
							elseif aSegment(lColumn) <> g_SEG_SPANNED then
								Call writeBlankTD(aSegment(lColumn), m_aGrid(g_SEG_HEIGHT), sClass, lSegment)
							end if
						next
					else
						' empty segments were assigned colspan value
						Call writeBlankTD(aSegment, m_aGrid(g_SEG_HEIGHT), sClass, lSegment)
					end if
				else
					' no events
					Call writeBlankTD(0, m_aGrid(g_SEG_HEIGHT), sClass, lSegment)
				end if
				
				if bEndRow then Call writeRowEndTD(lHour, m_aGrid(g_SEG_PER_HOUR), lSegment, sRootClass)
			end if
		next
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		writeHourTD()
	'	Purpose: 	write HTML for hour label
	' Modifications:
	'	Date:		Name:	Description:
	'	10/8/03		JEA		Creation
	'-------------------------------------------------------------------------
	Private Sub writeHourTD(ByVal v_lHour, ByVal v_lRowSpan, ByVal v_lSegment)
		with response
			.write "<tr>"
			if (v_lSegment - 1) Mod v_lRowSpan = 0 then
				.write "<td class='rowEnd hourLeft' rowspan='"
				.write v_lRowSpan
				.write "'>"
				if v_lHour = 0 then
					.write "12<sup class='minute'>00</sup>"
				elseif v_lHour < 12 then
					.write v_lHour
					.write "<sup class='minute'>00</sup>" ' AM"
				elseif v_lHour = 12 then
					.write "<b>noon</b>"
				else
					.write v_lHour - 12
					.write "<sup class='minute'>00</sup>" ' PM"
				end if
				.write "</td>"
			end if
		end with
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		writeEventTD()
	'	Purpose: 	write HTML for single event
	' Modifications:
	'	Date:		Name:	Description:
	'	10/7/03		JEA		Creation
	'-------------------------------------------------------------------------
	Private Sub writeEventTD(ByVal v_aEvent, ByVal v_lWidthPercent)
		with response
			.write "<td class='event'"
			if v_aEvent(g_EVENT_SEG_SPAN) > 1 then
				.write " rowspan='"
				.write v_aEvent(g_EVENT_SEG_SPAN)
				.write "'"
			end if
			if v_aEvent(g_EVENT_COL_SPAN) > 1 then
				'v_lWidthPercent = Int(v_lWidthPercent / v_aEvent(g_EVENT_COL_SPAN))
				.write " colspan='"
				.write v_aEvent(g_EVENT_COL_SPAN)
				.write "' width='"
				.write v_lWidthPercent
				.write "'"
			end if
			.write "><img src='./images/tiny_blank.gif'>"
			.write "<div class='eventContainer' style='height: "
			.write ((m_aGrid(g_SEG_HEIGHT) + 1) * v_aEvent(g_EVENT_SEG_SPAN))
			.write "px;"
			if m_bAltStyle then
				.write " width: "
				.write v_lWidthPercent - 1
				.write "%;"
			end if
			.write "' onMouseOver='expandEvent(this);' onMouseOut='collapseEvent(this);' "
			.write " onClick='alert(""detail"");'><div class='event'>"
			Call m_oLayout.writeSymbol(g_CHAR_RECUR, "1.5em")
			.write v_aEvent(g_EVENT_TITLE)
			.write "<p>"
			.write v_aEvent(g_EVENT_DESC)
			.write "<br>"
			.write v_aEvent(g_TIME_START)
			.write "</div></div></td>"
		end with
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		writeBlankTD()
	'	Purpose: 	write a blank TD
	' Modifications:
	'	Date:		Name:	Description:
	'	10/8/03		JEA		Creation
	'-------------------------------------------------------------------------
	Private Sub writeBlankTD(ByVal v_lColSpan, ByVal v_lHeight, ByVal v_sClass, ByVal v_lSegment)
		with response
			.write "<td"
			if v_lColSpan > 1 then
				.write " colspan='"
				.write v_lColSpan
				.write "'"
			end if
			.write " class='"
			.write v_sClass
			.write " segment' style='height: "
			.write v_lHeight
			.write "px; font-size: 6pt; text-align: center;'>"
			.write "<img src='./images/tiny_blank.gif' height='"
			.write m_aGrid(g_SEG_HEIGHT)
			.write "'></td>"
		end with
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		writeRowEndTD()
	'	Purpose: 	write TD at end of grid row
	' Modifications:
	'	Date:		Name:	Description:
	'	10/9/03		JEA		Creation
	'-------------------------------------------------------------------------
	Private Sub writeRowEndTD(ByVal v_lHour, ByVal v_lRowSpan, ByVal v_lSegment, ByVal v_sClass)
		with response
			if (v_lSegment - 1) Mod v_lRowSpan = 0 then
				.write "<td class='rowEnd "
				.write v_sClass
				.write " hourRight' rowspan='"
				.write v_lRowSpan
				.write "'>"
				.write IIf((v_lHour <= 12), v_lHour, v_lHour - 12)
				.write "</td>"
			end if
			.write "<td class='"
			.write IIf(isBusinessHour(v_lHour, 1, 1), "biz", "spacer")
			.write "'><img src='./images/tiny_blank.gif' width='3' height='"
			.write m_aGrid(g_SEG_HEIGHT)
			.write "'></td>"
		end with
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		isBusinessHour()
	'	Purpose: 	determine if given hour is during business day
	'	Return: 	boolean
	' Modifications:
	'	Date:		Name:	Description:
	'	10/14/03	JEA		Creation
	'-------------------------------------------------------------------------
	Private Function isBusinessHour(ByVal v_lHour, ByVal v_lDay, ByVal v_lDaysInView)
		isBusinessHour = CBool((v_lHour >= g_BIZ_START And v_lHour < g_BIZ_END And v_lHour <> 12) _
			And (v_lDaysInView = 1 Or (v_lDay > 0 And v_lDay < 6)))
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		getEvents()
	'	Purpose: 	retrieve week's events from database
	'				structure is aEvent(lDay, lSegment)(lColumn)(EVENT_PROPERTY)
	'	Return: 	array
	' Modifications:
	'	Date:		Name:	Description:
	' 	2/28/01		JEA		Creation
	'	9/30/03		JEA		Altered data structure to allow multiple columns/day
	'-------------------------------------------------------------------------
	Private Function getEvents(ByVal v_aDates)
		dim aData			' returned from db
		dim aEvents			' array of events for each day of week
		dim aSegment		' array of event columns within a segment
		dim lRowStart		' event start segment (row)
		dim lRowEnd			' event end segment
		dim lRowSpan		' segments spanned
		dim lPrevRowEnd		' previous event end segment
		dim lDaysInView		' number of days displayed in given view
		dim lDay			' day in event array
		dim lNewDay
		dim lColumn			' current column within day (allows overlapping events)
		dim lDayColumns		' number of columns within day
		dim oData
		dim oSession
		dim bTimed			' is event timed
		dim x, y			' loop counters
	
		Set oData = New wcData : Set oSession = New wcSession
		aData = oData.GetArray(oSession.getViewQuery(v_aDates(g_FIRST_DATE), v_aDates(g_LAST_DATE)))
		Set oSession = Nothing : Set oData = Nothing

		if IsArray(aData) then
			lDay = -1		' force into day change condition
			lDaysInView = v_aDates(g_LAST_DAY) - v_aDates(g_FIRST_DAY)
			ReDim aEvents(lDaysInView, m_aGrid(g_SEG_TOTAL))
			
			for x = 0 to UBound(aData, 2)
				'if x > 5 then exit for
			
				lNewDay = Weekday(aData(g_EVENT_DATE, x)) - Weekday(v_aDates(g_FIRST_DATE))
				bTimed = Not IsVoid(aData(g_TIME_START, x))
				if lDay < lNewDay then
					' new day
					if x > 0 then m_aColumnsPerDay(lDay) = lDayColumns
					lDay = lNewDay
					lDayColumns = 0
					lPrevRowEnd = 0
				end if
				
				if bTimed then
					lRowStart = getSegment(aData(g_TIME_START, x))
					lRowEnd = getSegment(aData(g_TIME_END, x)) - 1
				else
					' untimed event
					lRowStart = 0
					lRowEnd = 0
				end if
				
				lRowSpan = (lRowEnd - lRowStart) + 1
				aSegment = aEvents(lDay, lRowStart)
				
				if IsArray(aSegment) then
					' segment already has event(s), figure out any overlap
					lColumn = firstAvailableColumn(aEvents, lDay, lRowStart, lRowEnd, lDayColumns)
					if lColumn > lDayColumns then lDayColumns = lColumn
					if UBound(aSegment) < lColumn then Redim Preserve aSegment(lColumn)
				else
					' initialize column array to hold event inside segment
					lColumn = 0
					aSegment = Array(0)
				end if
				
				' put event in segment column
				aSegment(lColumn) = transferEvent(aData, x, bTimed, lRowSpan)
				aEvents(lDay, lRowStart) = aSegment
				
				' mark spanned segments (rows) -- column spanning is calculated later
				for y = lRowStart + 1 to lRowEnd
					aSegment = aEvents(lDay, y)
					if Not IsArray(aSegment) then
						ReDim aSegment(lColumn)
					elseif Ubound(aSegment) < lColumn then
						ReDim Preserve aSegment(lColumn)
					end if
					aSegment(lColumn) = g_SEG_SPANNED
					aEvents(lDay, y) = aSegment
				next
				lPrevRowEnd = lRowEnd	' used to compare in lext loop
			next
			m_aColumnsPerDay(lDay) = lDayColumns
		end if
		If IsArray(aEvents) then Call computeSpanning(aEvents)
		getEvents = aEvents
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		firstAvailableColumn()
	'	Purpose: 	find first column available for event
	' Modifications:
	'	Date:		Name:	Description:
	' 	10/16/03	JEA		Creation
	'-------------------------------------------------------------------------
	Private Function firstAvailableColumn(ByVal v_aEvents, ByVal v_lDay, _
		ByVal v_lRowStart, ByVal v_lRowEnd, ByVal v_lDayColumns)
		dim aSegment
		dim lSegment
		dim lColumn
		dim bCanFit
		
		bCanFit = false
		for lColumn = 0 to v_lDayColumns
			for lSegment = v_lRowStart to v_lRowEnd
				aSegment = v_aEvents(v_lDay, lSegment)
				if IsArray(aSegment) then
					if UBound(aSegment) >= lColumn then
						if IsArray(aSegment(lColumn)) Or aSegment(lColumn) = g_SEG_SPANNED then
							bCanFit = false
							exit for
						end if
					end if
				end if
				bCanFit = true
			next
			if bCanFit then exit for
		next
		firstAvailableColumn = lColumn
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		computeSpanning()
	'	Purpose: 	add column spanning values to segments
	' Modifications:
	'	Date:		Name:	Description:
	' 	10/7/03		JEA		Creation
	'-------------------------------------------------------------------------
	Private Sub computeSpanning(ByRef r_aEvents)
		dim lDayColumns		' total columns in day
		dim lSpanRemain		' span remaining from last event
		dim lSpanPrevious	' previous event column span
		dim lDay			' weekday in loop
		dim lSegment		' segment in loop
		dim lColumn			' column in loop
		dim lSegColumns		' columns in given segment
		dim lDayCount		' total days in view
		dim lSegCount		' total segments in day
		dim aSegment		' segment array of columns
		
		lDayCount = UBound(r_aEvents)
		lSegCount = UBound(r_aEvents, 2)
		
		for lDay = 0 to lDayCount
			' go through each day
			lDayColumns = m_aColumnsPerDay(lDay)
			lSpanPrevious = 0
			
			for lSegment = 0 to lSegCount
				' go through each segment in day
				aSegment = r_aEvents(lDay, lSegment)
				if IsArray(aSegment) then		' contains event
					lSegColumns = UBound(aSegment)
					lSpanRemain = 0

					' resize segment array to match total for day
					if lSegColumns < lDayColumns then Redim Preserve aSegment(lDayColumns)
					
					for lColumn = 0 to lDayColumns
						' go through each column in segment
						if IsArray(aSegment(lColumn)) then
							' event start contains array of event details
							' NOTE that r_aEvents is not yet updated with Redimmed aSegment
							lSpanRemain = getEventColSpan(r_aEvents, lDay, lSegment, lColumn, lDayColumns)
							aSegment(lColumn)(g_EVENT_COL_SPAN) = lSpanRemain
							lSpanRemain = lSpanRemain - 1
							lSpanPrevious = lSpanRemain
							
						elseif aSegment(lColumn) = g_SEG_SPANNED then
							' row spanned by event--set local span to last event column span value
							lSpanRemain = lSpanPrevious
							
						elseif lSpanRemain > 0 then
							' newly created column spanned by last event
							aSegment(lColumn) = g_SEG_SPANNED
							lSpanRemain = lSpanRemain - 1
						else
							' newly created column
							lSpanRemain = getBlankColSpan(aSegment, lColumn)
							aSegment(lColumn) = lSpanRemain
							lSpanRemain = lSpanRemain - 1
						end if
					next
				else
					' blank segment spans all columns (make one-based)
					aSegment = lDayColumns + 1
					lSpanPrevious = 0
				end if
				r_aEvents(lDay, lSegment) = aSegment
			next
		next
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		getEventColSpan()
	'	Purpose: 	get maximum columns that this event can span
	'	Return: 	integer
	' Modifications:
	'	Date:		Name:	Description:
	' 	10/7/03		JEA		Creation
	'-------------------------------------------------------------------------
	Private Function getEventColSpan(ByVal v_aEvents, ByVal v_lWeekday, ByVal v_lSegment, _
		ByVal v_lColumn, ByVal v_lDayColumns)
		dim aSegment		' array of segment columns
		dim aEvent			' array of single event properties
		dim lColSpan		' total columns that can be spanned by event
		dim lSegColumns		' columns in given segment
		dim lThisSegment	' index of segment in loop
		dim lThisColumn		' index of column in loop
		
		lColSpan = 0
		aEvent = v_aEvents(v_lWeekday, v_lSegment)(v_lColumn)
		
		if v_lColumn < v_lDayColumns then
			' potentially spannable columns exist beyond event start column
			for lThisSegment = v_lSegment to v_lSegment + (aEvent(g_EVENT_SEG_SPAN) - 1)
				' go through every segment (row) covered by event
				aSegment = v_aEvents(v_lWeekday, lThisSegment)
				lSegColumns = UBound(aSegment)
				
				if lSegColumns = v_lColumn then
					' subsequent columns are undefined for this segment, hence spannable
					lColSpan = getUpdatedSpan(lColSpan, v_lDayColumns - v_lColumn)
				else
					for lThisColumn = v_lColumn + 1 to lSegColumns
						' check every column next to event
						if IsArray(aSegment(lThisColumn)) then
							' an event starts here, can't span
							lColSpan = getUpdatedSpan(lColSpan, (lThisColumn - v_lColumn) - 1)
							exit for
						elseif aSegment(lThisColumn) = g_SEG_SPANNED then
							' another event already spans this segment
							lColSpan = getUpdatedSpan(lColSpan, (lThisColumn - v_lColumn) - 1)
							exit for
						else
							lColSpan = getUpdatedSpan(lColSpan, lThisColumn - v_lColumn)
						end if
					next
				end if
				if lColSpan = 0 then exit for
			next
		end if
		getEventColSpan = lColSpan + 1		' convert to 1-based
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		getBlankColSpan()
	'	Purpose: 	count contiguous blank columns
	'	Return: 	integer
	' Modifications:
	'	Date:		Name:	Description:
	' 	10/8/03		JEA		Creation
	'-------------------------------------------------------------------------
	Private Function getBlankColSpan(ByVal v_aSegment, ByVal v_lColumn)
		dim lColSpan		' total columns that can be spanned by event
		dim lThisColumn		' index of column in loop
		
		lColSpan = 0
		for lThisColumn = v_lColumn + 1 to UBound(v_aSegment)
			if IsVoid(v_aSegment(lThisColumn)) then
				lColSpan = lColSpan + 1
			else
				exit for
			end if
		next
		getBlankColSpan = lColSpan + 1		' convert to 1-based
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		getUpdatedSpan()
	'	Purpose: 	apply updated span value
	'	Return: 	integer
	' Modifications:
	'	Date:		Name:	Description:
	' 	10/7/03		JEA		Creation
	'-------------------------------------------------------------------------
	Private Function getUpdatedSpan(ByVal v_lColSpan, ByVal v_lProposed)
		if v_lColSpan = 0 Or v_lColSpan > (v_lProposed) then
			' no spanning defined yet or existing spanning is too much
			getUpdatedSpan = v_lProposed
		else
			getUpdatedSpan = v_lColSpan
		end if
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		getGridProperties()
	'	Purpose: 	get grid properties for any view
	'	Returns:	array
	' Modifications:
	'	Date:		Name:	Description:
	'	2/28/01		JEA		Creation
	'-------------------------------------------------------------------------
	Private Function getGridProperties(ByVal v_lView)
		Const MINS_PER_DAY = 1440
		Const MINS_PER_HOUR = 60
		Const HOURS_PER_DAY = 24
		dim aDefaults		' default properties
		dim aSegments(6)	' array of segment properties
		
		aDefaults = Session(g_sDB_NAME & "Segments")(v_lView)
		aSegments(g_SEG_MINS) = aDefaults(g_SEG_MINS)
		aSegments(g_SEG_START) = 0 'v_aSegDefault(g_SEG_START)
		aSegments(g_SEG_END) = aDefaults(g_SEG_END)
		aSegments(g_SEG_PER_HOUR) = MINS_PER_HOUR / aSegments(g_SEG_MINS)
	
		' one-base; zero-position reserved for untimed events
		if aSegments(g_SEG_START) = 0 Or aSegments(g_SEG_END) = 0 then
			' set defaults--begin at 6:00 AM (6:00 hours) and end at 9:00 PM (22:00 hours)
			aSegments(g_SEG_START) = (g_GRID_START_HOUR * aSegments(g_SEG_PER_HOUR)) + 1
			aSegments(g_SEG_END) = ((g_GRID_END_HOUR + 1) * aSegments(g_SEG_PER_HOUR))
		end if
		
		' total minutes in a day (1440) / minutes in segment
		aSegments(g_SEG_MAX) = (MINS_PER_DAY / aSegments(g_SEG_MINS)) - 1
		aSegments(g_SEG_TOTAL) = aSegments(g_SEG_END) - aSegments(g_SEG_START) + 1
		' pixel height for each segment
		aSegments(g_SEG_HEIGHT) = HOURS_PER_DAY / aSegments(g_SEG_PER_HOUR) - 1
		getGridProperties = aSegments
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		getSegment()
	'	Purpose: 	Convert time to segment based on specified interval
	'	Returns:	integer (one-based)
	' Modifications:
	'	Date:		Name:	Description:
	'	2/28/01		JEA		Creation
	'	10/13/03	JEA		Simplified logic
	'-------------------------------------------------------------------------
	Private Function getSegment(ByVal v_dtTime)
		dim lMinute		' minutes past the hour of start time
		dim lAdd		'
		dim x			' loop counter
	
		lMinute = Minute(v_dtTime)
		lAdd = 0
		for x = 0 to 60 step (60 / m_aGrid(g_SEG_PER_HOUR))
			if x >= lMinute then exit for
			lAdd = lAdd + 1
		next
		getSegment = lAdd + (Hour(v_dtTime) * m_aGrid(g_SEG_PER_HOUR)) - m_aGrid(g_SEG_START) + 2
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		needNewColumn()
	'	Purpose: 	is new column needed for event; update grid values
	'	Return: 	boolean
	' Modifications:
	'	Date:		Name:	Description:
	' 	10/6/03		JEA		Creation
	'-------------------------------------------------------------------------
	Private Function needNewColumn(ByRef r_lColumn, ByVal v_aSegment)
		dim bNeedColumn
		dim x				' loop counter
	
		bNeedColumn = true
		for x = 0 to r_lColumn
			if IsVoid(v_aSegment(r_lColumn)) then
				if bNeedColumn then
					r_lColumn = x
					bNeedColumn = false
					if x > 0 then
						' check for previous events with column spanning
						v_aSegment(x - 1)
					end if
				end if
			end if
		next
		needNewColumn = bNeedColumn
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		transferEvent()
	'	Purpose: 	transfer event to column array, part of full grid array
	'	Return: 	array
	' Modifications:
	'	Date:		Name:	Description:
	' 	9/29/03		JEA		Creation
	'-------------------------------------------------------------------------
	Private Function transferEvent(ByVal v_aEvent, ByVal x, ByVal v_bTimed, ByVal v_lRowSpan)
		dim aEvent()
		dim y

		' transfer event details from db into column array
		Redim aEvent(g_EVENT_UBOUND)
		for y = g_EVENT_ID to g_EVENT_SKIP_WE
			aEvent(y) = v_aEvent(y, x)
		next
		
		' insert generated values ...
		aEvent(g_EVENT_SEG_SPAN) = v_lRowSpan
	
		' generate mouse-over description
		if v_bTimed then
			aEvent(g_EVENT_MOUSE_OVER) = SimpleTime(v_aEvent(g_TIME_START, x)) _
				& " to " & SimpleTime(v_aEvent(g_TIME_END, x))
		else
			aEvent(g_EVENT_MOUSE_OVER) = g_sMORE_DETAILS
		end if
		transferEvent = aEvent
	End Function
		
	'-------------------------------------------------------------------------
	'	Name: 		normalizeColumnsPerDay()
	'	Purpose: 	make all values in array numeric
	' Modifications:
	'	Date:		Name:	Description:
	'	10/14/03	JEA		Creation
	'-------------------------------------------------------------------------
	Private Sub normalizeColumnsPerDay()
		dim x
		for x = 0 to UBound(m_aColumnsPerDay)
			m_aColumnsPerDay(x) = MakeNumber(m_aColumnsPerDay(x))
		next
	End Sub
End Class
%>