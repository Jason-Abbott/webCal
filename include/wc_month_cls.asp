<%
' Copyright 1996-2004 Jason Abbott (webcal@webott.com)
' methods for displaying month view

Class wcMonth

	Private m_bUserCanAdd
	Private m_oLayout
	
	Private Sub Class_Initialize()
		 m_bUserCanAdd = userCanAdd()
		 Set m_oLayout = New wcLayout
	End Sub
	
	Private Sub Class_Terminate()
		Set m_oLayout = nothing
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		writeHTML()
	'	Purpose: 	generates the body of the month calendar
	' Modifications:
	'	Date:		Name:	Description:
	'	3/14/01		JEA		Creation
	'	9/23/03		JEA		Update to use .writes
	'-------------------------------------------------------------------------
	Public Function writeHTML(ByVal v_lYear, ByVal v_lMonth, ByVal v_dtDate)
		Const sVIEW = "month"
		dim y, m, d			' year, month, day, event
		dim lColumn			' current calendar column
		dim lRow			' current calendar row
		dim lLastPrevDay	' last day of previous month
		dim sHTML			' hold all this stuff
		dim aDates
		dim aEvents
		dim sTitle
		dim x
		
		aDates = getDates(v_lYear, v_lMonth, v_dtDate)
		lLastPrevDay = Day(aDates(g_FIRST_DATE) - 1)
		aEvents = getEvents(aDates(g_FIRST_DATE), aDates(g_LAST_DATE))
		y = Year(aDates(g_FIRST_DATE))
		m = Month(aDates(g_FIRST_DATE))
		sTitle = MonthName(Month(aDates(g_FIRST_DATE))) & " " & Year(aDates(g_FIRST_DATE))
		
		with response
			.write "<table width='100%' border='0' cellspacing='0' cellpadding='1'><tr><td>"
			Call m_oLayout.writeButtons(sTitle, sVIEW, aDates)
			.write "</td><tr><td>"
			.write "<table width='100%' border='0' cellspacing='1' cellpadding='0' class='monthView'><tr>"

			' write weekday names as header
			for lColumn = 1 to 7
				.write "<td width='14%' class='dayName'>"
				.write WeekDayName(lColumn ,0)
				.write "</td>"
			next
	 		.write "<td></td><tr>"
			
			' write leftover days from last month
			lColumn = 0
			for d = 1 to aDates(g_FIRST_DAY) - 1
				.write "<td valign='top' class='other'>"
				.write lLastPrevDay - aDates(g_FIRST_DAY) + d + 1
				.write "</td>"
				lColumn = lColumn + 1
			next
			
			' go through days of month
			lRow = 1
			for d = 1 to aDates(g_LAST_DAY)
				lColumn = lColumn + 1

				' container cell
				.write "<td height='45' valign='top' class='day"
				if y & m & d = Year(now) & Month(now) & Day(now) then
					.write "This"
				elseif lColumn = 1 Or lColumn = 7 then
					.write "Weekend"
				else
					.write "Common"
				end if
				.write "'>"
				
				' table within cell to allow right-aligned day link
				.write "<table width='100%' cellspacing='0' cellpadding='0' border='0'>"
				.write "<tr><td class='dayAdd'>"
				if m_bUserCanAdd then
					' make day number a link
					.write "<a href='"
					.write g_sFILE_PREFIX
					.write "event-edit.asp?date="
					.write Dateserial(y, m, d)
					.write "&view=month' "
					Call m_oLayout.writeStatusJS("Add a new event to " & DateSerial(y, m, d))
					.write ">"
					.write d
					.write "</a>"
				else
					.write d
				end if
				.write "</td>"
				
				if IsArray(aEvents(d)) then
					' write events if any
					.write "<td class='dayLink'><a href='"
					.write g_sFILE_PREFIX
					.write "day.asp?date="
					.write DateSerial(y, m, d)
					.write "'>Show Day</a></td></table><div class='eventTitle'>"
					Call writeDayEvents(aEvents(d))
					.write "</div></td>"
				else
					.write "</table>"	
				end if
					
				' make link to week view on last column
				if lColumn = 7 AND d <= aDates(g_LAST_DAY) then
					.write "<td valign='center'><a href='"
					.write g_sFILE_PREFIX
					.write "week.asp?date="
					.write DateSerial(y, m, d)
					.write "' "
					Call m_oLayout.writeRolloverJS("Week" & lRow, "Week", "View week " & lRow)
					.write "><img name='Week"
					.write lRow
					.write "' src='./images/week_grey.gif' border=0></a></td>"
					
					' only start a new row if days of the month remain
					if d < aDates(g_LAST_DAY) then .write "<tr>"
					lColumn = 0 : lRow = lRow + 1
				end if
			next
			
			' first days of next month
			if lColumn > 0 then
				d = 1
				do while lColumn < 7
					.write "<td valign='top' class='other'>"
					.write d
					.write "</td>"
					d = d + 1 : lColumn = lColumn + 1
				loop
				.write "<td valign='center'><a href='"
				.write g_sFILE_PREFIX
				.write "week.asp?date="
				.write aDates(g_NEXT_DATE)
				.write "' "
				Call m_oLayout.writeRolloverJS("Week", "Week", "View week " & lRow)
				.write "><img name='Week' src='./images/week_grey.gif' border=0></a></td>"
			end if
			.write "</table></td>"
			'.write "<tr><td valign='top'><div class='footnote'>"
			'.write showLoadTime(m_strQuery, m_strLoadFrom)
			'.write "<a href='"
			'.write g_sHOME_PAGE
			'.write "' target='_top'>webCal 4.0</a></div></td>"
			'.write "<td align='right'><form>"
			'.write makeButton(g_sBTN_LOGOUT,"logout();",12,60)
			'.write "&nbsp;"
			'.write makeButton(g_sBTN_DISPLAY,"showSettings();",12,160)
			.write "</form></td></table>"
		end with
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		writeDayEvents()
	'	Purpose: 	write events for the day
	' Modifications:
	'	Date:		Name:	Description:
	'	9/25/03		JEA		Creation
	'-------------------------------------------------------------------------
	Private Sub writeDayEvents(ByVal v_aEvents)
		dim x
		with response
			for x = 0 to UBound(v_aEvents)
				.write "<img src='./images/arrow_right_"
				.write v_aEvents(x)(g_EVENT_COLOR)
				.write ".gif' width='4' height='7'> "
				.write "<a href='"
				.write g_sFILE_PREFIX
				.write "detail.asp?event_id="
				.write v_aEvents(x)(g_EVENT_ID)
				.write "&date="
				.write v_aEvents(x)(g_EVENT_DATE)
				.write "&view=month' "
				Call m_oLayout.writeStatusJS(v_aEvents(x)(g_EVENT_MOUSE_OVER))
				.write ">"
				.write v_aEvents(x)(g_EVENT_TITLE)
				.write "</a><br>"
			next
		end with
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		writeNavMonthHTML()
	'	Purpose: 	write small month to navigate to days
	' Modifications:
	'	Date:		Name:	Description:
	'	3/3/01		JEA		Creation
	'	10/16/03	JEA		Use .writes
	'-------------------------------------------------------------------------
	Public Sub writeNavMonthHTML(ByVal v_dtDate)
		dim y, m, d			' year, month, day, event
		dim aHasEvents(31)	' boolean array indicating events per day
		dim lColumn			' current calendar column
		dim lRow			' current calendar row
		dim lLastPrevDay	' last day of previous month
		dim sHTML			' hold all this stuff
		dim aDates
		dim aEvents
		dim sTitle
		dim x
	
		aDates = getDates("", "", v_dtDate)
		lLastPrevDay = Day(aDates(g_FIRST_DATE) - 1)
		aEvents = getEvents(aDates(g_FIRST_DATE), aDates(g_LAST_DATE))
		y = Year(aDates(g_FIRST_DATE))
		m = Month(aDates(g_FIRST_DATE))
		
		for x = 0 to UBound(aHasEvents)
			aHasEvents(x) = IsArray(aEvents(x))
		next
		Erase aEvents
	  
		with response
			.write "<table border='0' cellspacing='1' cellpadding='1' class='navMonth'><tr>"

			' write weekday names
			for x = 1 to 7
				.write "<td width='14.3%' class='navWeekdayName'>"
				.write Left(WeekdayName(x), 1)
				.write "</td>"
			next
			.write "<tr>"
			
			' write leftover days from last month
			lColumn = 0
			for d = 1 to aDates(g_FIRST_DAY) - 1
				.write "<td class='navDay other'>"
				.write lLastPrevDay - aDates(g_FIRST_DAY) + d + 1
				.write "</td>"
				lColumn = lColumn + 1
			next

			' go through days of month
			lRow = 1

			for d = 1 to aDates(g_LAST_DAY)
				lColumn = lColumn + 1
				.write "<td height='45' valign='top' class='navDay "
				if y & m & d = Year(now) & Month(now) & Day(now) then
					.write "weekToday"
				elseif lColumn = 1 Or lColumn = 7 then
					.write "weekEnd"
				else
					.write "weekBiz"
				end if
				.write "'>"
				
				if aHasEvents(d) then
					.write "<a href='"
					.write g_sFILE_PREFIX
					.write "day.asp?date="
					.write DateSerial(y,m,d)
					.write "'>"
				end if
				
				.write d
				.write "</a></td>"
				
				if lColumn = 7 AND d <= aDates(g_LAST_DAY) then
					.write "<tr>"
					lColumn = 0
				end if
			next
		
			' first days of next month
			if lColumn > 0 then
				d = 1
				do while lColumn < 7
					.write "<td class='navDay other'>"
					.write d
					.write "</td>"
					d = d + 1 : lColumn = lColumn + 1
				loop
			end if
			.write "</table>"
		end with
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		getDates()
	'	Purpose: 	generate dates that define this month
	'	Return: 	array
	' Modifications:
	'	Date:		Name:	Description:
	'	3/3/01		JEA		Creation
	'-------------------------------------------------------------------------
	Private Function getDates(ByVal v_lYear, ByVal v_lMonth, ByVal v_dtDate)
		dim aDates(6)
			
		' retrieve selected date
		if IsNumber(v_lMonth) then
			aDates(g_THIS_DATE) = DateSerial(v_lYear, v_lMonth, 1)
		elseif IsVoid(v_dtDate) then
			aDates(g_THIS_DATE) = Date
		else
			aDates(g_THIS_DATE) = v_dtDate
		end if
	
		aDates(g_FIRST_DATE) = Dateserial(Year(aDates(g_THIS_DATE)), Month(aDates(g_THIS_DATE)), 1)
		aDates(g_NEXT_DATE) = DateAdd("m", 1, aDates(g_FIRST_DATE))
		aDates(g_PREV_DATE) = DateAdd("m", -1, aDates(g_FIRST_DATE))
		aDates(g_LAST_DATE) = DateAdd("d", -1, aDates(g_NEXT_DATE))
		aDates(g_FIRST_DAY) = WeekDay(aDates(g_FIRST_DATE))
		aDates(g_LAST_DAY) = Day(aDates(g_LAST_DATE))
		
		getDates = aDates
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		getEvents()
	'	Purpose: 	put all matching events in an array indexed by day number
	'	Return: 	array
	' Modifications:
	'	Date:		Name:	Description:
	'	3/2/01		JEA		Creation
	'-------------------------------------------------------------------------
	Private Function getEvents(ByVal v_dtFirst, ByVal v_dtLast)
		dim aDay(31)			' array of events as HTML for each day of month
		dim aToday()			' array of events on single day
		dim aEvents				' array of events from query
		dim aTempEvent()		' template event array
		dim lIndex				' day of month index into array
		dim lDay				' day of month
		dim lUBound				' bound of per day array
		dim sDescription		' event description
		dim oData
		dim oSession
		dim x, y				' loop counters
		
		'if v_sQuery <> g_sNO_EVENTS then
			Set oData = New wcData : Set oSession = New wcSession
			aEvents = oData.GetArray(oSession.getViewQuery(v_dtFirst, v_dtLast))
			Set oSession = Nothing : Set oData = Nothing
		'end if
		
		if IsArray(aEvents) then
			lDay = 0
			lUBound = 0
			ReDim aTempEvent(g_EVENT_MOUSE_OVER)
			for x = 0 to UBound(aEvents, 2)
				if lDay <> Day(aEvents(g_EVENT_DATE, x)) then
					' create new day array
					if lDay > 0 then aDay(lDay) = aToday
					lDay = Day(aEvents(g_EVENT_DATE, x))
					lUBound = 0
				else
					' increment existing event array within day
					lUBound = lUBound + 1
				end if
				ReDim Preserve aToday(lUBound)
				aToday(lUBound) = aTempEvent
				for y = g_EVENT_ID to g_EVENT_SKIP_WE
					' transfer event details into month array
					aToday(lUBound)(y) = aEvents(y, x)
				next
				
				' generate mouse-over description
				if aEvents(g_TIME_START, x) <> "" then
					aToday(lUBound)(g_EVENT_MOUSE_OVER) = SimpleTime(aEvents(g_TIME_START, x)) _
						& " to " & SimpleTime(aEvents(g_TIME_END, x))
				else
					aToday(lUBound)(g_EVENT_MOUSE_OVER) = g_sMORE_DETAILS
				end if
			next
			aDay(lDay) = aToday
		end if
		getEvents = aDay
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		userCanAdd()
	'	Purpose: 	determine whether user has sufficient permissions to add events
	'	Return: 	boolean
	' Modifications:
	'	Date:		Name:	Description:
	'	3/3/01		JEA		Creation
	'-------------------------------------------------------------------------
	Private Function userCanAdd()
		dim bCanAdd
		dim aGroups
		dim x
		
		bCanAdd = false
		aGroups = Session(g_sDB_NAME & "Groups")
		for x = 0 to UBound(aGroups, 2)
			if aGroups(g_GROUP_ACCESS, x) > g_READ_ACCESS then
				bCanAdd = true : exit for
			end if
		next
		userCanAdd = bCanAdd
	End Function
End Class
%>