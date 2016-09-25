<!--#include file="wc_grid_cls.asp"-->
<%
' Copyright 1996-2004 Jason Abbott (webcal@webott.com)
' methods for displaying week view

Class wcWeek

	Private m_bUserCanAdd
	Private m_oLayout
	Private m_oGrid
	
	Private Sub Class_Initialize()
		'm_bUserCanAdd = userCanAdd()
		Set m_oLayout = New wcLayout
		Set m_oGrid = New wcGrid
	End Sub
	
	Private Sub Class_Terminate
		Set m_oGrid = nothing
		Set m_oLayout = nothing
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		writeHTML()
	'	Purpose: 	generates complete week table
	' Modifications:
	'	Date:		Name:	Description:
	'	2/28/01		JEA		Creation
	'	9/25/03		JEA		Update to use .writes
	'-------------------------------------------------------------------------
	Public Sub writeHTML(ByVal v_dtDate)
		Const sVIEW = "week"
		dim sTitle
		dim aDates				' array of date parameters for this view
		dim aColumnsPerDay		' array of column count per day
		dim lDayColumns			' event columns spanned within day column
		dim lDaysInWeek			' five or seven day week, zero-based
		dim lDayWidthPercent	' table percentage width per column
		dim lDayOfWeek			' current day of week
		dim dtPrevDate			' date from previous week
		dim dtNextDate			' date in next week
		dim dtToday				' date of day, looped for week
		dim sClass				' style sheet class
		
		aDates = getDates(v_dtDate)
		lDaysInWeek = aDates(g_LAST_DAY) - aDates(g_FIRST_DAY)	' zero-based
		m_oGrid.columnsPerDay = lDaysInWeek
		lDayWidthPercent = Round(90/(lDaysInWeek + 1), 0)
		dtToday = aDates(g_FIRST_DATE)
		Call m_oGrid.initialize(aDates, g_WEEK)
		aColumnsPerDay = m_oGrid.columnsPerDay
		sTitle = "Week " & DatePart("ww", aDates(g_FIRST_DATE)) & " in " & Year(aDates(g_FIRST_DATE))
		
		with response
			.write "<table width='100%' border='0' cellspacing='0' cellpadding='0'>"
			Call m_oLayout.writeButtons(sTitle, sVIEW, aDates)
			.write "<table width='100%' border='0' cellspacing='0' cellpadding='0' class='gridView'>"
			
			.write "<tr><td rowspan='2' width='5%' class='gridCorner'><img src='./images/tiny_blank.gif'></td>"
			
			' month title and link
			Call writeMonthLinkTD(dtToday, lDaysInWeek)
			.write "<td rowspan='2' width='1%' class='gridCorner'><img src='./images/tiny_blank.gif'></td><tr>"
			
			' weekday titles and day links
			for lDayOfWeek = aDates(g_FIRST_DAY) to aDates(g_LAST_DAY)
				lDayColumns = aColumnsPerDay(lDayOfWeek - 1)
				sClass = "Biz"
			
				if dtToday = Date then
					' highlight current day
					sClass = "Today"
				elseif lDayOfWeek = 1 or lDayOfWeek = 7 then
					' shade weekend
					sClass = "End"
				end if
				
				.write "<td width='"
				.write lDayWidthPercent
				.write "%'"
				if lDayColumns > 0 then
					.write " colspan='"
					.write lDayColumns + 1
					.write "'"
				end if
				.write " class='week week"
				.write sClass
				.write "'>"
				' day heading
				.write "<table cellspacing='0' cellpadding='0' border='0' width='100%'><tr>"
				.write "<td class='WeekDate'>"
				.write Day(dtToday)								' day date
				.write "</td><td class='WeekdayName'>"
				.write "<a href='"								' link to events
				.write g_sFILE_PREFIX
				.write "day.asp?date="
				'.write "event-edit.asp?action=new&view=week&date="
				.write dtToday
				.write "' "
				Call m_oLayout.writeStatusJS("View " & dtToday)
				.write ">"
				.write WeekDayName(lDayOfWeek, 0)
				.write "</a></td></table></td>"
				dtToday = DateAdd("d", 1, dtToday)
			next
			
			Call m_oGrid.writeEvents(lDaysInWeek, lDayWidthPercent)
			
			.write "</table></table>"
		end with
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		writeMonthLinkTD()
	'	Purpose: 	create month title and link
	' Modifications:
	'	Date:		Name:	Description:
	'	2/28/01		JEA		Creation
	'	9/29/03		JEA		Converted to sub with .writes
	'-------------------------------------------------------------------------
	Private Sub writeMonthLinkTD(ByVal v_dtDate, ByVal v_lDaysInWeek)
		dim lSpan
		dim dtDate
		dtDate = v_dtDate
		with response
			for each lSpan in getMonthColSpans(v_dtDate, v_lDaysInWeek)
				.write "<td colspan='"
				.write lSpan
				.write "' class='monthLink'><a href='"
				.write g_sFILE_PREFIX
				.write "month.asp?date="
				.write v_dtDate
				.write "' "
				Call m_oLayout.writeStatusJS("View all of " & MonthName(Month(v_dtDate)))
				.write ">"
				.write MonthName(Month(dtDate))
				.write " "
				.write Year(dtDate)
				.write "</a></td>"
				dtDate = DateAdd("m", 1, dtDate)
			next
		end with
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		getMonthColSpans()
	'	Purpose: 	calculate days spanned by month titles for weeks crossing months
	'	Return: 	array
	' Modifications:
	'	Date:		Name:	Description:
	'	9/29/03		JEA		Creation
	'-------------------------------------------------------------------------
	Private Function getMonthColSpans(ByVal v_dtDate, ByVal v_lDaysInWeek)
		dim aSpan()				' array of days spanned by month titles
		dim aColumnsPerDay
		dim lMonth				' month in loop
		dim x					' loop counter
		
		aColumnsPerDay = m_oGrid.columnsPerDay
		ReDim aSpan(1)
		aSpan(0) = 0 : aSpan(1) = 0
		lMonth = Month(v_dtDate)
		for x = 0 to v_lDaysInWeek
			if Month(v_dtDate) <> lMonth then
				aSpan(1) = (v_lDaysInWeek - aSpan(0)) + 1
				exit for
			else
				aSpan(0) = aSpan(0) + aColumnsPerDay(x) + 1
				v_dtDate = DateAdd("d", 1, v_dtDate)
			end if
		next
		if aSpan(1) = 0 then ReDim Preserve aSpan(0)
		getMonthColSpans = aSpan
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		getDates()
	'	Purpose: 	generate dates that define this week
	'	Return: 	array
	' Modifications:
	'	Date:		Name:	Description:
	'	2/28/01		JEA		Creation
	'-------------------------------------------------------------------------
	Private Function getDates(ByVal v_dtDate)
		dim aDates(6)
		
		aDates(g_THIS_DATE) = IIf(IsDate(v_dtDate), v_dtDate, Date)
		
		' define first and last days to display
		if Session(g_sDB_NAME & "Weekends") or true then
			aDates(g_FIRST_DAY) = 1 : aDates(g_LAST_DAY) = 7
		else
			aDates(g_FIRST_DAY) = 2 : aDates(g_LAST_DAY) = 6
		end if
		
		aDates(g_FIRST_DATE) = DateAdd("d", aDates(g_FIRST_DAY) - WeekDay(aDates(g_THIS_DATE)), aDates(g_THIS_DATE))
		aDates(g_LAST_DATE) = DateAdd("d", aDates(g_LAST_DAY) - 1, aDates(g_FIRST_DATE))
		aDates(g_PREV_DATE) = DateAdd("d", -aDates(g_LAST_DAY), aDates(g_FIRST_DATE))
		aDates(g_NEXT_DATE) = DateAdd("d", aDates(g_FIRST_DAY), aDates(g_FIRST_DATE))
		getDates = aDates
	End Function
End Class
%>
