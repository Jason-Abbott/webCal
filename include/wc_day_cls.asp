<!--#include file="wc_grid_cls.asp"-->
<!--#include file="wc_month_cls.asp"-->
<%
' Copyright 1996-2004 Jason Abbott (webcal@webott.com)
' methods for displaying week view

Class wcDay

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
		Const sVIEW = "day"
		dim sTitle
		dim aDates				' array of date parameters for this view
		dim aColumnsPerDay		' array of column count per day
		dim oMonth
		
		aDates = getDates(v_dtDate)
		m_oGrid.columnsPerDay = 0
		Call m_oGrid.initialize(aDates, g_DAY)
		aColumnsPerDay = m_oGrid.columnsPerDay
		sTitle = "Day " & DatePart("y", aDates(g_THIS_DATE)) & " in " & Year(aDates(g_THIS_DATE))
		
		with response
			.write "<table width='100%' border='0' cellspacing='0' cellpadding='0'>"
			.write "<tr><td>"
			Call m_oLayout.writeButtons(sTitle, sVIEW, aDates)
			.write "</td><td></td><tr><td width='80%'>"
			.write "<table width='100%' border='0' cellspacing='0' cellpadding='0' class='gridView'>"
			.write "<td width='5%'></td><td width='92%'"
			if aColumnsPerDay(0) > 0 then
				.write " colspan='"
				.write aColumnsPerDay(0) + 1
				.write "'"
			end if
			.write "></td><td colspan='2' width='3%'></td>"
			Call m_oGrid.writeEvents(0, 100)
			.write "</table></td><td valign='top' align='center'>"
			.write WeekdayName(WeekDay(aDates(g_THIS_DATE)))
			.write "<br>"
			.write Day(aDates(g_THIS_DATE))
			.write "<br>"
			.write "<a href='"
			.write g_sFILE_PREFIX
			.write "month.asp?date="
			.write aDates(g_THIS_DATE)
			.write "'>"
			.write MonthName(Month(aDates(g_THIS_DATE)),1)
			.write "</a><br>"
			.write Year(aDates(g_THIS_DATE))

			Set oMonth = New wcMonth
			Call oMonth.writeNavMonthHTML(v_dtDate)
			Set oMonth = Nothing
			
			.write "</td></table>"
			
		end with
	End Sub
	
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
		aDates(g_FIRST_DATE) = aDates(g_THIS_DATE)
		aDates(g_NEXT_DATE) = DateAdd("d", 1, aDates(g_FIRST_DATE))
		aDates(g_PREV_DATE) = DateAdd("d", -1, aDates(g_FIRST_DATE))
		aDates(g_LAST_DATE) = aDates(g_THIS_DATE)
		aDates(g_FIRST_DAY) = Day(aDates(g_FIRST_DATE))
		aDates(g_LAST_DAY) = Day(aDates(g_LAST_DATE))

		getDates = aDates
	End Function
End Class
%>
