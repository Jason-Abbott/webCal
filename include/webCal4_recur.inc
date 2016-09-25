<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' Last updated 02/29/2001

' creates specialized views of the different recurrence types (updated 2/24/01)
'   parameters are an array of event dates, current
'   date, event id and the type of recurrence
' returns string ---------------------------------------------------------
Function showRecur(ByVal v_arDates, ByVal v_strDate, ByVal v_intEventID, _
	ByVal v_strRecur, ByVal v_strView)
	
	dim oYears			' years on which the event occurs
	dim strYear			' temporarily hold year of event
	dim strHTML			' HTML returned by function
	dim strDateList		' list of dates in year definition
	dim x				' loop counter

	Set oYears = CreateObject("Scripting.Dictionary")

	' group event occurences by year
	for x = 0 to UBound(v_arDates)
		'response.write Year(v_arDates(x)) : response.flush
		strYear = Year(v_arDates(x))
		if oYears.Exists(strYear) then
			' event already occurs in this year so add date
			' to list, then purge item and re-add all
			strDateList = oYears.Item(strYear)
			oYears.Remove(strYear)
			' insert delimiter for subsequent splitting by month
			oYears.Add strYear, strDateList & " " & v_arDates(x)
		else
			oYears.Add strYear, v_arDates(x)
		end if
	next

	' generate the text that describes the type of recurrence
	' for each event and a view of recurrence dates
	select case v_strRecur
		case "yearly"
			strHTML = RecurByYear(v_arDates, v_strDate, oYears, v_strView)
		case "monthly"
			strHTML = RecurByMonth(v_arDates, v_strDate, oYears, v_strView)
		case "2weeks", "weekly"
			strHTML = RecurByWeek(v_arDates, v_strDate, oYears, v_strView, v_intEventID)
		case "daily"
			strHTML = RecurByDay(v_arDates, v_strDate, oYears, v_strView, v_intEventID)
		case "none"
			strHTML = NoRecur(v_arDates, v_intEventID, v_strView)
	end select
	oYears.RemoveAll	
	showrecur = strHTML
End Function

' build HTML to show recurrence by year (updated 2/23/01)
' returns string ---------------------------------------------------------
Private Function RecurByYear(ByVal v_arDates, ByVal v_strDate, _
	ByVal v_oYears, ByVal v_strView)
	
	dim strYear
	dim strHTML
	
	strHTML = "<font face='" & g_arFont(1) & "' size=2>" _
		& "scheduled yearly on " _
		& MonthName(Month(v_arDates(0))) _
		& " " & Day(v_arDates(0)) _
		& " during these oYears:" _
		& "<select onChange=""location.href='" _
			& "webCal4_detail.asp?event_id=" _
		& v_intEventID & "&view=" & v_strView _
		& "&date=" & Month(v_arDates(0)) & "/" _
		& Day(v_arDates(0)) & "/' + " _
		& "options[selectedIndex].value;"">" & VbCrLf
		
	for each strYear in v_oYears.Keys
		strHTML = strHTML & "<option value=" & strYear
		if strYear = Year(v_strDate) then
			strHTML = strHTML & " selected"
		end if
		strHTML = strHTML & ">" & strYear & VbCrLf
	next
	strHTML = strHTML & "</select></font>" & VbCrLf
	
	RecurByYear = strHTML
End Function

' build HTML to show recurrence by month (updated 2/23/01)
' returns string ---------------------------------------------------------
Private Function RecurByMonth(ByVal v_arDates, ByVal v_strDate, _
	ByVal v_oYears, ByVal v_strView)
	
	dim strHTML
	dim arMonths(12)	' array of months in monthly recurrence
	dim x
	
	strHTML = "<font face='" & g_arFont(1) & "' size=1>" _
		& "scheduled for the " & Day(v_arDates(0))

	' generate the correct wording for the day number
	select case Right(Day(v_arDates(0)), 1)
		case 1
			strHTML = strHTML & "st"
		case 2
			strHTML = strHTML & "nd"
		case 3
			strHTML = strHTML & "rd"
		case else
			strHTML = strHTML & "th"
	end select
	strHTML = strHTML & " of these months:<br>"

	' go through each year
	for each strYear in oYears.Keys
		' for each year create an array of scheduled months
		for each x in split(v_oYears.Item(strYear))
			arMonths(Month(x)) = MonthName(Month(x))
		next
		
		' then display each item in the month array
		strHTML = strHTML & "<table><tr><td>" & VbCrLf _
			& "<select onChange=""location.href=" _
			& "'webCal4_detail.asp?event_id=" _
			& v_intEventID & "&view=" & v_strView _
			& "&date=' + options[selectedIndex].value + " _
			& "'/" & Day(v_arDates(0)) & "/" & strYear & "';"">" & VbCrLf
			
		for each x in arMonths
			if x <> "" then
				strHTML = strHTML & "<option value=" & x
				if x = MonthName(Month(v_strDate)) then
					strHTML = strHTML & " selected"
				end if
				strHTML = strHTML & ">" & x & VbCrLf
			end if
		next
		Erase arMonths
		strHTML = strHTML & "</select></td>" & VbCrLf _
			& "<td><font face='" & g_arFont(1) & "' size=2>" _
			& strYear & "</font></td>" & VbCrLf & "</table>"
	next
	
	RecurByMonth = strHTML
End Function

' built HTML to show recurrence by week (updated 2/23/01)
' returns string ---------------------------------------------------------
Private Function RecurByWeek(ByVal v_arDates, ByVal v_strDate, _
	ByVal v_oYears, ByVal v_strView, ByVal v_intEventID)

	dim strHTML
	strHTML = "<font face='" & g_arFont(1) & "' size=1>" _
		& "scheduled for every other " _
		& WeekDayName(WeekDay(v_arDates(0))) _
		& " on these dates:<br>" _
		& RecurDayGrid(v_arDates, v_strDate, v_oYears, v_strView, v_intEventID)
	RecurByWeek = strHTML
End Function

' build HTML to show recurrence by day (updated 2/23/01)
' returns string ---------------------------------------------------------
Private Function RecurByDay(ByVal v_arDates, ByVal v_strDate, _
	ByVal v_oYears, ByVal v_strView, ByVal v_intEventID)
	
	dim strHTML
	strHTML = "<font face='" & g_arFont(1) & "' size=1>" _
		& "scheduled daily for these dates:<br>" _
		& RecurDayGrid(v_arDates, v_strDate, v_oYears, v_strView, v_intEventID)
	RecurByDay = strHTML
End Function

' build HTML to show no recurrence (updated 2/23/01)
' returns string ---------------------------------------------------------
Private Function NoRecur(ByVal v_arDates, ByVal v_intEventID, ByVal v_strView)
	dim strHTML
	dim strLink
	
	strLink = MonthName(Month(v_arDates(0))) & " " _
		& Day(v_arDates(0)) & ", " & Year(v_arDates(0))
		
	strHTML = "<font face='" & g_arFont(1) & "' size=2>" _
		& makeLink(v_intEventID, v_strView, v_arDates(0), "View details", strLink)
	NoRecur = strHTML
End Function

' build HTML to show day distribution by months and years (updated 2/23/01)
' returns string ---------------------------------------------------------
Private Function RecurDayGrid(ByVal v_arDates, ByVal v_strDate, _
	ByVal v_oYears, ByVal v_strView, ByVal v_intEventID)
	 
	dim strYear
	dim strHTML
	dim arMonths(12)	' array of months in monthly recurrence
	dim strCombine		' boolean combination text
	dim strLink			' text linked to event detail
	dim strStatus		' text displayed in status bar
	dim strDate			' local event date
	dim x

	for each strYear in v_oYears.Keys
		strHTML = strHTML & "<table width='100%'><tr>" & VbCrLf _
			& "<td valign='top' align='center'>" _
			& "<font face='" & g_arFont(1) & "' size=2>" _
			& strYear & "</font></td>" & VbCrLf
	
		' go through each event date in that year and create an array
		' of days for each month
		for each x in split(v_oYears.Item(strYear))
			if arMonths(Month(x)) <> "" then
				strCombine = "<br>"
			else
				strCombine = ""
			end if
			strStatus = "View "	& MonthName(Month(x)) & " " & Day(x) & ", " _
				& strYear & " detail"
			strDate = Month(x) & "/" & Day(x) & "/" & strYear

			arMonths(Month(x)) = arMonths(Month(x)) & strCombine _
				& makeLink(v_intEventID, v_strView, strDate, strStatus, Day(x))
		next
	
		' create month headings and make them link to the calendar
		' view for that month and year
		strHTML = strHTML & "<td><table width='100%' " _
			& "cellpadding=2 cellspacing=0 border=0>" _
			& "<tr>" & VbCrLf
				
		for x = 1 to 12
			strHTML = strHTML & "   <td align='center' " _
				& "width='8.3%' bgcolor='#" & g_arColor(0) & "'>" _
				& "<font face='" & g_arFont(0) & "' size=1"
			if arMonths(x) <> "" then
				strHTML = strHTML & "><a href='webCal4_month.asp?date=" _
					& x & "/1/" & strYear & "' " _
					& showStatus("View " _
						& MonthName(x) & " " & strYear _
						& " calendar") _
					& ">" & MonthName(x,1) & "</a>"
			else
				strHTML = strHTML & " color=""#" & g_arColor(5) & """>" & MonthName(x,1)
			end if
			strHTML = strHTML & "</font></td>" & VbCrLf
		next
		
		strHTML = strHTML & "<tr>" & VbCrLf
		for x = 1 to 12
			strHTML = strHTML & "   <td valign='top' align='right'>" _
				& "<font face='" & g_arFont(1) & "' size=1>" _
				& arMonths(x) & "</font></td>" & VbCrLf
		next
		strHTML = strHTML & "</table></td>" & VbCrLf & "</table>"
		Erase arMonths
		Erase v_arDates
	next

	RecurDayGrid = strHTML
End Function

' build common link to event detail (updated 2/23/01)
' returns string ---------------------------------------------------------
Private Function makeLink(ByVal v_intEventID, ByVal v_strView, _
	ByVal v_strDate, ByVal v_strStatus, ByVal v_strLink)
	
	dim strHTML
	strHTML = "<a href='webCal4_detail.asp?event_id=" _
		& v_intEventID & "&view=" & v_strView _
		& "&date=" & v_strDate & "' " & showStatus(v_strStatus) _
		& ">" & v_strLink & "</a>"
	makeLink = strHTML
End Function
%>