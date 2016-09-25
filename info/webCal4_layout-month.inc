<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' Last updated 3/4/2000

' this generates the body of the month calendar

Function makeMonth(arEvents,strFirstDate,strLastDate)

	dim d			' day of month
	dim intCol		' current calendar column
	dim intRow		' current calendar row
	dim strFont		' normal font definition
	dim strNonDay	' layout for cells that aren't days of this month
	dim strHTML		' hold all this stuff
	
	
	strFirstDay		= WeekDay(strFirstDate)
	strLastDay		= Day(strLastDate)
	strLastPrevDay	= Day(strFirstDate - 1)
	
	strHTML = "<table width='100%' border=0 cellspacing=1 cellpadding=1><tr>" & vbCrLf
	
	' print all the day names as headings
	for intCol = 1 to 7
		strHTML = strHTML & "   <td width='14%' align='center' bgcolor='#" & g_arColor(2) & "'>" _
			& VbCrLf & "<font face='" & g_arFont(1) & "' size=1>" _
			& WeekDayName(intCol,0) & "</font></td>" & VbCrLf
	next
	
	strHTML = strHTML & "<td></td><tr>"
	
	' ---------------------------------------------------------
	' now generate calendar body
	' ---------------------------------------------------------
	
	' the column variable keeps constant track of the
	' current calendar column
	
	intCol = 0
	
	strFont = "<font face='" & g_arFont(0) & "' size=2><b>"
	strNonDay = "<td valign='top' bgcolor='#" & g_arColor(11) _
		& "'>" & strFont & "<font color='#" & g_arColor(13) & "'>" & VbCrLf
	
	' cycle through all the days previous to the first
	' day of the active month
	
	for d = 1 to strFirstDay - 1
		strHTML = strHTML & strNonDay & strLastPrevDay - strFirstDay + d + 1 _
			& "</b></font></font></td>" & VbCrLf
		intCol = intCol + 1
	next
	
	' now cycle through all the days of the current month
	
	intRow = 1
	for d = 1 to strLastDay
		intCol = intCol + 1
		strHTML = strHTML & "<td height=45 valign='top' bgcolor='#"
		if y & m & d = Year(now) & Month(now) & Day(now) then
			strHTML = strHTML & g_arColor(8)
		elseif intCol = 1 or intCol = 7 then
			strHTML = strHTML & g_arColor(10)
		else
			strHTML = strHTML & g_arColor(9)
		end if
		strHTML = strHTML & "'>" & strFont & VbCrLf _
			& VbCrLf & "<a href='webCal4_event-edit.asp?" _
			& "date=" & Dateserial(y, m, d) & "&view=month' " _
			& VbCrLf _
			& showStatus("Add a new event to " & DateSerial(y, m, d)) _
			& ">" & d & "</a></b></font>"
			
		' if the day contains events then generate link to day detail
			
		if arEvents(d) <> "" then
			strHTML = strHTML & " <a href='webCal4_day.asp?date=" _
				& DateSerial(y, m, d) & "' " & switchIcon("Day" & d, "Day", "View " & DateSerial(y, m, d) & " detail") _
				& "><img name='Day" & d & "' src='./images/day_grey.gif' border=0></a>"
		end if
	
		strHTML = strHTML & "<br>" & VbCrLf & "<font face='" & g_arFont(2) & "' size=1>" _
			& arEvents(d) & "</font></td>" & VbCrLf
		if intCol = 7 AND d <= strLastDay then
		
		' if we're at the last column then generate link to week view
		
			strHTML = strHTML & "<td valign='center'><a href='webCal4_week.asp?date=" _
				& DateSerial(y, m, d) & "' " & switchIcon("Week" & intRow, "Week", "View week " & intRow) _
				& "><img name='Week" & intRow _
				& "' src='./images/week_grey.gif' border=0></a></td>"
			
			' only start a new row if days of the month remain
	
			if d < strLastDay then
				strHTML = strHTML & "<tr>"
			end if
			strHTML = strHTML & VbCrLf
			
			intCol = 0
			intRow = intRow + 1
		end if
	next
	
	' finally, cycle through as many days of the next month as
	' necessary to fill the calendar grid through column 7
	
	if intCol > 0 then
		d = 1
		do while intCol < 7
			strHTML = strHTML & strNonDay & d & "</font></b></font></td>" & VbCrLf
			d = d + 1
			intCol = intCol + 1
		loop
		strHTML = strHTML & "<td valign='center'><a href='webCal4_week.asp?date=" _
			& strNext & "' " & switchIcon("Week", "Week", "View week " & intRow) _
			& "><img name='Week' src='./images/week_grey.gif' border=0></a></td>"
	end if
	
	strHTML = strHTML & vbCrLf & "</table>"
	
	makeMonth = strHTML

End Function
%>