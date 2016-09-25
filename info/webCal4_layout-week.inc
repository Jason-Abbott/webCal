<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' Last updated 3/3/2000

Function makeWeek(arEvents,arDays)

	dim m, y			' month and year
	dim intCol			' current column
	dim intHour			' hour of the day (used for alternating colors)
	dim strEventClr		' event color
	dim strHourClr		' alternating hour colors
	dim strBsClr		' shade business hours
	dim strShow			' weekday column headings
	dim strDays			' header row of days
	dim arSpan()		' cells spanned
	dim intExtra		' not sure
	dim strQueryString	' not sure
	dim strHTML			' build HTML to return
	dim x				' loop counter
	
	strHTML = "<table width='100%' border='0' cellspacing='1' cellpadding='0'>" _
		& vbCrLf & "<tr>" & vbCrLf & "<td rowspan=3 width='10%'>&nbsp;</td>"
	
	' begin by displaying the first date
	strShow = strFirst
	ReDim arSpan(1)
	m = 0
	' go through each column-day to create day headings-------------------
	for intCol = intFirst to intLast
		if strShow = Date then
			' highlight current day
			strEventClr = g_arColor(8)
		elseif intCol = 1 or intCol = 7 then
			' shade weekend
			strEventClr = g_arColor(10)
		else
			strEventClr = g_arColor(9)
		end if
		
		strDays = strDays & "<td width='" & intRatio & "%' bgcolor='#" _
			& strEventClr & "'>" & VbCrLf _
			& "<table cellspacing='0' cellpadding='1' border='0' width='100%'><tr>" & VbCrLf _
			& "<td align='center' bgcolor='#000000' width='20%' valign='top'>" & VbCrLf _
			& "<font face='Verdana, Arial, Helvetica' color='#ffffff' size='2'><b>" _
			& Day(strShow) & "</b></font></td>" _
			& VbCrLf _
			& "<td align='center' rowspan='2'>" _
			& "<font face='Verdana, Arial, Helvetica' size='1'>" _
			& VbCrLf _
			& "<a href='webCal4_edit.asp?action=new&view=week&date=" & strShow & "' " _
			& showStatus("Add a new event to " & strShow) & ">" _
			& WeekDayName(intCol,0) & "</a></td>" _
			& VbCrLf _
			& "<tr>"
	
		' if this day has events, display the link to day detail
		if arDays(intCol-intFirst) then
			strDays = strDays & "<td>" _
			& "<a href='webCal4_day.asp?date=" & strShow & "' " _
			& switchIcon("Day" & intCol, "Day", "View " & strShow & " detail") _
			& "><img name='Day" & intCol & "' src='./images/day_grey.gif' width='18' " _
			& "height='7' border='0'>"
		else
			strDays = strDays & "<td bgcolor='#000000'>" _
				& "<img src='./images/tiny_blank.gif' height='7'>"
		end if
		
		strDays = strDays & "</td>" & VbCrLf & "</table></td>" & VbCrLf
			
		arSpan(m) = arSpan(m) + 1
		' get the next date to show
		strNext = DateAdd("d", 1, strShow)
		if Month(strNext) <> Month(strShow) then m = 1
		strShow = strNext
	next
	
	' display link to month-----------------------------------------------
	strHTML = strHTML & "<tr><td bgcolor='#" & g_arColor(2) & "' colspan=" _
		& arSpan(0) & " align='center'><font face='" & g_arFont(0) _
		& "' size=2 color='#" & g_arColor(6) & "'>" & vbCrLf _
		& "<a href='webCal4_month.asp?date=" & strFirst & "' " _
		& showStatus("View all of " & MonthName(Month(strFirst))) _
		& "><b>" & MonthName(Month(strFirst)) & " " & Year(strFirst) _
		& "</b></a></font></td>"
	
	' if week spans two months, display second month----------------------
	if arSpan(0) < intDays + 1 then
		strHTML = strHTML & "<td bgcolor='#" & g_arColor(2) & "' colspan=" _
			& arSpan(1) & " align='center'><font face='" & g_arFont(0) _
			& "' size=2 color='#" & g_arColor(6) & "'>" & vbCrLf _
			& "<a href='webCal4_month.asp?date=" & strLast & "' " _
			& showStatus("View all of " & MonthName(Month(strLast))) _
			& "><b>" & MonthName(Month(strLast)) & " " & Year(strLast) _
			& "</b></a></font></td>"
	end if
	
	strHTML = strHTML & "<td rowspan=2><img src='./images/tiny_blank.gif' height=1>" _
		& "</td><tr>" & strDays
		
	ReDim arSpan(intDays)
	
	' initialize default spanning for all days to one
	for x = 0 to intDays
		arSpan(x) = 1
	next
	
	' go through each time segment of the day
	' (index to 2D of event array)
	
	for intTime = intSegStart to intSegEnd
	
		strHTML = strHTML & "<tr>" & VbCrLf
		
		if (intTime Mod intFactor = 0 OR intTime = 0) AND intTime <> intSegEnd then
			' this segment begins a new hour--display hour information
			
			strHTML = strHTML & "<td rowspan='" & intFactor & "' align='right' bgcolor='#" _
				& g_arColor(1) & "' height='" & intHeight * intFactor & "'>" _
				& "<font face='Tahoma, Arial, Helvetica' size=2><nobr>"
				
			if intTime > 0 then
				intHour = intTime/intFactor
			else
				intHour = 0
			end if
	
			if intHour = 0 then
				strHTML = strHTML & "<b>midnight</b>"
			elseif intHour < 12 then
				strHTML = strHTML & intHour & ":00 AM"
			elseif intHour = 12 then
				strHTML = strHTML & "<b>noon</b>"
			else
				strHTML = strHTML & intHour - 12 & ":00 PM"
			end if
		
			' alternate hour colors
			if intHour Mod 2 = 0 then
				strHourClr = "ffffff"
			else
				strHourClr = "dfdfdf"
			end if
	
			strHTML = strHTML & "</nobr></font></td>" & VbCrLf
			
		end if
		
		' for every day we're showing
		' (index to 1D of event array)
		for x = 0 to intDays
			if arSpan(x) = 1 then
				' an event is not being shown
				strHTML = strHTML & "<td bgcolor='#"
				
				' 0=event text; 1=segments spanned; 2=color
				
				if arEvents(x,intTime,0) <> "" then
				
				
					' an event exists at this segment (table cell)
					' get correct row span value
					arSpan(x) = arEvents(x,intTime,1)
					
					strHTML = strHTML & arEvents(x,intTime,2) _
						& "' rowspan='" & arSpan(x) & "' align='center'>" _
						& "<font face='Tahoma, Arial, Helvetica' size='1'>" _
						& arEvents(x,intTime,0) & "</font>"
					
					intExtra = 0
				else
					strHTML = strHTML & strHourClr & "' height='" & intHeight & "'>" _
						& "<img src='./images/tiny_blank.gif'>"
				end if
				strHTML = strHTML & "</td>" & VbCrLf
			else
				' decrement the spanning of the current event
				' as we cycle to a new row
				arSpan(x) = arSpan(x) - 1
			end if
		next
		
		' mark business hours
		if intHour < 8 OR intHour > 16 OR (intHour < 13 AND intHour > 11) then
			strBsClr = g_arColor(1)
		else
			strBsClr = g_arColor(6)
		end if
		
		strHTML = strHTML & "<td bgcolor='#" & strBsClr & "' height=" _
			& intHeight & " width='1%'>" _
			& "<img src='./images/tiny_blank.gif' height=1></td>" & VbCrLf
	next

	strHTML = strHTML & "</table></td></table>"
	
	makeWeek = strHTML
	
End Function

%>
