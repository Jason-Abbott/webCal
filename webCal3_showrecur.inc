<%
' Copyright 1999 Jason Abbott (jason@webott.com)
' Last updated 06/04/1999

' this script creates specialized views of the different
' recurrence types

' it expects the following things:
' 1) the pre-existence of dateList
' 2) the pre-existence of eventID, eventRecur
' 3) the showStatus function already defined by container
' 4) these declarations should be in the container script:
'    dim monthList(12), eventYear, x, years, dateList


Set years = CreateObject("Scripting.Dictionary")

for x = 0 to UBound(dateList)
	eventYear = Year(dateList(x))
	if years.Exists(eventYear) then
		yearList = years.Item(eventYear)
		years.Remove(eventYear)
		years.Add eventYear, yearList & " " & dateList(x)
			
' insert delimiter for subsequent splitting by month
			
	else
		years.Add eventYear, dateList(x)
	end if
next

' now generate the text that describes the type of recurrence
' for each event and a view recurrence dates

select case eventRecur
'------------------------------------------------	
	case "yearly"
'------------------------------------------------
' if there's yearly recurrence then create a list
' of those years

		response.write "<font face=""Verdana, Arial, Helvetica"" size=2>" _
			& "scheduled yearly on " _
			& MonthName(Month(dateList(0))) _
			& " " & Day(dateList(0)) _
			& " during these years:" _
			& "<select onChange=""location.href='" _
			& "webCal3_detail.asp?event_id=" _
			& eventID & "&view=" & Request("view") _
			& "&date=" & Month(dateList(0)) & "/" _
			& Day(dateList(0)) & "/' + " _
			& "options[selectedIndex].value;"">" & VbCrLf
		for each eventYear in years.Keys
			response.write "<option value=" & eventYear
			if eventYear = Year(eventDate) then
				response.write " selected"
			end if
			response.write ">" & eventYear & VbCrLf
		next
		response.write "</select></font>" & VbCrLf
			
'------------------------------------------------
	case "monthly"
'------------------------------------------------
		response.write "<font face=""Verdana, Arial, Helvetica"" size=1>" _
			& "scheduled for the " & Day(dateList(0))

' generate the correct wording for the day number
' (no detail overlooked)
				
		select case Right(Day(dateList(0)), 1)
			case 1
				response.write "st"
			case 2
				response.write "nd"
			case 3
				response.write "rd"
			case else
				response.write "th"
		end select
		response.write " of these months:<br>"

' go through each year
			
		for each eventYear in years.Keys
			
' for each year create an array of scheduled months
				
			for each x in split(years.Item(eventYear))
				monthList(Month(x)) = MonthName(Month(x))
			next
				
' then display each item in the month array

			response.write "<table><tr><td>" & VbCrLf _
				& "<select onChange=""location.href=" _
				& "'webCal3_detail.asp?event_id=" _
				& eventID & "&view=" & Request("view") _
				& "&date=' + options[selectedIndex].value + " _
				& "'/" & Day(dateList(0)) & "/" & eventYear & "';"">" & VbCrLf
			for each x in monthList
				if x <> "" then
					response.write "<option value=" & x
					if x = MonthName(Month(eventDate)) then
						response.write " selected"
					end if
					response.write ">" & x & VbCrLf
				end if
			next
			Erase monthList
			response.write "</select></td>" & VbCrLf _
				& "<td><font face=""Verdana, Arial, Helvetica"" size=2>" _
				& eventYear & "</font></td>" & VbCrLf & "</table>"
		next
		
'------------------------------------------------
	case "2weeks"
'------------------------------------------------
' the formatting for the three day based recurrence types
' (2weeks, weekly and daily) are handled together, below,
' so there is only basic formatting in this case statement

		response.write "<font face=""Verdana, Arial, Helvetica"" size=1>" _
			& "scheduled for every other " _
			& WeekDayName(WeekDay(dateList(0))) _
			& " on these dates:<br>"
			
'------------------------------------------------
	case "weekly"
'------------------------------------------------
		response.write "<font face=""Verdana, Arial, Helvetica"" size=1>" _
			& "scheduled for every " _
			& WeekDayName(WeekDay(dateList(0))) _
			& " on these dates:<br>"
			
'------------------------------------------------
	case "daily"
'------------------------------------------------
		response.write "<font face=""Verdana, Arial, Helvetica"" size=1>" _
			& "scheduled daily for these dates:<br>"
			
'------------------------------------------------
	case "none"
'------------------------------------------------
		response.write "<font face=""Verdana, Arial, Helvetica"" size=2>" _
			& "<a href=""webCal3_detail.asp?event_id=" _
			& eventID & "&view=" & Request("view") _
			& "&date=" & dateList(0) & """ " _
			& showStatus("View details") & ">" _
			& MonthName(Month(dateList(0))) & " " _
			& Day(dateList(0)) & ", " _
			& Year(dateList(0)) & "</a></font>"
end select

'------------------------------------------------
' now finish the formatting for the day based events
'------------------------------------------------
if eventRecur = "daily" OR _
	eventRecur = "weekly" OR _
	eventRecur = "2weeks" then
' this is the biggie
' for every event, go through each year of occurence and break
' out dates by month

	for each eventYear in years.Keys
		response.write "<table width=""100%""><tr>" & VbCrLf _
			& "<td valign=""top"" align=""center"">" _
			& "<font face=""Verdana, Arial, Helvetica"" size=2>" _
			& eventYear & "</font></td>" & VbCrLf
		
' go through each event date in that year and create an array
' of days for each month

		for each x in split(years.Item(eventYear))
			if monthList(Month(x)) <> "" then
				combine = "<br>"
			else
				combine = ""
			end if
				
' now use that array to make a simple list of the days on
' which the event occurs for every month
				
			monthList(Month(x)) = monthList(Month(x)) & combine _
				& "<a href=""webCal3_detail.asp?event_id=" _
				& eventID & "&view=" & Request("view") _
				& "&date=" & DateSerial(eventYear,Month(x),Day(x)) _
				& """ " & showStatus("View " _
					& MonthName(Month(x)) & " " & Day(x) & ", " _
					& eventYear & " detail") _
				& ">" & Day(x) & "</a>"
		next

' create month headings and make them link to the calendar
' view for that month and year
		
		response.write "<td><table width=""100%"" " _
			& "cellpadding=2 cellspacing=0 border=0>" _
			& "<tr>" & VbCrLf
				
		for x = 1 to 12
			response.write "   <td align=""center"" " _
				& "width=""8.3%"" bgcolor=""#" & color(0) & """>" _
				& "<font face=""Tahoma, Arial, Helvetica"" size=1"
			if monthList(x) <> "" then
				response.write "><a href=""webCal3_month.asp?date=" _
					& DateSerial(eventYear, x, 1) & """ " _
					& showStatus("View " _
						& MonthName(x) & " " & eventYear _
						& " calendar") _
					& ">" & MonthName(x,1) & "</a>"
			else
				response.write " color=""#" & color(5) & """>" & MonthName(x,1)
			end if
			response.write "</font></td>" & VbCrLf
		next
			
		response.write "<tr>" & VbCrLf
		for x = 1 to 12
			response.write "   <td valign=""top"" align=""right"">" _
				& "<font face=""Verdana, Arial, Helvetica"" size=1>" _
				& monthList(x) & "</font></td>" & VbCrLf
		next
		response.write "</table></td>" & VbCrLf & "</table>"
		Erase monthList
		Erase dateList
	next
end if
years.RemoveAll	
%>