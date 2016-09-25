<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' Last updated 02/23/2000

dim intErval		' number of minutes (time interval) per table cell (segment)
dim intFactor
dim intSegments		' number of cells (segments) per column
dim intSegStart		' week start time measured in segments
dim intSegEnd		' ending time measured in segments
dim intSegMax		' total segments per day
dim intHeight		' cell height in pixels
dim bWeekend		' display weekend?

if Request.Cookies("week")("interval") <> "" then
	' read saved cookie values
	intErval = Request.Cookies("week")("interval")
	intFactor = 60/intErval	
	intSegStart = Request.Cookies("week")("range1") * intFactor
	intSegStart = Request.Cookies("week")("range2") * intFactor - 1
	bWeekend = Request.Cookies("week")("weekend")
else
	' default values
	intErval = 15
	intFactor = 60/intErval
	intSegStart = 6 * intFactor
	intSegEnd = 22 * intFactor - 1
	bWeekend = 1
end if

intSegments = intFactor

' ZERO-BASED count of time segements/day
intSegMax = (1440/intErval) - 1
intHeight = 24/intFactor - 1

' this function takes a time and converts it to the
' proper number of table segments based on the
' specified interval
Function segments(strTime) '----------------------------------------------
	dim intMin		' minutes past the hour of start time
	dim intAdd		'
	dim x			' loop counter

	intMin = Minute(strTime)
	intAdd = intErval/2
	for x = 0 to 60/intErval - 1
		if intMin < intAdd then
			intSegments = x
			exit for
		end if
		intAdd = intAdd + intErval
	next
	segments = intSegments + (Hour(strTime) * intFactor)
End Function


Function hexColor (strNameClr) '------------------------------------------
	dim strColor
	select case strNameClr
		case "blue"
			strColor = "0000ff"
		case "purple"
			strColor = "aa00aa"
		case "red"
			strColor = "ff0000"
		case "green"
			strColor = "00cc00"
		case "orange"
			strColor = "ffcc44"
		case else
			strColor = "ffffff"
	end select
	hexColor = strColor
End Function
%>