<%
' Copyright 1997 Jason Abbott (jabbott@uidaho.edu)
' Last updated 11/17/97

' Determine values for the month and year of the next
' and previous calendar pages

dim m,y,mnext,mprev,ynext,yprev

sub cal_nav (month,year)
	m = month
	y = year
	if m = 12 then
		mnext = 1
		mprev = m - 1
		ynext = y + 1
		yprev = y
	elseif m = 1 then
		mnext = m + 1
		mprev = 12
		ynext = y
		yprev = y - 1
	else
		mnext = m + 1
		mprev = m - 1
		ynext = y
		yprev = y
	end if
end sub
%>