<%
' determine values for the month and year
' of the next and previous calendar pages

if m = 12 then
	mNext = 1
	mPrev = m - 1
	yNext = y + 1
	yPrev = y
elseif m = 1 then
	mNext = m + 1
	mPrev = 12
	yNext = y
	yPrev = y - 1
else
	mNext = m + 1
	mPrev = m - 1
	yNext = y
	yPrev = y
end if
%>