<input type="submit" value="Show">
<select name="interval">
<%
for each x in Array(60,30,15,5)
	response.write "<option"
	if CInt(intErval) = x then
		response.write " selected"
	end if
	response.write ">" & x & VbCrLf
next
%>
</select>
minute intervals from

<select name="range1">
<option value=0>12:00
<%
for x = 1 to 11
	response.write "<option value=" & x
	if intRange1/intFactor = x then
		response.write " selected"
	end if
	response.write ">" & x & ":00"
next
%>
</select>
AM to
<select name="range2">
<%
for x = 13 to 23
	response.write "<option value=" & x
	if (intRange2 + 1)/intFactor - 1 = x then
		response.write " selected"
	end if
	response.write ">" & x - 12 & ":59"
next
%>
</select>
PM