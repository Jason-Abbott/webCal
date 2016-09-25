<%
' Copyright 1999 Jason Abbott (webcal@webott.com)
' updated 09/13/1999


' if a conflicting event was found, show it

	do while not rsEvents.EOF
		response.write rsEvents("event_title") & "<br>"
		rsEvents.MoveNext
	loop

%>