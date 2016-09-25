<%
' this updates the status bar with a message

function showStatus(strMsg)
	strMsg = Replace(strMsg, "'", "\'")
	showStatus = "onMouseOver=""status='" & strMsg & "'; " _
		& "return true;"" " & VbCrLf _
		& "onMouseOut=""status=''; return true;"""
end function

' this switches an icon

function switchIcon(strName, strSource)
	if strSource = "" then
		strSource = strName
	end if

	switchIcon = "onMouseOver=""iconOver('" & strName _
		& "','" & strSource & "'); return true;"" onMouseOut=""iconOut('" _
		& strName & "','" & strSource & "'); return true;"""
end function
%>