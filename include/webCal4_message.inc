<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' Last updated 2/21/2001

' Displays a feedback message (updated 2/21/01)
' returns string ---------------------------------------------------------
Function showMsg(ByVal v_strMessage, ByVal v_strBorder, ByVal v_strBox, ByVal v_strFont)
	dim strAlign
	if Trim(v_strMessage) <> "" then
		if Len(v_strMessage) < 45 then
			strAlign = "center"
		else
			strAlign = "left"
		end if
		showMsg = "<layer><div id='msg'><p><center><table bgcolor='#" & v_strBorder _
			& "' border=0 cellpadding=2 cellspacing=0 width='50%'><tr><td>" _
			& "<table border=0 cellpadding=3 cellspacing=0 width='100%'>" _
			& "<tr><td bgcolor='#" & v_strBox & "' align='" & strAlign & "'>" _
			& "<font face='" & v_strFont & "' size=2><b>" _
			& v_strMessage & "</b></font>" _
			& "</td></table></td></table></center></div></layer>"
	else
		showMsg = "<layer><div id='msg'></div></layer>"
	end if
End Function
%>

<script language="javascript">
// hide message layer after specified interval (updated 2/22/01)
// updates dom -----------------------------------------------------------
function initMessage() {
	var lSecs = 5000;
	var sAction = (document.layers) ? "document.layers['msg'].visibility = 'hide';" : "msg.style.display = 'none';"
	window.setTimeout(sAction, lSecs);
}
</script>