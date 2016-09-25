<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./data/webCal4_data.inc"-->
<!--#include file="./include/webCal4_constants.inc"-->
<!--#include file="./include/webCal4_verify.inc"-->
<!--#include file="./include/webCal4_settings.inc"-->
<!--#include file="./include/webCal4_functions.inc"-->
<!--#include file="./include/webCal4_functions-event.inc"-->
<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' Last updated 2/24/2001

dim m_arRecurType		' types of recurrence
dim m_arRecurName		' descriptions of recurrence types
dim m_arHourName		' hours for selection list
dim m_strRecurList		' list of recurrence options
dim m_strStartHrList	' list of hours
dim m_strEndHourList
dim m_strStartMinList	' list of minutes
dim m_strEndMinList
dim m_strJSEventScopes

' arrays to populate form lists
m_arRecurType = Array("none","daily","weekly","2weeks","monthly","yearly")
m_arRecurName = Array("None","Daily","Weekly","Every other wk","Monthly","Yearly")
m_arHourName = Array("12 AM","1 AM","2 AM","3 AM","4 AM","5 AM","6 AM","7 AM","8 AM","9 AM","10 AM","11 AM","12 PM","1 PM","2 PM","3 PM","4 PM","5 PM","6 PM","7 PM","8 PM","9 PM","10 PM","11 PM")

Call getEventFields(Request.Form("fldEdit"), Request.Form("fldEventID"), _
	Request.Form("fldScope"), Request.Form("fldDate"), m_strJSEventScopes)

' generate the recurrence options
for x = 0 to UBound(m_arRecurType)
	m_strRecurList = m_strRecurList & "<option value='" & m_arRecurType(x) _
		& "'>" & m_arRecurName(x) & VbCrLf
next
m_strRecurList = makeSelected(m_strRecurList, m_strRecur)

' generate the hours lists
for x = 0 to 23
	m_strStartHrList = m_strStartHrList & "<option value='" _
		& x & "'>" & m_arHourName(x) & VbCrLf
next
m_strEndHourList = makeSelected(m_strStartHrList, m_intEndHour)
m_strStartHrList = makeSelected(m_strStartHrList, m_intStartHour)

' generate the minutes form list
for x = 0 to 55 step 5
	if x < 10 then x = "0" & x
	m_strStartMinList = m_strStartMinList & "<option value='" _
		& x & "'>:" & x & VbCrLf
next
m_strEndMinList = makeSelected(m_strStartMinList, m_strEndMin)
m_strStartMinList = makeSelected(m_strStartMinList, m_strStartMin)
%>
<html>
<head>
<style><!--#include file="./style/webCal4_common.css"--></style>
<script language="javascript" src="./script/webCal4_functions.js"></script>
<script language="javascript" src="./script/webCal4_validate.js"></script>
<script language="javascript" src="./script/webCal4_functions-event.js"></script>
<script language="javascript">
var m_arEditScopes = [<%=m_strJSEventScopes%>];
</script>
</head>
<body onLoad="initPage();" bgcolor="#<%=g_arColor(1)%>" link="#<%=g_arColor(7)%>" vlink="#<%=g_arColor(7)%>" alink="#<%=g_arColor(6)%>">
<center>

<!-- framing table -->
<table bgcolor="#<%=g_arColor(6)%>" border=0 cellpadding=2 cellspacing=0><tr><td>
<!-- end framing table -->

<table bgcolor="#<%=g_arColor(11)%>" border=0 cellpadding=4 cellspacing=0>
<form name="frmEdit" method="post">

<tr bgcolor="#<%=g_arColor(3)%>" valign="bottom">
	<td colspan=2><font face="<%=g_arFont(0)%>" size=4>
	<b>Event Details</b></font></td>
<tr>
	<td valign="top"><b><font face="<%=g_arFont(0)%>" color="#<%=g_arColor(14)%>" size=3>Title</font></b><br>
		<input name="fldTitle" type="text" size="35" max="50" value="<%=m_strTitle%>">
	</td>
	<td rowspan=2 width=256 valign="top"><font face="<%=g_arFont(0)%>" color="#<%=g_arColor(14)%>"><b>Description</b></font><br>
		<textarea cols="24" name="fldDescription" type="text" rows="14" wrap="virtual"><%=m_strDescription%></textarea>
	</td>
<tr>
	<td valign="top">

<!-- timing table -->

	<table cellpadding=2 cellspacing=2 border=0 width="100%">
	<tr>
		<td bgcolor="#<%=g_arColor(12)%>"><font face="<%=g_arFont(0)%>">
			<font color="#<%=g_arColor(14)%>" size=3><b>Date</b></font>
			<br>
			<input name="fldStartDate" type="text" size="10" value="<%=m_strStartDate%>"><font size=2><input type="button" value="&gt;" onClick="calPop('frmEdit','fldStartDate');">
			<br>
			Recurrence<br>
			<select name='fldEventRecur' onChange='newRecur(this);'>
			<%=m_strRecurList%></select><br>
			until</font><br>
			<input name="fldEndDate" type="text" size="10" value="<%=m_strEndDate%>"><font size=2><input type="button" value="&gt;" onClick="calPop('frmEdit','fldEndDate');"></font>
			<br>
			<input type="checkbox" name="fldSkipWE"<%=m_strSkipWE%>><font size=2>Skip weekends</font></font>
		</font>
		</td>

		<td valign="top" bgcolor="#<%=g_arColor(12)%>"><font face="<%=g_arFont(0)%>">
			<font color="#<%=g_arColor(14)%>" size=3><b>Time</b></font>
			<font size=2>
			<br>
			<nobr>
			<select name="fldStartHour"<%=m_strShowTime%>>
			<%=m_strStartHrList%>
			</select>
			<select name="fldStartMin"<%=m_strShowTime%>>
			<%=m_strStartMinList%>
			</select>
			</nobr>
			<br>until<br>

			<nobr>
			<select name="fldEndHour"<%=m_strShowTime%>>
			<%=m_strEndHourList%>
			</select>
			<select name="fldEndMin"<%=m_strShowTime%>>
			<%=m_strEndMinList%>
			</select>
			</nobr>
			<p>
			<input type="checkbox" name="fldNoTime"<%=m_strNoTime%>
			onClick="newTimeCheck(this);">No Specific Time</font>
		</td>

<!-- end timing table -->

	<tr>
		<td colspan=2 bgcolor="#<%=g_arColor(12)%>"><font face="<%=g_arFont(0)%>">
			<font color="#<%=g_arColor(14)%>" size=3><b>Display</b></font><br>

			<font size=2>
			<table cellspacing=1 cellpadding=0 border=0 width="100%">
			<tr>
				<td align="center" valign="bottom" bgcolor="#000000">
				<input type="radio" name="fldEventColor" value="black"
				<%if m_strEventClr = "black" then%>checked<%end if%>></td>

				<td align="center" valign="bottom" bgcolor="#0000ff">
				<input type="radio" name="fldEventColor" value="blue"
				<%if m_strEventClr = "blue" then%>checked<%end if%>></td>

				<td align="center" valign="bottom" bgcolor="#aa00aa">
				<input type="radio" name="fldEventColor" value="purple"
				<%if m_strEventClr = "purple" then%>checked<%end if%>></td>

				<td align="center" valign="bottom" bgcolor="#ff0000">
				<input type="radio" name="fldEventColor" value="red"
				<%if m_strEventClr = "red" then%>checked<%end if%>></td>

				<td align="center" valign="bottom" bgcolor="#00cc00">
				<input type="radio" name="fldEventColor" value="green"
				<%if m_strEventClr = "green" then%>checked<%end if%>></td>

				<td align="center" valign="bottom" bgcolor="#ffbb00">
				<input type="radio" name="fldEventColor" value="orange"
				<%if m_strEventClr = "orange" then%>checked<%end if%>></td>
			</table>
			<center>
			In
			<select name="fldGroup" onChange="newGroup(this);"></select>,
			visible to
			<select name="fldShowTo" onChange="newUserScope(this);">
			<option value="0">none
			<option value="1">only me
			<option value="2">group
			<option value="3">public
			</select>

			</font>
			</center>
		</td>
	</table>
	</td>
<tr>
	<td colspan=2 align="center">
	<input type="button" value="Save" onClick='saveEvent("frmEdit", false);'>
	<input type="button" value="Save & Add Another"	onClick='saveEvent("frmEdit", true);'>
	<input type="button" value="Cancel" onClick='goPage("<%=Request.Form("fldURL")%>","frmEdit");'>
	</td>
</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->

<input type="hidden" name="fldUserScopes" value="0">
<input type="hidden" name="fldEditType" value="<%=m_strEditType%>">
<input type="hidden" name="fldEventID" value="<%=Request.Form("fldEventID")%>">
<input type="hidden" name="fldURL" value="<%=Request.Form("fldURL")%>">
<input type="hidden" name="fldView" value="<%=Request.QueryString("view")%>">
</form>

</center>
</body>
</html>