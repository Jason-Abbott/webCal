<!--#include file="data/webCal3_data.inc"-->
<!--#include file="webCal3_verify.inc"-->

<html>
<head>
<!--#include file="webCal3_themes.inc"-->

<%
' Copyright 1999 Jason Abbott (jason@webott.com)
' Last updated 05/27/1999

dim rs, query, view

if Request.Form("delete") = "Delete" then
' ----------------------------------
' DELETION FORM
' ----------------------------------
' if the delete button was hit then display the deletion form
' get the info on the event to be deleted

	dim say
	query = "SELECT * FROM cal_events WHERE " _
		& "(event_id)=" & Request.Form("event_id")
	Set rs = db.Execute(query,,&H0001)

' how many did they want to delete?

	Select Case Request.Form("scope")
		Case "one"
			say = "this (" & Request.Form("date") & ") instance of"
		Case "future"
			say = "this (" & Request.Form("date") & ") and <b>all future</b> instances of"
		Case "all"
			say = "<b>all " & Request.Form("count") & "</b> instances of"
		Case else
			say = ""
	End Select
%>

</head>
<body bgcolor="#<%=color(1)%>" link="#<%=color(7)%>" vlink="#<%=color(7)%>" alink="#<%=color(6)%>">
<center>

<!-- framing table -->
<table bgcolor="#<%=color(6)%>" border=0 cellpadding=2 cellspacing=0><tr><td>
<!-- end framing table -->

<table bgcolor="#<%=color(11)%>" border=0 cellpadding=4 cellspacing=0>
<form action="webCal3_deleted.asp" method="post">
<tr bgcolor="#<%=color(3)%>" valign="bottom">
	<td colspan=3><font face="Tahoma, Arial, Helvetica" size=4>
	<b>Event Deletion</b></font></td>
<tr>
	<td align="center" colspan=3><font face="Arial, Helvetica">
	Are you sure you want to erase <%=say%> <i><%=rs("event_title")%></i>?</font></td>
<tr>
	<td colspan=3 align="center" bgcolor="#<%=color(12)%>">
		<input type="submit" name="delete" value="Yes">
		<input type="submit" name="cancel" value="No">
	</td>
<tr>
	<td align="center" colspan=3><font face="Tahoma, Arial, Helvetica" size=2>
	<b><font color="#cc0000">Caution</font>: erased events cannot be restored</b></font></td>
</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->

<input type="hidden" name="event_id" value="<%=Request.Form("event_id")%>">
<input type="hidden" name="date" value="<%=Request.Form("date")%>">
<input type="hidden" name="scope" value="<%=Request.Form("scope")%>">
<input type="hidden" name="url" value="<%=Request.Form("url")%>">
<input type="hidden" name="view" value="<%=Request.Form("view")%>">
</form>

<%
	rs.Close
	Set rs = nothing
else
' ----------------------------------
' EDIT FORM
' ----------------------------------
' if any button other than delete was hit, display the edit form

	dim eventTitle, eventDescription, eventRecur, startDate, endDate
	dim startHour, startMin, endHour, endMin, recurType, recurName
	dim x, hourName, editType, noTime, showTime

' arrays to populate form lists
' (I miss associative arrays)

	recurType = Array("none","daily","weekly","2weeks","monthly","yearly")
	recurName = Array("None","Daily","Weekly","Every other wk","Monthly","Yearly")
	hourName = Array("12 AM","1 AM","2 AM","3 AM","4 AM","5 AM","6 AM","7 AM","8 AM","9 AM","10 AM","11 AM","12 PM","1 PM","2 PM","3 PM","4 PM","5 PM","6 PM","7 PM","8 PM","9 PM","10 PM","11 PM")

' default values
' use military time

	startHour = 8
	startMin = "00"
	endHour = 17
	endMin = "00"
	noTime = ""
	showTime = ""
	skipWE = ""
	eventColor = "black"

	if Request.Form("edit") = "Edit" then

' we're editing an existing event
' ----------------------------------
' get existing data
' ----------------------------------

		query = "SELECT * FROM cal_events E INNER JOIN cal_dates D" _
			& " ON (E.event_id = D.event_id) " _
			& "WHERE (E.event_id)=" & Request.Form("event_id") _
			& " ORDER BY event_date"

		Set rs = Server.CreateObject("ADODB.RecordSet")

' DSN was defined by data include
' adOpenStatic = 3
' adLockReadOnly = 1
' adCmdText = &H0001

		rs.Open query, DSN, 3, 1, &H0001

		eventTitle = rs("event_title")
		eventDescription = rs("event_description")
		eventColor = rs("event_color")

' these need to be broken out for separate form fields

		if rs("time_start") <> "" then
			startHour = Hour(rs("time_start"))
			startMin = Minute(rs("time_start"))
			endHour = Hour(rs("time_end"))
			endMin = Minute(rs("time_end"))
		else
			noTime = " checked"
			showTime = " disabled"
		end if

		if rs("skip_weekends") = 1 then
			skipWE = " checked"
		else
			skipWE = ""
		end if

		if rs("private") = 1 then
			eventHide = " checked"
		else
			eventHide = ""
		end if

		Select Case Request.Form("scope")
			Case "future"
				eventRecur = rs("event_recur")
				startDate = Request.Form("date")
				rs.MoveLast
				endDate = DateValue(rs("event_date"))
			Case "all"
				eventRecur = rs("event_recur")
				startDate = DateValue(rs("event_date"))
				rs.MoveLast
				endDate = DateValue(rs("event_date"))
			Case else
				eventRecur = "none"
				startDate = Request.Form("date")
				endDate = ""
				skipWE = ""
		End Select

		if Request.Form("scope") <> "" then
			editType = Request.Form("scope")
		else
			editType = "all"
		end if

		view = Request.Form("view")
		rs.Close
		Set rs = nothing
	else

' otherwise create a new event
' ----------------------------------
' prepare new data
' ----------------------------------

		eventTitle = ""
		eventDescription = ""
		eventRecur = "none"
		startDate = Request.QueryString("date")
		endDate = ""
		editType = "new"
		view = Request.QueryString("view")

' if the user is currently hiding public events then assume
' this new event should be private

		if Session(dataName & "Public") then
			eventHide = ""
		else
			eventHide = " checked"
		end if
	end if

' now include the JavaScript for the popup calendar
' and populate the edit form with values
%>

<SCRIPT TYPE="text/javascript" LANGUAGE="javascript" SRC="webCal3_popup.js"></SCRIPT>
<SCRIPT LANGUAGE="javascript">
<!--
function Validate() {
	if (document.editform.title.value.length <= 0) {
		alert("You must enter a title for the event");
		document.editform.title.select();
		document.editform.title.focus();
		return false;
	}
}

function disableTime() {
// Netscape doesn't support the disabled property as of 4.08
	if (document.editform.notime.checked) {
		document.editform.start_hour.disabled=1;
		document.editform.start_min.disabled=1;
		document.editform.end_hour.disabled=1;
		document.editform.end_min.disabled=1;
	} else {
		document.editform.start_hour.disabled=0;
		document.editform.start_min.disabled=0;
		document.editform.end_hour.disabled=0;
		document.editform.end_min.disabled=0;
	}
}

function updateEnd() {
//	if (document.editform.end_date.value == "") {
		var r = document.editform.event_recur.options[document.editform.event_recur.selectedIndex].value;
		var d = document.editform.start_date.value;
		var day = d.split("/")[1];
		var month = d.split("/")[0];
		var year = d.split("/")[2];

		if (r == "none") {
			d = "";
		}
		if (r == "monthly") {
			if (month != 12) {
				month = month - 1 + 2;
			} else {
				month = 1;
				year = year - 1 + 2;
			}
			d = month + "/" + day + "/" + year;
		}
		if (r == "yearly") {
			year = year - 1 + 2;
			d = month + "/" + day + "/" + year;
		}
		document.editform.end_date.value = d;
//	}
}

//-->
</SCRIPT>
</head>
<body onload="init();" bgcolor="#<%=color(1)%>" link="#<%=color(7)%>" vlink="#<%=color(7)%>" alink="#<%=color(6)%>">
<center>

<!-- framing table -->
<table bgcolor="#<%=color(6)%>" border=0 cellpadding=2 cellspacing=0><tr><td>
<!-- end framing table -->

<table bgcolor="#<%=color(11)%>" border=0 cellpadding=4 cellspacing=0>
<form name="editform" method="post" id="event" action="webCal3_updated.asp">

<tr bgcolor="#<%=color(3)%>" valign="bottom">
	<td colspan=2><font face="Tahoma, Arial, Helvetica" size=4>
	<b>Event Details</b></font></td>
<tr>
	<td valign="top"><b><font face="Tahoma, Arial, Helvetica" color="#<%=color(14)%>" size=3>Title</font></b><br>
		<input name="title" id="title" type="text" size="35" max="50" value="<%=eventTitle%>">
	</td>
	<td rowspan=2 width=256 valign="top"><font face="Tahoma, Arial, Helvetica" color="#<%=color(14)%>"><b>Description</b></font><br>
		<textarea cols="24" name="description" type="text" rows="14" wrap="virtual"><%=eventDescription%></textarea>
	</td>
<tr>
	<td valign="top">

<!-- timing table -->

	<table cellpadding=2 cellspacing=2 border=0 width="100%">
	<tr>
		<td bgcolor="#<%=color(12)%>"><font face="Tahoma, Arial, Helvetica">
			<font color="#<%=color(14)%>" size=3><b>Date</b></font>
			<br>
			<input name="start_date" id="date" type="text" size="10" value="<%=startDate%>"><font size=2><input type="button" value="&gt;" onClick="calpopup('editform','start_date');">
			<br>
			Recurrence<br>
<%
' generate the recurrence options
' select the option that matches the current event

	response.write "<select name=""event_recur"" " _
		& "onChange=""updateEnd();"">"
	for x = 0 to UBound(recurType)
		response.write("<option value=""" & recurType(x) & """")
		if recurType(x) = eventRecur then
			response.write(" selected")
		end if
		response.write(">" & recurName(x) & VbCrLf)
	next
%>
			</select><br>
			until</font><br>
			<input name="end_date" id="recurend" type="text" size="10" value="<%=endDate%>"><font size=2><input type="button" value="&gt;" onClick="calpopup('editform','end_date');"></font>
			<br>
			<input type="checkbox" name="skip"<%=skipWE%>><font size=2>Skip weekends</font></font>
		</font>
		</td>

		<td valign="top" bgcolor="#<%=color(12)%>"><font face="Tahoma, Arial, Helvetica">
			<font color="#<%=color(14)%>" size=3><b>Time</b></font>
			<font size=2>
			<br>
			<nobr>
			<select name="start_hour"<%=showTime%>>
<%
' generate the hours form list and select the
' one assigned above

	for x = 0 to 23
		response.write("<option value=" & x)
		if x = startHour then
			response.write(" selected")
		end if
		response.write(">" & hourName(x) & VbCrLf)
	next
%>
			</select>
			<select name="start_min"<%=showTime%>>
<%
' generate the minutes form list and select the
' one assigned above

	for x = 0 to 55 step 5
		if x < 10 then
			x = "0" & x
		end if
		response.write("<option value=""" & x & """")
		if x = startMin then
			response.write(" selected")
		end if
		response.write(">:" & x & VbCrLf)
	next
%>
			</select>
			</nobr>
			<br>until<br>

			<nobr>
			<select name="end_hour"<%=showTime%>>
<%
' hours list

	for x = 0 to 23
		response.write("<option value=" & x)
		if x = endHour then
			response.write(" selected")
		end if
		response.write(">" & hourName(x) & VbCrLf)
	next
%>
			</select>
			<select name="end_min"<%=showTime%>>
<%
' minutes list

	for x = 0 to 55 step 5
		if x < 10 then
			x = "0" & x
		end if
		response.write("<option value=""" & x & """")
		if x = endMin then
			response.write(" selected")
		end if
		response.write(">:" & x & VbCrLf)
	next
%>
			</select>
			</nobr>
			<p>
			<input type="checkbox" name="notime"<%=noTime%>
			onClick="disableTime();">No Specific Time</font>
		</td>
	</table>

<!-- end timing table -->

	<font face="Tahoma, Arial, Helvetica" size=2>
	Display color:
	<table cellspacing=1 cellpadding=0 border=0 width="100%">
	<tr>
		<td align="center" valign="bottom" bgcolor="#000000">
		<input type="radio" name="event_color" value="black"
		<%if eventColor = "black" then%>checked<%end if%>></td>

		<td align="center" valign="bottom" bgcolor="#0000ff">
		<input type="radio" name="event_color" value="blue"
		<%if eventColor = "blue" then%>checked<%end if%>></td>

		<td align="center" valign="bottom" bgcolor="#aa00aa">
		<input type="radio" name="event_color" value="purple"
		<%if eventColor = "purple" then%>checked<%end if%>></td>

		<td align="center" valign="bottom" bgcolor="#ff0000">
		<input type="radio" name="event_color" value="red"
		<%if eventColor = "red" then%>checked<%end if%>></td>

		<td align="center" valign="bottom" bgcolor="#00cc00">
		<input type="radio" name="event_color" value="green"
		<%if eventColor = "green" then%>checked<%end if%>></td>

		<td align="center" valign="bottom" bgcolor="#ffbb00">
		<input type="radio" name="event_color" value="orange"
		<%if eventColor = "orange" then%>checked<%end if%>></td>
	</table>

	<input type="checkbox" name="private"<%=eventHide%>><font size=2>Private (visible only to you)
	</font>

	</td>
<tr>
	<td colspan=2 align="center">
		<input type="submit" name="save" value="Save" onClick="return Validate();">
		<input type="submit" name="saveadd" value="Save & Add Another" onClick="return Validate();">
      <input type="submit" name="cancel" value="Cancel">
	</td>
</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->

<input type="hidden" name="edit_type" value="<%=editType%>">
<input type="hidden" name="event_id" value="<%=Request.Form("event_id")%>">
<input type="hidden" name="url" value="<%=Request.Form("url")%>">
<input type="hidden" name="view" value="<%=view%>">
</form>

<script lang="javascript">
<!--
// focus on the title form element

	document.forms[0].elements[0].focus();

// -->
</script>

<%
' ----------------------------------
' END OF SEPERATE FORMS
' ----------------------------------
end if
db.Close
Set db = nothing
%>
</center>
</body>
</html>