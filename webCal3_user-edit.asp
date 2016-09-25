<!--#include file="data/webCal3_data.inc"-->
<!--#include file="webCal3_verify.inc"-->

<html>
<head>
<!--#include file="webCal3_themes.inc"-->

<%
' Copyright 1999 Jason Abbott (jason@webott.com)
' Last updated 06/05/1999

dim rs, query

if Request.Form("delete") = "Delete" then
' ----------------------------------
' DELETION FORM
' ----------------------------------
' if the delete button was hit then display the deletion form
' get the info on the user to be deleted

	dim eventCount
	query = "SELECT * FROM cal_users WHERE " _
		& "(user_id)=" & Request.Form("user_id")
	Set rs = db.Execute(query,,&H0001)
		nameFirst = rs("name_first")
	rs.Close
	Set rs = nothing
	
	query = "SELECT user_id FROM cal_events WHERE " _
		& "user_id=" & Request.Form("user_id")
	Set rs = Server.CreateObject("ADODB.RecordSet")

' DSN was defined by data include
' adOpenStatic = 3
' adLockReadOnly = 1
' adCmdText = &H0001

	rs.Open query, DSN, 3, 1, &H0001
	eventCount = rs.Recordcount
%>

</head>
<body bgcolor="#<%=color(1)%>" link="#<%=color(7)%>" vlink="#<%=color(7)%>" alink="#<%=color(6)%>">
<center>

<!-- framing table -->
<table bgcolor="#<%=color(6)%>" border=0 cellpadding=2 cellspacing=0><tr><td>
<!-- end framing table -->

<table bgcolor="#<%=color(11)%>" border=0 cellpadding=3 cellspacing=0>
<form action="webCal3_user-deleted.asp" method="post">
<tr>
	<td bgcolor="#<%=color(4)%>" colspan=3>
	<font face="Tahoma, Arial, Helvetica" size=4>
	<b>User Deletion</b></font></td>
<tr>
	<td colspan=3><font face="Tahoma, Arial, Helvetica" size=2>
	
<% if eventCount > 0 then %>
	
	<b>What should happen to the <%=eventCount%> events scheduled by <%=nameFirst%>?</b><br>
	<input type="radio" name="do" value="delete">erase them all<br>
	<input type="radio" name="do" value="some" checked>erase the private but transfer the public events<br>
	<input type="radio" name="do" value="move">transfer them all<br>
	<center>transfer to <select name="recipient">
<%
		dim showName

		query = "SELECT user_id, name_first, name_last FROM cal_users " _
			& "ORDER BY name_last, name_first"
		Set rs = db.Execute(query,,&H0001)
		do while not rs.EOF
			if rs("name_last") <> "" then
				showName = rs("name_last") & ", " & rs("name_first")
			else
				showName = rs("name_first")
			end if
			response.write "<option value=" & rs("user_id") _
				& ">" & showName & VbCrLf
			rs.MoveNext
		loop
%>
	</select></center>
	</font></td>
	
<% else %>

Are you sure you want to erase the user <%=nameFirst%>?

<% end if %>
	
<tr>
	<td colspan=3 align="center" bgcolor="#<%=color(12)%>">
		<input type="submit" name="delete" value="Continue">
		<input type="submit" name="cancel" value="Cancel">
	</td>
<tr>
	<td align="center" colspan=3><font face="Tahoma, Arial, Helvetica" size=2>
	<b><font color="#cc0000">Caution</font>: erased users and events cannot be restored</b></font></td>
</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->

<input type="hidden" name="user_id" value="<%=Request.Form("user_id")%>">
<input type="hidden" name="event_count" value="<%=eventCount%>">
<input type="hidden" name="view" value="<%=Request.QueryString("view")%>">
</form>

<%
	rs.Close
	Set rs = nothing
else
' ----------------------------------
' EDIT FORM
' ----------------------------------
' if any button other than delete was hit, display the edit form

	dim nameFirst, nameLast, accessLevel, emailName, emailSite
	dim login, password, editType

	if Request.Form("edit") = "Edit" then

' we're editing an existing event
' ----------------------------------
' get existing data
' ----------------------------------
		editType = "update"
		query = "SELECT * FROM cal_users WHERE " _
			& "(user_id)=" & Request.Form("user_id")
		Set rs = db.Execute(query,,&H0001)
			nameFirst = rs("name_first")
			nameLast = rs("name_last")
			accessLevel = rs("access_level")
			emailName = rs("email_name")
			emailSite = rs("email_site")
			login = rs("login")
			password = rs("password")
		rs.Close
		Set rs = nothing
	else

' otherwise create a new event
' ----------------------------------
' prepare new data
' ----------------------------------

		editType = "new"
		nameFirst = ""
		nameLast = ""
		accessLevel = "user"
		emailName = ""
		emailSite = ""
		login = ""
		password = ""
	end if

' now include the JavaScript for the popup calendar
' and populate the edit form with values
%>

<SCRIPT LANGUAGE="javascript">
<!--
function Validate() {
	if (document.editform.name_first.value.length <= 0) {
		alert("You must enter a first name");
		document.editform.name_first.select();
		document.editform.name_first.focus();
		return false;
	}
	if (document.editform.login.value.length <= 0) {
		alert("You must enter a login name");
		document.editform.login.select();
		document.editform.login.focus();
		return false;
	}
	if (document.editform.password.value.length <= 0) {
		alert("You must enter a password");
		document.editform.password.select();
		document.editform.password.focus();
		return false;
	}
	if (document.editform.password.value != document.editform.confirm.value) {
		alert("The password values do not match");
		document.editform.confirm.select();
		document.editform.confirm.focus();
		return false;
	}
}
//-->
</SCRIPT>
</head>
<body bgcolor="#<%=color(1)%>" link="#<%=color(7)%>" vlink="#<%=color(7)%>" alink="#<%=color(6)%>">
<center>

<!-- framing table -->
<table bgcolor="#<%=color(6)%>" border=0 cellpadding=2 cellspacing=0><tr><td>
<!-- end framing table -->

<table bgcolor="#<%=color(11)%>" border=0 cellpadding=4 cellspacing=0>
<form name="editform" method="post" id="event" action="webCal3_user-updated.asp">
<tr bgcolor="#<%=color(4)%>" valign="bottom">
	<td colspan=2><font face="Tahoma, Arial, Helvetica" size=4>
	<b>User Details</b></font></td>
<tr>
	<td align="right" bgcolor="#<%=color(12)%>">
	<font face="Tahoma, Arial, Helvetica" size=2 color="#bb0000">First Name</font></td>
	<td><input type="text" name="name_first" value="<%=nameFirst%>" size=15></td>
<tr>
	<td align="right" bgcolor="#<%=color(12)%>">
	<font face="Tahoma, Arial, Helvetica" size=2>Last Name</font></td>
	<td><input type="text" name="name_last" value="<%=nameLast%>" size=15></td>
<tr>
	<td align="right" bgcolor="#<%=color(12)%>">
	<font face="Tahoma, Arial, Helvetica" size=2>e-mail</td>
	<td><input type="text" name="email_name" value="<%=emailName%>" size=10>@<input type="text" name="email_site" value="<%=emailSite%>" size=15></td>
<tr>
	<td align="right" bgcolor="#<%=color(12)%>">
	<font face="Tahoma, Arial, Helvetica" size=2 color="#bb0000">Login Name</font></td>
	<td><input type="text" name="login" value="<%=login%>" size=15></td>
<tr>
	<td align="right" bgcolor="#<%=color(12)%>">
	<font face="Tahoma, Arial, Helvetica" size=2 color="#bb0000">Password<br>
	confirm</font></td>
	<td><input type="password" name="password" value="<%=password%>" size=15><br>
	<input type="password" name="confirm" value="<%=password%>" size=15></td>
<tr>
	<td align="right" bgcolor="#<%=color(12)%>">
	<font face="Tahoma, Arial, Helvetica" size=2>Access</font></td>
	<td><select name="access_level">
	<option value="user"<% if accessLevel = "user" then%> selected<%end if%>>User
	<option value="admin"<% if accessLevel = "admin" then%> selected<%end if%>>Administrator
	</select>
	</td>
<tr>
	<td colspan=2 align="center" bgcolor="#<%=color(12)%>">
		<input type="submit" name="save" value="Save" onClick="return Validate();">
		<input type="submit" name="saveadd" value="Save & Add Another" onClick="return Validate();">
      <input type="submit" name="cancel" value="Cancel">
	</td>

</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->

<input type="hidden" name="edit_type" value="<%=editType%>">
<input type="hidden" name="user_id" value="<%=Request.Form("user_id")%>">
<input type="hidden" name="url" value="<%=Request.Form("url")%>">
<input type="hidden" name="view" value="<%=Request.QueryString("view")%>">
</form>

<font face="Tahoma, Arial, Helvetica" size=2 color="#bb0000">Red fields are required</font>

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