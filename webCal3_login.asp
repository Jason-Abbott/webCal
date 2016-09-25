<%
' Copyright 1999 Jason Abbott (jason@webott.com)
' Last updated 05/27/1999

dim error, status, query, rs

status = "This action is available only to registered users"
error = "The information you entered could not be validated. " _
	& "Please try again."
%>
<!--#include file="data/webCal3_data.inc"-->
<%
if Request.Form("login") <> "" then
	query = "SELECT * FROM cal_users WHERE " _
		& "login = '" & Request.Form("login") & "'"
	Set rs = db.Execute(query,,&H0001)
	if rs.EOF = -1 then
		status = error
	else
		if rs("password") = Request.Form("password") then
			Session(dataName & "User") = rs("user_id")
			Session(dataName & "Access") = rs("access_level")
			rs.Close
			db.Close
			Set rs = nothing
			Set db = nothing
			response.redirect Request.Form("url")
		else
			status = error
		end if
	end if
	rs.Close : Set rs = nothing
	db.Close : Set db = nothing
end if

' obtain the administrator's e-mail address
'query = "SELECT email_name, email_site, user_id " _
'	& "FROM cal_users WHERE (user_id) = 1"
'Set rs = db.Execute(query,,&H0001)
%>

<html>
<!--#include file="webCal3_themes.inc"-->
<body bgcolor="#<%=color(1)%>" link="#<%=color(7)%>" vlink="#<%=color(7)%>" alink="#<%=color(6)%>">
<center>

<!-- framing table -->
<table bgcolor="#<%=color(6)%>" width="60%" border=0 cellpadding=2 cellspacing=0><tr><td>
<!-- end framing table -->

<table bgcolor="#<%=color(11)%>" border=0 cellpadding=3 cellspacing=0 width="100%">
<form action="webCal3_login.asp" method="post">
<tr bgcolor="#<%=color(4)%>" valign="bottom">
	<td colspan=4><font face="Tahoma, Arial, Helvetica" size=4>
	<b>Login</b></font></td>
<tr>
	<td colspan=4 align="center"><font face="Arial, Helvetica" size=2>
	<%=status%><br></font></td>
<tr>
	<td>&nbsp;</td>
	<td bgcolor="#<%=color(12)%>" align="right"><font face="Arial, Helvetica">Username:&nbsp;</td>
	<td bgcolor="#<%=color(12)%>"><input type="text" name="login" size=10></td>
	<td>&nbsp;</td>
<tr>
	<td>&nbsp;</td>
	<td bgcolor="#<%=color(12)%>" align="right"><font face="Arial, Helvetica">Password:&nbsp;</td>
	<td bgcolor="#<%=color(12)%>"><input type="password" name="password" size=10></td>
	<td>&nbsp;</td>
<tr>
	<td colspan=4 align="center"><!-- <font face="Verdana, Arial, Helvetica" size=2>
	<a href="mailto:<%'rs("email_name")%>@<%'rs("email_site")%>">Request an account</a></font>
	<br> -->
	<input type="submit" value="Continue"></td>
</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->

<input type="hidden" name="url" value="<%=Request("url")%>">
</form>

</center>
</body>
</html>