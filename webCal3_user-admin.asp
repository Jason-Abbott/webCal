<!--
Copyright 1999 Jason Abbott (jason@webott.com)
Last updated 06/05/1999
 -->
 
<!--#include file="./data/webCal3_data.inc"-->
<!--#include file="webCal3_verify.inc"-->

<html>
<head>
<script language="javascript"><!--
//preload images and text for faster operation

if (document.images) {
// back to calendar icon
	var iconMonth = new Image();
	iconMonth.src = "images/icon_calprev_grey.gif";
	var iconMonthOn = new Image();
	iconMonthOn.src = "images/icon_calprev.gif"
	statusMonth = "Return to calendar";
}

function iconOver(name){
	if (document.images) {
  		document.images[name].src=eval("icon"+name+"On.src");
		status=eval("status"+name);
	}
}

function iconOut(name){
	if (document.images) {
  		document.images[name].src=eval("icon"+name+".src");
		status="";
	}
}
//-->
</script>

<!--#include file="webCal3_themes.inc"-->
<body bgcolor="#<%=color(1)%>" link="#<%=color(7)%>" vlink="#<%=color(7)%>" alink="#<%=color(6)%>">
<center>

<!-- framing table -->
<table bgcolor="#<%=color(6)%>" border=0 cellpadding=2 cellspacing=0><tr><td>
<!-- end framing table -->

<table bgcolor="#<%=color(11)%>" border=0 cellpadding=3 cellspacing=0>
<form name="adminform" action="webCal3_user-edit.asp?view=<%=Request.QueryString("view")%>" method="post">
<tr>
	<td bgcolor="#<%=color(4)%>" colspan=2>
	<a href="webCal3_<%=Request.QueryString("view")%>.asp"
	onMouseOver="iconOver('Month'); return true;" 
   onMouseOut="iconOut('Month'); return true;">
	<img name="Month" src="./images/icon_calprev_grey.gif" width=15 height=16 alt="" border=0></a>
	<font face="Tahoma, Arial, Helvetica" size=4>
	<b>User Management</b></font></td>
<tr>
	<td align="right" valign="top" bgcolor="#<%=color(12)%>">
<input type="submit" name="add" value="Add">
	</td>
	<td><font face="Verdana, Arial, Helvetica" size=2>
	a new user</font></td>
<tr>
	<td align="right" valign="top" bgcolor="#<%=color(12)%>">
<input type="submit" name="edit" value="Edit">
	</td>
	<td><font face="Verdana, Arial, Helvetica" size=2>
	the selected user</font></td>
<tr>
	<td align="right" valign="top" bgcolor="#<%=color(12)%>">
<input type="submit" name="delete" value="Delete">
	</td>
	<td><font face="Verdana, Arial, Helvetica" size=2>
	the selected user</font></td>
<tr>
	<td bgcolor="#<%=color(12)%>" align="right">
	<font face="Verdana, Arial, Helvetica" size=2>select</font></td>
	<td bgcolor="#<%=color(12)%>">
	<select name="user_id">
<%
dim query, rs, showName

query = "SELECT user_id, name_first, name_last FROM cal_users " _
	& "WHERE user_id <> 1 ORDER BY name_last, name_first"
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
rs.Close
db.Close
Set rs = nothing
Set db = nothing
%>

	</select>
	</td>
</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->

<input type="hidden" name="view" value="<%=Request.QueryString("view")%>">
</form>

</center>
</body>
</html>