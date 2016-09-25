<!--
Copyright 2000 Jason Abbott (webcal@webott.com)
Last updated 2/25/2000
-->

<html>
<head>
<SCRIPT TYPE="text/javascript" LANGUAGE="javascript" SRC="./include/webCal4_popup.js"></SCRIPT>
<SCRIPT LANGUAGE="javascript">
<!--

function Validate() {
	var oForm = document.findform;
	if (oForm.title.value.length <= 0
     && oForm.description.value.length <= 0
	 && oForm.date_start.value.length <= 0
	 && oForm.date_end.value.length <= 0) {
		alert("You must enter criteria in at least one field");
		oForm.title.select();
		oForm.title.focus();
		return false;
	}
}
//-->
</SCRIPT>

<!--#include file="./include/webCal4_settings_inc.asp"-->
<!--#include file="./include/webCal4_constants_inc.asp"-->
</head>
<body onload="init();" bgcolor="#<%=g_arColor(1)%>" link="#<%=g_arColor(7)%>" vlink="#<%=g_arColor(7)%>" alink="#<%=g_arColor(6)%>">
<center>
<font face="<%=g_arFont(1)%>">

<% if Request.QueryString("retry") then %>
<font size=2>
No events matched your query.<br>Please try different parameters:<br>
</font>
<% end if %>

<!-- framing table -->
<table bgcolor="#<%=g_arColor(5)%>" cellspacing=0 cellpadding=2 border=0><tr><td>
<!-- end framing table -->

<table bgcolor="#<%=g_arColor(11)%>" cellspacing=0 cellpadding=2 border=0>
<form name="findform" action="webCal4_found.asp" method="post" onSubmit="return Validate();">
<tr>
	<td colspan=2 bgcolor="#<%=g_arColor(3)%>">
	<font face="<%=g_arFont(0)%>" size=4>
	<b>Find Events</b></font></td>
<tr>
	<td align="right"><font face="<%=g_arFont(2)%>" size=2>Title: </font></td>
	<td><input type="text" name="title" size="10"></td>
<tr>
	<td align="right"><font face="<%=g_arFont(2)%>" size=2>Description: </font></td>
	<td><input type="text" name="description" size="10"></td>
<tr>
	<td align="right"><font face="<%=g_arFont(2)%>" size=2>Between: </font></td>
	<td><input type="text" name="date_start" size="10"><input type="button" value="&gt;" onClick="calpopup(2);"></td>
<tr>
	<td align="right"><font face="<%=g_arFont(2)%>" size=2>and: </font></td>
	<td><input type="text" name="date_end" size="10"><input type="button" value="&gt;" onClick="calpopup(4);"></td>
<tr>
	<td align="right" colspan=2 bgcolor="#<%=g_arColor(12)%>"><input type="submit" value="Find"></td>
</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->

<input type="hidden" name="view" value="<%=Request.QueryString("view")%>">
</form>

<table cellspacing=4 cellpadding=2 border=0 width="50%">
<tr>
	<td colspan=2 align="center">
	<font face="<%=g_arFont(0)%>" color="#<%=g_arColor(5)%>"><b>Examples</b></font></td>
<tr>
	<td align="center" bgcolor="#<%=g_arColor(2)%>">
	<font face="<%=g_arFont(1)%>" size=2>to match</font></td>
	<td align="center" bgcolor="#<%=g_arColor(2)%>">
	<font face="<%=g_arFont(1)%>" size=2>use</font></td>
<tr>
	<td align="right">
	<font face="<%=g_arFont(1)%>" size=2>"dog" <u>or</u> "cat"</font></td>
	<td bgcolor="#<%=g_arColor(11)%>">
	<font face="<%=g_arFont(1)%>" size=2>dog cat</td>
<tr>
	<td align="right">
	<font face="<%=g_arFont(1)%>" size=2>both "dog" <u>and</u> "cat"</font></td>
	<td bgcolor="#<%=g_arColor(11)%>">
	<font face="<%=g_arFont(1)%>" size=2>dog+cat</td>
<tr>
	<td align="right">
	<font face="<%=g_arFont(1)%>" size=2>the <u>phrase</u> "dog cat"</font></td>
	<td bgcolor="#<%=g_arColor(11)%>">
	<font face="<%=g_arFont(1)%>" size=2>"dog cat"</td>
<tr>
	<td align="right">
	<font face="<%=g_arFont(1)%>" size=2><u>without</u> "dog"</font></td>
	<td bgcolor="#<%=g_arColor(11)%>">
	<font face="<%=g_arFont(1)%>" size=2>-dog</td>
<tr>
	<td align="right" valign="top">
	<font face="<%=g_arFont(1)%>" size=2>scheduled in 1997</font></td>
	<td bgcolor="#<%=g_arColor(11)%>" valign="top">
	<font face="<%=g_arFont(1)%>" size=2><%=DateSerial(1997,1,1)%><br>
	<%=DateSerial(1998,1,1)%></td>
<tr>
	<td align="right" valign="top">
	<font face="<%=g_arFont(1)%>" size=2>scheduled between<br>October 1998 and now</font></td>
	<td bgcolor="#<%=g_arColor(11)%>" valign="top">
	<font face="<%=g_arFont(1)%>" size=2><%=DateSerial(1998,10,1)%><br>[leave blank]*</td>
<tr>
	<td colspan=2><br><font face="<%=g_arFont(2)%>" size=1>
	*if you enter a value for one date and leave the other blank, the program will assume the current date for the blank field
	</font></td>
</table>
</center>
</font>
</body>
</html>