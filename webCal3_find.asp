<!--
Copyright 1999 Jason Abbott (jason@webott.com)
Last updated 06/04/1999
-->

<html>
<head>
<SCRIPT TYPE="text/javascript" LANGUAGE="javascript" SRC="webCal3_popup.js"></SCRIPT>
<SCRIPT LANGUAGE="javascript">
<!--

function Validate() {
	if (document.findform.title.value.length <= 0
    && document.findform.description.value.length <= 0
	 && document.findform.date_start.value.length <= 0
	 && document.findform.date_end.value.length <= 0) {
		alert("You must enter criteria in at least one field");
		document.findform.title.select();
		document.findform.title.focus();
		return false;
	}
}
//-->
</SCRIPT>

<!--#include file="webCal3_themes.inc"-->
</head>
<body onload="init();" bgcolor="#<%=color(1)%>" link="#<%=color(7)%>" vlink="#<%=color(7)%>" alink="#<%=color(6)%>">
<center>
<font face="Verdana, Arial, Helvetica">
<% if Request.QueryString("retry") then %>
<font size=2>
No events matched your query.<br>Please try different parameters:<br>
</font>
<% end if %>

<!-- framing table -->
<table bgcolor="#<%=color(5)%>" cellspacing=0 cellpadding=2 border=0><tr><td>
<!-- end framing table -->

<table bgcolor="#<%=color(11)%>" cellspacing=0 cellpadding=2 border=0>
<form name="findform" action="webCal3_found.asp" method="post" onSubmit="return Validate();">
<tr>
	<td colspan=2 bgcolor="#<%=color(3)%>">
	<font face="Tahoma, Arial, Helvetica" size=4>
	<b>Find Events</b></font></td>
<tr>
	<td align="right"><font face="Arial, Helvetica" size=2>Title: </font></td>
	<td><input type="text" name="title" size="10"></td>
<tr>
	<td align="right"><font face="Arial, Helvetica" size=2>Description: </font></td>
	<td><input type="text" name="description" size="10"></td>
<tr>
	<td align="right"><font face="Arial, Helvetica" size=2>Between: </font></td>
	<td><input type="text" name="date_start" size="10"><input type="button" value="&gt;" onClick="calpopup('findform','date_start');"></td>
<tr>
	<td align="right"><font face="Arial, Helvetica" size=2>and: </font></td>
	<td><input type="text" name="date_end" size="10"><input type="button" value="&gt;" onClick="calpopup('findform','date_end');"></td>
<tr>
	<td align="right" colspan=2 bgcolor="#<%=color(12)%>"><input type="submit" value="Find"></td>
</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->

<input type="hidden" name="view" value="<%=Request.QueryString("view")%>">
</form>

<table cellspacing=4 cellpadding=2 border=0>
<tr>
	<td colspan=2 align="center">
	<font face="Tahoma, Arial, Helvetica" color="#<%=color(5)%>"><b>Examples</b></font></td>
<tr>
	<td align="center" bgcolor="#<%=color(2)%>">
	<font face="Verdana, Arial, Helvetica" size=2>to match</font></td>
	<td align="center" bgcolor="#<%=color(2)%>">
	<font face="Verdana, Arial, Helvetica" size=2>use</font></td>
<tr>
	<td align="right">
	<font face="Verdana, Arial, Helvetica" size=2>"dog" <u>or</u> "cat"</font></td>
	<td bgcolor="#<%=color(11)%>">
	<font face="Verdana, Arial, Helvetica" size=2>dog cat</td>
<tr>
	<td align="right">
	<font face="Verdana, Arial, Helvetica" size=2>both "dog" <u>and</u> "cat"</font></td>
	<td bgcolor="#<%=color(11)%>">
	<font face="Verdana, Arial, Helvetica" size=2>dog+cat</td>
<tr>
	<td align="right">
	<font face="Verdana, Arial, Helvetica" size=2>the <u>phrase</u> "dog cat"</font></td>
	<td bgcolor="#<%=color(11)%>">
	<font face="Verdana, Arial, Helvetica" size=2>"dog cat"</td>
<tr>
	<td align="right">
	<font face="Verdana, Arial, Helvetica" size=2><u>without</u> "dog"</font></td>
	<td bgcolor="#<%=color(11)%>">
	<font face="Verdana, Arial, Helvetica" size=2>-dog</td>
<tr>
	<td align="right" valign="top">
	<font face="Verdana, Arial, Helvetica" size=2>scheduled in 1997</font></td>
	<td bgcolor="#<%=color(11)%>" valign="top">
	<font face="Verdana, Arial, Helvetica" size=2>1/1/97<br>1/1/98</td>
<tr>
	<td align="right" valign="top">
	<font face="Verdana, Arial, Helvetica" size=2>scheduled between<br>October 1998 and now</font></td>
	<td bgcolor="#<%=color(11)%>" valign="top">
	<font face="Verdana, Arial, Helvetica" size=2>10/1/98<br>[leave blank]*</td>
</table>
</center>
<p>
<font size=1>
*if you enter a value for one date and leave the other blank, the program will assume the current date for the blank field
</font></font>

</body>
</html>