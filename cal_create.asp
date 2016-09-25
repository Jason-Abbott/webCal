<!-- 
Copyright 1999 Jason Abbott (jabbott@uidaho.edu)
Last updated 6/3/98
-->

<!--#include virtual="/header_start.inc"-->
Create a New Calendar
<!--#include virtual="/header_end.inc"-->

<center>

<table cellpadding=4 cellspacing=0 border=0>
<form action="cal_created.asp" method="post">
<tr>
	<td align=right valign=top <%=light%>>
	<font face="arial" size=2>Title</font></td>
	<td><input type="text" name="name"><br>
	Try to keep your name short so that<br>it displays well in the menus.
	</td>
<tr>
	<td align=right valign=top <%=light%>>
	<font face="arial" size=2>Restricted</font></td>
	<td><input type="checkbox" name="private"><br>
	Checking this box will prevent users<br>outside of the Boise Center from<br>seeing your calendar.
	</td>
<tr>
	<td colspan=2 align=center>
	<input type="submit" value="Add"></td>
</table>
</form>
</center>

<!--#include virtual="/footer.inc"-->