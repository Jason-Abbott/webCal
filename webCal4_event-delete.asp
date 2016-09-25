<% Option Explicit %>
<% Response.Buffer = True %>
<html>
<head>
<!--#include file="./data/webCal4_data.inc"-->
<!--#include file="./include/webCal4_verify.inc"-->
<!--#include file="./include/webCal4_constants.inc"-->
<!--#include file="./include/webCal4_settings.inc"-->
<script language="javascript" src="./script/webCal4_functions.js"></script>
<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' Last updated 2/23/2001

dim m_strSay
dim m_strQuery		' query string passed to database
dim m_oRS			' recordset object
dim m_strEventTitle

if Request.Form("fldCount") = 2 then
	m_strSay = "both"
else
	m_strSay = "all " & Request.Form("fldCount")
end if

' how many did they want to delete?
Select Case Request.Form("fldTimeScope")
	Case "one"
		m_strSay = "this (" & Request.Form("fldDate") & ") instance of"
	Case "future"
		m_strSay = "this (" & Request.Form("fldDate") & ") and <b>all future</b> instances of"
	Case "all"
		m_strSay = "<b>" & m_strSay & "</b> instances of"
	Case else
		m_strSay = ""
End Select

m_strQuery = "SELECT event_title FROM tblEvents WHERE " _
	& "(event_id)=" & Request.Form("fldEventID")

Set m_oRS = Server.CreateObject("ADODB.Recordset")
m_oRS.Open m_strQuery, g_strDSN, adOpenForwardOnly, adLockReadOnly, adCmdText
m_strEventTitle = m_oRS("event_title")
m_oRS.Close : Set m_oRS = nothing
%>

</head>
<body onLoad="showMessage();" bgcolor="#<%=g_arColor(1)%>" link="#<%=g_arColor(7)%>" vlink="#<%=g_arColor(7)%>" alink="#<%=g_arColor(6)%>">
<center>

<!-- framing table -->
<table bgcolor="#<%=g_arColor(6)%>" border=0 cellpadding=2 cellspacing=0><tr><td>
<!-- end framing table -->

<table bgcolor="#<%=g_arColor(11)%>" border=0 cellpadding=4 cellspacing=0>
<form name="frmDel" method="post">
<tr bgcolor="#<%=g_arColor(3)%>" valign="bottom">
	<td colspan=3><font face="<%=g_arFont(0)%>" size=4>
	<b>Event Deletion</b></font></td>
<tr>
	<td align="center" colspan=3><font face="<%=g_arFont(2)%>">
	Are you sure you want to erase <%=m_strSay%> <i><%=m_strEventTitle%></i>?</font></td>
<tr>
	<td colspan=3 align="center" bgcolor="#<%=g_arColor(12)%>">
		<input type="button" value="Yes" onClick='goPage("webCal4_event-deleted.asp","frmDel");'>
		<input type="button" value="No" onClick='goPage("<%=Request.Form("fldURL")%>","frmDel");'>
	</td>
<tr>
	<td align="center" colspan=3><font face="<%=g_arFont(0)%>" size=2>
	<b><font color="#cc0000">Caution</font>: erased events cannot be restored</b></font></td>
</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->

<input type="hidden" name="fldEventID" value="<%=Request.Form("fldEventID")%>">
<input type="hidden" name="fldDate" value="<%=Request.Form("fldDate")%>">
<input type="hidden" name="fldTimeScope" value="<%=Request.Form("fldTimeScope")%>">
<input type="hidden" name="fldURL" value="<%=Request.Form("fldURL")%>">
<input type="hidden" name="fldView" value="<%=Request.Form("fldView")%>">
</form>

</body>
</html>

