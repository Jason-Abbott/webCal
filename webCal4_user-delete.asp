<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./data/webCal4_data.inc"-->
<!--#include file="./include/webCal4_verify.inc"-->
<!--#include file="./include/webCal4_constants.inc"-->
<!--#include file="./include/webCal4_settings.inc"-->
<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' Last updated 2/25/2001

dim m_strUserName	' user first name
dim m_strLastName	' user last name
dim m_intUserID		' user id
dim m_strQuery		' query string passed to db
dim m_strUserList	' selection list of other users
dim m_intEventCount	' events created by this user
dim m_oConn			' connection object
dim m_oRS			' recordset object

m_intUserID = Request.Form("fldUserID")
Set m_oConn = Server.CreateObject("ADODB.Connection") : m_oConn.Open g_strDSN
Set m_oRS = Server.CreateObject("ADODB.RecordSet")

' get info on the user to be deleted
m_strQuery = "SELECT name_first, name_last FROM tblUsers WHERE " _
	& "(user_id)=" & m_intUserID
m_oRS.Open m_strQuery, m_oConn, adOpenForwardOnly, adLockReadOnly, adCmdText
m_strUserName = m_oRS("name_first") & " " & m_oRS("name_last")
m_oRS.Close

' get a count of the events scheduled by this user
m_strQuery = "SELECT COUNT(user_id) AS event_count " _
	& "FROM tblEvents WHERE user_id=" & m_intUserID
m_oRS.Open m_strQuery, m_oConn, adOpenForwardOnly, adLockReadOnly, adCmdText
m_intEventCount = m_oRS("event_count")
m_oRS.Close

if m_intEventCount > 0 then
	' get a list of all other users
	m_strQuery = "SELECT user_id, name_first, name_last FROM tblUsers" _
		& " WHERE (user_id)<>" & m_intUserID _
		& " ORDER BY name_last, name_first"
	m_strUserList = makeList(m_strQuery, "", m_oConn)
end if

Set m_oRS = nothing
m_oConn.Close : Set m_oConn = nothing
%>
<html>
<head>
<script language="javascript" src="./include/webCal4_functions.js"></SCRIPT>
</head>
<body bgcolor="#<%=g_arColor(1)%>" link="#<%=g_arColor(7)%>" vlink="#<%=g_arColor(7)%>" alink="#<%=g_arColor(6)%>">
<center>

<!-- framing table -->
<table bgcolor="#<%=g_arColor(6)%>" border=0 cellpadding=2 cellspacing=0><tr><td>
<!-- end framing table -->

<table bgcolor="#<%=g_arColor(11)%>" border=0 cellpadding=3 cellspacing=0>
<form name="frmDel" method="post">
<tr>
	<td bgcolor="#<%=g_arColor(4)%>" colspan=3>
	<font face="<%=g_arFont(0)%>" size=4>
	<b>User Deletion</b></font></td>
<tr>
	<td colspan=3><font face="<%=g_arFont(0)%>" size=2>
	
<% if m_intEventCount > 0 then %>
	
	<b>What should happen to the <%=m_intEventCount%> events scheduled by <%=m_strUserName%>?</b><br>
	<input type="radio" name="fldAction" value="delete">erase them all<br>
	<input type="radio" name="fldAction" value="some" checked>erase the private but transfer the public events<br>
	<input type="radio" name="fldAction" value="move">transfer them all<br>
	<center>transfer to <select name="fldRecipient"><%=m_strUserList%></select></center>
	</font></td>
	
<% else %>

Are you sure you want to erase the user <u><%=m_strUserName%></u>?

<% end if %>
	
<tr>
	<td colspan=3 align="center" bgcolor="#<%=g_arColor(12)%>">
		<input type="button" value="Continue" onClick='goPage("webCal4_user-deleted.asp","frmDel");'>
		<input type="button" value="Cancel" onClick='goPage("webCal4_admin.asp","frmDel");'>
	</td>
<tr>
	<td align="center" colspan=3><font face="<%=g_arFont(0)%>" size=2>
	<b><font color="#bb0000">Caution</font>: erased users and events cannot be restored</b></font></td>
</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->

<input type="hidden" name="fldUserID" value="<%=m_intUserID%>">
<input type="hidden" name="fldEventCount" value="<%=m_intEventCount%>">
<input type="hidden" name="fldView" value="<%=Request.Form("fldView")%>">
</form>

</body>
</html>