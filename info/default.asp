<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./data/webCal4_data.inc"-->
<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' Last updated 3/29/2000

dim strQuery		' query passed to db
dim strPage			' user's preferred start location
dim arTemp
dim g_arColor(14)  	' array of colors
dim arCache			' local copy of page cache
dim oConn			' connection object
dim oRS				' recordset object
dim intSize			' cache size
dim x, y			' loop counters

intSize = 9			' making this too big or small will degrade performance
Set oRS = Server.CreateObject("ADODB.RecordSet")
Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open g_strDSN

' get correct start page
if Request.QueryString("start") <> "" then
	strPage = Request.QueryString("start")
elseif Request.Cookies("start") <> "" then
	strPage = Request.Cookies("start")
else
	strPage = "month"
end if

' retrieve location id from cookie or database
'if Request.Cookies("lcid") <> "" then
if false then
	Session(unique & "LCID") = Request.Cookies("lcid")
else
	strQuery = "SELECT lcid FROM tblLCIDs WHERE lcid = " _
		& "(SELECT default_lcid FROM tblSettings)"
	oRS.Open strQuery, oConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	Session(unique & "LCID") = oRS("lcid")
	oRS.Close
end if

' DEBUG ONLY
Application(unique & "Cache") = ""
' DEBUG ONLY

if Not IsArray(Application(unique & "Cache")) then
	' initialize cache--create empty array of pages-----------------------
	ReDim arCache(intSize)
	for x = 0 to intSize
		' array holds query string, HTML page, expiration date
		arCache(x) = Array("","","")
	next
	
	' retrieve color settings from database-------------------------------
	strQuery = "SELECT * FROM tblColors WHERE (color_id) = " _
		& "(SELECT cal_color FROM tblSettings)"
	oRS.Open strQuery, oConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	arTemp = oRS.GetRows
	oRS.Close
	' skip first two fields (id and name)
	for x = 2 to UBound(arTemp)
		g_arColor(x - 2) = arTemp(x,0)
	next
	
	Application.Lock
	Application(unique & "Colors") = g_arColor
	Application(unique & "Cache") = arCache
	Application(unique & "Date") = Date
	Application.Unlock
elseif Application(unique & "Date") < Date then
	' cache exists but hasn't been updated today
	' clear expired pages from cache--------------------------------------
	arCache = Application(unique & "Cache")
	for x = 0 to intSize
		if arCache(x)(2) < Date then
			' the page has expired
			for y = x to intSize - 1
				' shift following pages up in the cache
				arCache(y) = arCache(y + 1)
			next
		end if
	next
	Application.Lock
	Application(unique & "Date") = Date
	Application(unique & "Cache") = arCache
	Application.Unlock
end if

Set oRS = nothing
oConn.Close
Set oConn = nothing
%>

<title>webCal 4.0</title>

<FRAMESET rows="*,0" border=0 framespacing=0 frameborder=0>
	<FRAME src="webCal4_<%=strPage%>.asp" name="body" scrolling="auto" marginwidth=8 marginheight=8 noresize>
	<FRAME src="./include/line.html" scrolling="no" marginwidth=0 marginheight=0 noresize>
</FRAMESET>