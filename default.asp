<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./include/wc_constants_inc.asp"-->
<!--#include file="./include/wc_functions_inc.asp"-->
<!--#include file="./include/wc_data_cls.asp"-->
<!--#include file="./include/wc_session_cls.asp"-->
<!--#include file="./include/wc_cache_cls.asp"-->
<%
' Copyright 1996-2004 Jason Abbott (webcal@webott.com)

dim m_oSession
dim m_dtDate
dim m_sView

m_dtDate = Request.QueryString("date")
m_sView = Request.QueryString("view")

Set m_oSession = New wcSession
Call m_oSession.Validate(m_sView, m_dtDate)
Set m_oSession = Nothing


Sub WriteFonts()
	dim x
	with response
		.write "<table border='0'>"
		.write "<tr><th></th><th>Webdings</th><th>Wingdings</th><th>Marlett</th><th>Tahoma</th>"
		for x = 32 to 255
			' if x mod 3 then .write "</tr><tr>"
			.write "<tr><td>"
			.write x
			.write "</td><td style='font-family: Webdings; font-size: 10pt;'>"
			.write Chr(x)
			.write "</font></td><td style='font-family: Wingdings; font-size: 10pt;'>"
			.write Chr(x)
			.write "</font></td><td style='font-family: Marlett; font-size: 10pt;'>"
			.write Chr(x)
			.write "</font></td><td style='font-family: Tahoma; font-size: 10pt;'>"
			.write Chr(x)
			.write "</font></td>"
		next
		.write "</table>"
		.end
	end with
End Sub
%>
<html><head>
<script language="javascript" src="./script/wc_initialize.js"></script>
</head><body onLoad="initCalendar('<%=m_sView%>','<%=m_dtDate%>')";></body></html>

<% 'Call WriteFonts() %>