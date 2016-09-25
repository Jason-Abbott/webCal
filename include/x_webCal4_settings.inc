<%
' Copyright 1996-2004 Jason Abbott (webcal@webott.com)

dim g_arFont(2)		' array of font strings used
dim g_arColor		' array of colors
dim g_strBrowser	' browser

If Session(g_sDB_NAME & "LCID") = "" Then goInitFirst

g_arFont(0) = "Tahoma, Arial, Helvetica"
g_arFont(1) = "Verdana, Arial, Helvetica"
g_arFont(2) = "Arial, Helvetica"
g_arColor = Application(g_sDB_NAME & "Colors")
Session.LCID = Session(g_sDB_NAME & "LCID")
g_strBrowser = Session(g_sDB_NAME & "Browser")

if Request.QueryString("logout") = 1 then Session.Abandon

'-------------------------------------------------------------------------
'	Name: 		goInitFirst
'	Purpose: 	initialize user session by redirecting to default page
'Modifications:
'	Date:		Name:	Description:
'	3/4/01		JEA		Creation
'-------------------------------------------------------------------------
Sub goInitFirst()
	dim strView
	dim strPage
	dim strDate
	dim x
	
	strDate = Trim(Request.QueryString("date"))
	strPage = Request.ServerVariables("SCRIPT_NAME")
	for each x in Array("month", "week", "day", "year")
		if InStr(strPage, x) Then strView = x : exit for
	next
	response.redirect "./?view=" & strView & "&date=" & strDate
End Sub
%>