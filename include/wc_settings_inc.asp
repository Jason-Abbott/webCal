<%
' Copyright 1996-2004 Jason Abbott (webcal@webott.com)

dim g_arFont(2)		' array of font strings used
dim g_arColor		' array of colors
dim g_aBrowser		' browser

If Not IsArray(Session(g_sDB_NAME & "Groups")) Then Call goInitFirst()
If Not IsArray(Session(g_sDB_NAME & "Symbols")) Then Call initSymbols()

g_arFont(0) = "Tahoma, Arial, Helvetica"
g_arFont(1) = "Verdana, Arial, Helvetica"
g_arFont(2) = "Arial, Helvetica"
g_arColor = Application(g_sDB_NAME & "Colors")
Session.LCID = Session(g_sDB_NAME & "LCID")
g_aBrowser = Session(g_sDB_NAME & "Browser")

if Request.QueryString("logout") = 1 then Session.Abandon

'-------------------------------------------------------------------------
'	Name: 		goInitFirst
'	Purpose: 	initialize user session by redirecting to default page
'Modifications:
'	Date:		Name:	Description:
'	3/4/01		JEA		Creation
'-------------------------------------------------------------------------
Sub goInitFirst()
	dim sView
	dim sPage
	dim x
	sPage = Request.ServerVariables("SCRIPT_NAME")
	for each x in Array("month", "week", "day", "year")
		if InStr(sPage, x) Then sView = x : exit for
	next
	response.redirect "./?view=" & sView & "&date=" & Trim(Request.QueryString("date"))
End Sub

Sub initSymbols()
	dim oSession
	Set oSession = New wcSession : Call oSession.initSymbols() : Set oSession = Nothing
End Sub
%>