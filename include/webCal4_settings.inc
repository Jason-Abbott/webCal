<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' Last updated 3/4/2001

dim g_arFont(2)		' array of font strings used
dim g_arColor		' array of colors
dim g_strBrowser	' browser

If Session(g_unique & "LCID") = "" Then goInitFirst
If Not IsArray(Session(g_unique & "Symbols")) Then initSymbols

g_arFont(0) = "Tahoma, Arial, Helvetica"
g_arFont(1) = "Verdana, Arial, Helvetica"
g_arFont(2) = "Arial, Helvetica"
g_arColor = Application(g_unique & "Colors")
Session.LCID = Session(g_unique & "LCID")
g_strBrowser = Session(g_unique & "Browser")

if Request.QueryString("logout") = 1 then Session.Abandon

' build an array of font maps for this client (updated 3/4/01)
' updates session variable -----------------------------------------------
Sub initSymbols()
	dim strSupport
	dim arSupport		' fonts supported by client
	dim arSymbol(1,26)	' array of character and font face
	dim arChar			' temporary array of character codes
	dim x
	
	Const FONT_WEBDINGS = 0
	Const FONT_WINGDINGS = 1
	Const FONT_ZAPFDINGS = 2
	
	strSupport = Trim(Request.QueryString("fs"))
	if strSupport = "" Then goInitFirst
	arSupport = Split(strSupport, ",")
	If arSupport(FONT_WEBDINGS) And arSupport(FONT_WINGDINGS) Then
		arChar = Array(39,51,52,53,54,76,105,113,120,121,115,183,184, _
			185,186,187,188,189,190,191,192,194,207,208,235,253,254)
		for x = 0 to UBound(arChar)
			arSymbol(g_FONT_CHAR, x) = arChar(x)
			arSymbol(g_FONT_FACE, x) = "Webdings"
		next
		for x = g_CHAR_CLOCK_1 to g_CHAR_CLOCK_12
			arSymbol(g_FONT_FACE, x) = "Wingdings"
		next
		for each x in Array(g_CHAR_CLOSE, g_CHAR_OPEN, g_CHAR_XBOX, g_CHAR_CHECKBOX)
			arSymbol(g_FONT_FACE, x) = "Wingdings"
		next		
	ElseIf arSupport(FONT_WINGDINGS) Then
		' probably Netscape (32 is space)
		arChar = Array(32,231,232,233,234,32,105,32,120,121,63,183,184, _
			185,186,187,188,189,190,191,192,194,32,32,32,253,254)
		for x = 0 to UBound(arChar)
			arSymbol(g_FONT_CHAR, x) = arChar(x)
			arSymbol(g_FONT_FACE, x) = "Wingdings"
		next
		for each x in Array(g_CHAR_INFO, g_CHAR_QUESTION)
			arSymbol(g_FONT_FACE, x) = "Arial"
		next
	ElseIf arSupport(FONT_ZAPFDINGS) Then
		' must be Macintosh
	
	End If
	Erase arChar : Session(g_unique & "Symbols") = arSymbol
End Sub

' user needs to hit default page first to init settings (updated 3/4/01)
' redirects to new page --------------------------------------------------
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
	response.redirect "default.asp?view=" & strView & "&date=" & strDate
End Sub
%>