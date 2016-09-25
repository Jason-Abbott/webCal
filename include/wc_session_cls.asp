<%
' Copyright 1996-2004 Jason Abbott (webcal@webott.com)

Class wcSession

	'-------------------------------------------------------------------------
	'	Name: 		Validate()
	'	Purpose: 	validate session
	'Modifications:
	'	Date:		Name:	Description:
	'	9/23/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub Validate(ByRef r_sView, ByVal v_sDate)
		dim bNewLoad
		dim oCache
	
		If true or Not IsArray(Application(g_sDB_NAME & "Cache")) then
			' webCal has not been run on this server
			Call initApp(r_sView) : bNewLoad = true
		else
			' set LCID so date functions are consistent
			Session.LCID = Application(g_sDB_NAME & "Settings")(g_LCID)
		end if
			
		if Not bNewLoad And Application(g_sDB_NAME & "Date") < Date then
			' cache exists but hasn't been updated today
			' clear expired pages from cache
			Set oCache = New wcCache : Call oCache.Refresh() : Set oCache = Nothing
		end if
		
		if IsArray(Session(g_sDB_NAME & "Groups")) then
			if IsVoid(r_sView) then r_sView = Session(g_sDB_NAME & "StartPage")
		else
			' create new user session
			r_sView = initGuestUser()
		end if
	End Sub

	'-------------------------------------------------------------------------
	'	Name: 		initApp()
	'	Purpose: 	retrieve and store settings in application object
	'Modifications:
	'	Date:		Name:	Description:
	'	3/4/01		JEA		Creation
	'-------------------------------------------------------------------------
	Private Sub initApp(ByRef r_sView)
		dim aTemp			' temporary color array
		dim aCache()		' local copy of page cache
		dim aColor(14)		' local copy of color array
		dim aSettings
		dim lColorID		' color id
		dim oData
		dim x				' loop counter
	
		' retrieve settings
		Set oData = New wcData
		aSettings = getAppSettings()
		aTemp = getColorPalette(aSettings(g_COLOR_ID))
		Set oData = Nothing
		
		aSettings(g_USE_CACHE) = CBool(aSettings(g_USE_CACHE) <> 0)
		aSettings(g_EASY_EDIT) = CBool(aSettings(g_EASY_EDIT) <> 0)
		aSettings(g_SHOW_WEEKEND) = CBool(aSettings(g_SHOW_WEEKEND) <> 0)
		Session.LCID = aSettings(g_LCID)
		if Not IsNumber(r_sView) then r_sView = aSettings(g_START_PAGE)
		
		' skip first two fields (id and name)
		for x = 2 to UBound(aTemp)
			aColor(x - 2) = aTemp(x,0)
		next
		Erase aTemp
		
		if aSettings(g_USE_CACHE) then
			' create empty array of pages
			ReDim aCache(aSettings(g_CACHE_SIZE))
			for x = 0 to aSettings(g_CACHE_SIZE)
				' array holds query string, HTML page, expiration date, grid settings
				aCache(x) = Array("","","","")
			next
		end if
		
		Application.Lock
		Application(g_sDB_NAME & "Colors") = aColor
		Application(g_sDB_NAME & "Settings") = aSettings
		Application(g_sDB_NAME & "Cache") = aCache
		Application(g_sDB_NAME & "Date") = Date
		Application.Unlock
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		initGuestUser()
	'	Purpose: 	creates user session variables for guests
	'Modifications:
	'	Date:		Name:	Description:
	'	1/7/01		JEA		Creation
	'-------------------------------------------------------------------------
	Private Function initGuestUser()
		dim aSettings		' application settings
		dim aSegments		' array of segment settings
	
		aSettings = Application(g_sDB_NAME & "Settings")
		aSegments = Array(aSettings(g_DEFAULT_SEG_MINS), aSettings(g_DEFAULT_SEG_START), aSettings(g_DEFAULT_SEG_END))

		Call initUser(0, aSettings(g_SHOW_WEEKEND), Array(aSegments, aSegments, Array(0,0,0)), _
			aSettings(g_START_PAGE), aSettings(g_LCID))

		Session(g_sDB_NAME & "Browser") = getClientInfo()
		initGuestUser = aSettings(g_START_PAGE)
		
		' retrieve location id from cookie or database
		'if Request.Cookies("lcid") <> "" then
		'	Session(g_sDB_NAME & "LCID") = Request.Cookies("lcid")
		'else
		'	Session(g_sDB_NAME & "LCID") = arSettings(g_LCID)
		'end if
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		initUser()
	'	Purpose: 	creates user session variables
	'Modifications:
	'	Date:		Name:	Description:
	'	9/24/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub initUser(ByVal v_sUserID, ByVal v_bShowWeekend, ByVal v_aSegments, _
		ByVal v_sStartPage, ByVal v_sLCID)
		
		Session(g_sDB_NAME & "UserID") = v_sUserID
		Session(g_sDB_NAME & "Weekends") = v_bShowWeekend
		Session(g_sDB_NAME & "Segments") = v_aSegments
		Session(g_sDB_NAME & "StartPage") = v_sStartPage
		Session(g_sDB_NAME & "LCID") = v_sLCID
		If Not IsArray(Session(g_sDB_NAME & "Browser")) then Session(g_sDB_NAME & "Browser") = getClientInfo()
		Call initDataAccess()
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		initSymbols
	'	Purpose: 	build an array of font maps for this client
	' Modifications:
	'	Date:		Name:	Description:
	'	3/4/01		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub initSymbols()
		dim sSupport
		dim aSupport		' fonts supported by client
		dim aSymbol(1,26)	' array of character and font face
		dim aChar			' temporary array of character codes
		dim x
		
		Const FONT_WEBDINGS = 0
		Const FONT_WINGDINGS = 1
		Const FONT_ZAPFDINGS = 2
		
		sSupport = Trim(Request.QueryString("fs"))
		aSupport = Split(sSupport, ",")
		If aSupport(FONT_WEBDINGS) And aSupport(FONT_WINGDINGS) Then
			aChar = Array(39,51,52,53,54,76,105,113,120,121,115,183,184, _
				185,186,187,188,189,190,191,192,194,207,208,235,253,254)
			for x = 0 to UBound(aChar)
				aSymbol(g_FONT_CHAR, x) = aChar(x)
				aSymbol(g_FONT_FACE, x) = "Webdings"
			next
			for x = g_CHAR_CLOCK_1 to g_CHAR_CLOCK_12
				aSymbol(g_FONT_FACE, x) = "Wingdings"
			next
			for each x in Array(g_CHAR_CLOSE, g_CHAR_OPEN, g_CHAR_XBOX, g_CHAR_CHECKBOX)
				aSymbol(g_FONT_FACE, x) = "Wingdings"
			next		
		ElseIf aSupport(FONT_WINGDINGS) Then
			' probably Netscape (32 is space)
			aChar = Array(32,231,232,233,234,32,105,32,120,121,63,183,184, _
				185,186,187,188,189,190,191,192,194,32,32,32,253,254)
			for x = 0 to UBound(aChar)
				aSymbol(g_FONT_CHAR, x) = aChar(x)
				aSymbol(g_FONT_FACE, x) = "Wingdings"
			next
			for each x in Array(g_CHAR_INFO, g_CHAR_QUESTION)
				aSymbol(g_FONT_FACE, x) = "Arial"
			next
		ElseIf aSupport(FONT_ZAPFDINGS) Then
			' must be Macintosh
		
		End If
		'Erase aChar
		Session(g_sDB_NAME & "Symbols") = aSymbol
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		initDataAccess()
	'	Purpose: 	get the list of groups and scopes available to this user
	'Modifications:
	'	Date:		Name:	Description:
	'	2/27/01		JEA		Creation
	'-------------------------------------------------------------------------
	Private Sub initDataAccess()
		dim lUserID
		dim aScopes				' array of scope settings
		dim aGroups				' array of group settings
		dim lAccessLevel		' default group access
		dim x					' loop counter
		
		lUserID = MakeNumber(Session(g_sDB_NAME & "UserID"))
		aGroups = getUserGroups(lUserID, lAccessLevel)
		aScopes = getUserScopes(lUserID)
		
		if lUserID <> 0 then
			' logged in user
			
			lAccessLevel = getDefaultAccess(lUserID)
			if lAccessLevel > g_NO_ACCESS then
				' assign default to groups with no explicit access
				if lAccessLevel = g_ADMIN_ACCESS then
					' admin access overrides all others
					for x = 0 to UBound(aGroups, 2)
						aGroups(g_GROUP_ACCESS, x) = g_ADMIN_ACCESS
					next
				else
					for x = 0 to UBound(aGroups, 2)
						if aGroups(g_GROUP_ACCESS, x) = "" Or aGroups(g_GROUP_ACCESS, x) = -1 then
							aGroups(g_GROUP_ACCESS, x) = lAccessLevel
						end if
					next
				end if
			end if
		else
			' only "public" scope is visible to guests
			aScopes(g_VISIBLE, 0) = 1
		end if
		Session(g_sDB_NAME & "Scopes") = aScopes
		Session(g_sDB_NAME & "Groups") = aGroups
		Session(g_sDB_NAME & "Query") = getUserQuery(aGroups, aScopes)
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		getUserQuery()
	'	Purpose: 	pre-build the user-specific portion of event query
	'Modifications:
	'	Date:		Name:	Description:
	'	3/3/01		JEA		Creation
	'-------------------------------------------------------------------------
	Private Function getUserQuery(ByVal v_aGroups, ByVal v_aScopes)
		dim sGroups				' query condition for groups
		dim sScopes				' query condition for scopes
		dim sQuery
	
		sGroups = makeWHERE(v_aGroups, "group_id", g_GROUP_ID)
		sScopes = makeWHERE(v_aScopes, "scope_id", g_SCOPE_ID)
		sQuery = ""
		if sGroups <> g_sNO_EVENTS And sScopes <> g_sNO_EVENTS then
			if sGroups <> "" And sScopes <> "" then sQuery = " AND "
			sQuery = Trim(sGroups & sQuery & sScopes)
			getUserQuery = IIf(IsVoid(sQuery), "", " AND " & sQuery)
		else
			getUserQuery = g_sNO_EVENTS
		end if
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		makeWHERE()
	'	Purpose: 	generate conditions of all queries for this user
	'	Return: 	string
	'Modifications:
	'	Date:		Name:	Description:
	'	3/3/01		JEA		Creation
	'-------------------------------------------------------------------------
	Private Function makeWHERE(ByVal v_aParms, ByVal v_sField, ByVal v_lID)
		dim sWhere			' SQL WHERE condition
		dim lAllowCount		' one-based count of matching parms
		dim lParmCount		' zero-based count of parms
		dim x				' loop counter
		
		lAllowCount = 0
		lParmCount = UBound(v_aParms, 2)
		for x = 0 to lParmCount
			if v_aParms(g_VISIBLE, x) then
				' add scopes and groups that are visible to this user
				sWhere = sWhere & v_aParms(v_lID, x) & ","
				lAllowCount = lAllowCount + 1
			end if
		next
		' trim exra comma
		if sWhere <> "" then sWhere = Left(sWhere, Len(sWhere) - 1)
		
		if lAllowCount = 1 then
			' match single parameter
			sWhere = "egs." & v_sField & "=" & sWhere
		elseif lAllowCount <> (lParmCount + 1) And lAllowCount > 0 then
			sWhere = "egs." & v_sField & " IN (" & sWhere & ")"
		elseif lAllowCount > 0 then
			' match all parameters, so no limitation
			sWhere = ""
		else
			' no parameters match
			sWhere = g_sNO_EVENTS
		end if
		makeWHERE = sWhere
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		getViewQuery()
	'	Purpose: 	Build the query string to retrieve events for view
	'	Return: 	string
	'Modifications:
	'	Date:		Name:	Description:
	'	1/7/01		JEA		Creation
	'	10/16/03	JEA		Sort first by event date
	'-------------------------------------------------------------------------
	Public Function getViewQuery(ByVal v_dtFirst, ByVal v_dtLast)
		if Session(g_sDB_NAME & "Query") <> g_sNO_EVENTS then
			getViewQuery = "SELECT e.event_id, e.event_title, e.event_recur, " _
				& "e.event_color, e.time_start, e.time_end, ed.event_date, " _
				& "e.event_description, e.skip_weekends " _
				& "FROM ((tblEvents AS e INNER JOIN tblEventDates AS ed " _
				& "ON e.event_id = ed.event_id) INNER JOIN tblEventGroupScopes AS egs " _
				& "ON e.event_id = egs.event_id) WHERE (ed.event_date " _
				& "BETWEEN " & g_sDB_DELIM _
				& sqlDate(v_dtFirst) _
				& " 12:00:00 AM" & g_sDB_DELIM & " AND " & g_sDB_DELIM _
				& sqlDate(v_dtLast)  _
				& " 11:59:59 PM" & g_sDB_DELIM & ")" & Session(g_sDB_NAME & "Query") _
				& " ORDER BY ed.event_date, e.time_start"
		else
			getViewQuery = g_sNO_EVENTS
		end if
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		getClientInfo()
	'	Purpose: 	set session values with browser info
	'	Return:		boolean
	'Modifications:
	'	Date:		Name:	Description:
	'				JEA		Creation
	'	8/28/01		JEA		Encapsulated in method
	'	5/29/02		JEA		Build array with more details
	'	1/14/03		JEA		add support for Safari browser
	'-------------------------------------------------------------------------
	Private Function getClientInfo()
		dim aClient(3)
		dim sBrowser
		dim oRegExp
		dim oMatches
		dim oMatch
		dim x
		
		sBrowser = LCase(Request.ServerVariables("HTTP_USER_AGENT"))
		
		Set oRegExp = New RegExp
		oRegExp.IgnoreCase = true
		oRegExp.Global = true
		
		' get browser info
		select case true
			case IsMatch(sBrowser, "WebTV", oRegExp)
				aClient(g_BROWSER_ID) = g_BROWSER_WEBTV
				aClient(g_BROWSER_VERSION) = getVersion(sBrowser, "WebTV/", "\d+\.\d+", oRegExp)
			case IsMatch(sBrowser, "msie", oRegExp)
				aClient(g_BROWSER_ID) = g_BROWSER_IE
				aClient(g_BROWSER_VERSION) = getVersion(sBrowser, "MSIE ", "\d+\.\d+", oRegExp)
			case IsMatch(sBrowser, "Netscape6", oRegExp)
				aClient(g_BROWSER_ID) = g_BROSWER_NS
				aClient(g_BROWSER_VERSION) = getVersion(sBrowser, "Netscape6/", "\d+\.\d", oRegExp)
			case IsMatch(sBrowser, "Safari", oRegExp)
				aClient(g_BROWSER_ID) = g_BROWSER_SAFARI
				aClient(g_BROWSER_VERSION) = getVersion(sBrowser, "Safari/", "\d+", oRegExp)
			case IsMatch(sBrowser, "Galeon", oRegExp)
				aClient(g_BROWSER_ID) = g_BROWSER_MOZILLA
				aClient(g_BROWSER_VERSION) = getVersion(sBrowser, "Galeon/", "\d+\.\d+", oRegExp)
			case IsMatch(sBrowser, "Gecko", oRegExp)
				aClient(g_BROWSER_ID) = g_BROWSER_MOZILLA
				aClient(g_BROWSER_VERSION) = getVersion(sBrowser, "rv:", "\d+\.\d+", oRegExp)
			case IsMatch(sBrowser, "Opera", oRegExp)
				aClient(g_BROWSER_ID) = g_BROWSER_OPERA
				aClient(g_BROWSER_VERSION) = getVersion(sBrowser, "Opera ", "\d+\.\d+", oRegExp)
			case else
				aClient(g_BROWSER_ID) = g_BROSWER_NS
				aClient(g_BROWSER_VERSION) = getVersion(sBrowser, "Mozilla/", "\d+\.\d+", oRegExp)
		end select
		
		' get operating system info
		select case true
			case IsMatch(sBrowser, "Windows NT", oRegExp)
				aClient(g_OS_ID) = g_OS_WINNT
				aClient(g_OS_VERSION) = getVersion(sBrowser, "Windows NT ", "\d+\.\d+", oRegExp)
			case IsMatch(sBrowser, "Windows 2000", oRegExp)
				aClient(g_OS_ID) = g_OS_WINNT
				aClient(g_OS_VERSION) = 5.0
			case IsMatch(sBrowser, "Windows", oRegExp)
				aClient(g_OS_ID) = g_OS_WIN
				aClient(g_OS_VERSION) = getVersion(sBrowser, "Windows ", "\d+\.*\d*", oRegExp)
			case IsMatch(sBrowser, "Win", oRegExp)
				aClient(g_OS_ID) = g_OS_WIN
				aClient(g_OS_VERSION) = getVersion(sBrowser, "Win", "\d+\.*\d*", oRegExp)
			case IsMatch(sBrowser, "Mac", oRegExp)
				aClient(g_OS_ID) = g_OS_MAC
				aClient(g_OS_VERSION) = 0
			case IsMatch(sBrowser, "IRIX", oRegExp), IsMatch(sBrowser, "Linux", oRegExp)
				aClient(g_OS_ID) = g_OS_UNIX
				aClient(g_OS_VERSION) = 0
			case else
				aClient(g_OS_ID) = g_OS_UNKNOWN
				aClient(g_OS_VERSION) = 0
		end select
	
		Set oRegExp = nothing
		getClientInfo = aClient
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		getVersion()
	'	Purpose: 	run regexp match
	'	Return:		float
	'Modifications:
	'	Date:		Name:	Description:
	'	5/29/02		JEA		Creation
	'-------------------------------------------------------------------------
	Private Function getVersion(ByVal v_sString, ByVal v_sExclude, ByVal v_sPattern, ByRef r_oRegExp)
		dim bNewObject
		dim oMatches
		
		if Not IsObject(r_oRegExp) then
			Set r_oRegExp = New RegExp
			r_oRegExp.IgnoreCase = true
			r_oRegExp.Global = true
			bNewObject = true
		else
			bNewObject = false
		end if
		v_sString = LCase(v_sString) : v_sExclude = LCase(v_sExclude)
		r_oRegExp.Pattern = v_sExclude & v_sPattern
		Set oMatches = r_oRegExp.Execute(v_sString)
		if oMatches.Count > 0 then getVersion = Replace(oMatches.Item(0).Value, v_sExclude, "")
		Set oMatches = nothing
		if bNewObject then set r_oRegExp = nothing
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		getAppSettings()
	'	Purpose: 	get app settings
	'	Return: 	array
	' Modifications:
	'	Date:		Name:	Description:
	'	9/24/03		JEA		Created
	'	10/8/03		JEA		Reduce to single dimension
	'-------------------------------------------------------------------------
	Private Function getAppSettings()
		dim oData
		dim sQuery
		sQuery = "SELECT cal_color, default_lcid, password_length, " _
			& "password_life, cache_pages, use_cache, easy_edit, " _
			& "show_weekend, seg_mins, seg_start, seg_end, start_page FROM tblSettings"
		Set oData = New wcData
		getAppSettings = oData.dimDown(oData.GetArray(sQuery), 0)
		Set oData = nothing
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		getColorPalette()
	'	Purpose: 	get colors for theme
	'	Return: 	array
	' Modifications:
	'	Date:		Name:	Description:
	'	9/24/03		JEA		Created
	'-------------------------------------------------------------------------
	Private Function getColorPalette(ByVal v_sTheme)
		dim oData
		dim sQuery
		sQuery = "SELECT * FROM tblColors WHERE (color_id) = " & v_sTheme
		Set oData = New wcData
		getColorPalette = oData.GetArray(sQuery)
		Set oData = nothing
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		getUserScopes()
	'	Purpose: 	get colors for theme
	'	Return: 	array
	' Modifications:
	'	Date:		Name:	Description:
	'	9/24/03		JEA		Created
	'-------------------------------------------------------------------------
	Private Function getUserScopes(ByVal v_lUserID)
		dim oData
		dim sQuery
		
		if v_lUserID = 0 then
			' 0 is default group visibility for guests
			sQuery = "SELECT scope_id, scope_name, 0 FROM tblScopes ORDER BY scope_id DESC"
		else	
			sQuery = "SELECT us.scope_id, s.scope_name, us.visible" _
				& " FROM (tblUserScopes AS us RIGHT OUTER JOIN" _
				& " tblScopes AS s ON (us.scope_id = s.scope_id))" _
				& " WHERE user_id = " & v_lUserID _
				& " ORDER BY us.scope_id DESC"
		end if
		
		Set oData = New wcData
		getUserScopes = oData.GetArray(sQuery)
		Set oData = nothing
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		getDefaultAccess()
	'	Purpose: 	get colors for theme
	'	Return: 	number
	' Modifications:
	'	Date:		Name:	Description:
	'	9/24/03		JEA		Created
	'-------------------------------------------------------------------------
	Private Function getDefaultAccess(ByVal v_lUserID)
		dim oData
		dim aData
		dim sQuery
		sQuery = "SELECT default_access FROM tblUsers WHERE (user_id) = " & v_lUserID
		Set oData = New wcData
		aData = oData.GetArray(sQuery)
		Set oData = nothing
		getDefaultAccess = aData(0,0)
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		getUserGroups()
	'	Purpose: 	get colors for theme
	'	Return: 	array
	' Modifications:
	'	Date:		Name:	Description:
	'	9/24/03		JEA		Created
	'-------------------------------------------------------------------------
	Private Function getUserGroups(ByVal v_lUserID, ByVal v_lAccessLevel)
		dim oData
		dim sQuery
		if v_lUserID = 0 then
			' for guests only return groups that have public events
			sQuery = "SELECT DISTINCT g.group_id, g.group_name, 1, 0 " _
				& "FROM tblGroups AS g INNER JOIN tblEventGroupScopes egs " _
				& "ON g.group_id = egs.group_id " _
				& "WHERE egs.scope_id = " & g_PUBLIC _
				& " ORDER BY g.group_name"
		elseif v_lAccessLevel > g_NO_ACCESS then
			' retrieve user->group specific permissions
			sQuery = "SELECT g.group_id, g.group_name, p.visible, p.access_level" _
				& " FROM (tblGroups AS g LEFT OUTER JOIN" _
				& " (SELECT group_id, visible, access_level FROM" _
				& " tblPermissions WHERE user_id = " & v_lUserID _
				& ") AS p ON p.group_id = g.group_id)" _
				& " ORDER BY group_name"
		else
			' return default group permissions
			sQuery = "SELECT g.group_id, g.group_name, p.visible, p.access_level" _
				& " FROM (tblPermissions AS p INNER JOIN" _
				& " tblGroups AS g ON (p.group_id = g.group_id))" _
				& " WHERE p.user_id = " & v_lUserID _
				& " ORDER BY group_name"
		end if
		Set oData = New wcData
		getUserGroups = oData.GetArray(sQuery)
		Set oData = nothing
	End Function
End Class
%>