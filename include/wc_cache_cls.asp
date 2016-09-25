<%
' Copyright 1996-2004 Jason Abbott (webcal@webott.com)
' allows entire HTML pages to be output cached

Class wcCache
	Public Page

	Private Sub Class_Initialize()
		' 2D jagged array of pages and settings
		Page = Application(g_sDB_NAME & "Cache")
	End Sub
	
	Private Sub Class_Terminate()
		' update application variable with page cache
		Application.Lock
		Application(g_sDB_NAME & "Date") = Date
		Application(g_sDB_NAME & "Cache") = Page
		Application.Unlock
		Erase Page
	End Sub

	'-------------------------------------------------------------------------
	'	Name: 		Read()
	'	Purpose: 	retrieve HTML from cache based on query and view
	'	Return: 	string
	'Modifications:
	'	Date:		Name:	Description:
	'	3/7/01		JEA		Creation
	'-------------------------------------------------------------------------
	Public Function Read(ByVal v_sQuery, ByVal v_lViewID)
		dim dtExpires
		dim sGrid
		dim dtFirst
		dim dtLast
		dim sHTML
		dim y
		
		sHTML = ""
		sGrid = getGridString(v_lViewID)
		v_sQuery = trimQuery(v_sQuery)
	
		for x = 0 to UBound(Cache)
			if Page(x)(g_CACHE_SQL) = v_sQuery And Page(x)(g_CACHE_GRID) = sGrid then
				' retrieve cached page
				dtExpires = Page(x)(g_CACHE_EXPIRE_DATE)
				dtFirst = Page(x)(g_CACHE_START_DATE)
				dtLast = Page(x)(g_CACHE_END_DATE)
				sHTML = Page(x)(g_CACHE_HTML)
				' demote other cached pages
				for y = 1 to x
					Page(y) = Page(y - 1)
				next
				' move page to head of cache
				Page(0) = Array(v_sQuery, sHTML, sGrid, dtFirst, dtLast, dtExpires)
				exit for
			end if
		next
		readCache = sHTML
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		Save()
	'	Purpose: 	save HTML to application cache
	'Modifications:
	'	Date:		Name:	Description:
	'	3/7/01		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub Save(ByVal v_sQuery, ByVal v_sHTML, ByVal v_dtFirst, _
		ByVal v_dtLast, ByVal v_lViewID)
		
		dim dtExpires
		dim sGrid
	
		sGrid = getGridString(v_lViewID)
		v_sQuery = trimQuery(v_sQuery)
		
		for x = UBound(Page) to 1 step - 1
			' push existing cached pages back for FIFO
			Page(x) = Page(x - 1)
		next
		
		if v_dtLast < Date then
			' page never needs to automatically expire
			dtExpires = DateSerial(2015, 1, 1)
		elseif v_dtFirst > Date then
			' expire cache when date is reached
			dtExpires = v_dtFirst
		elseif v_dtFirst <= Date AND v_dtLast >= Date then
			' page is only good for today
			dtExpires = Date
		end if
		
		' response.write dtLast & ", " & dtExpires
		Page(0) = Array(v_sQuery, v_sHTML, sGrid, v_dtFirst, v_dtLast, dtExpires)
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		Update()
	'	Purpose: 	update cached pages for dates that include new event
	'Modifications:
	'	Date:		Name:	Description:
	'	3/7/01		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub Update(ByVal v_arEventDates)
		dim x, y		' loop counters
		dim bExpire
		
		for x = 0 to UBound(Page)
			for	y = 0 to UBound(v_arEventDates)
				if v_arEventDates(y) >= Page(x)(g_CACHE_START_DATE) And _
				   v_arEventDates(y) <= Page(x)(g_CACHE_END_DATE) then
					' this cached page needs to be expired
					
					'bExpire = true
					exit for
				end if
			next
		next
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		Refresh
	'	Purpose: 	cycle through cache array and remove expired pages
	'Modifications:
	'	Date:		Name:	Description:
	'	3/4/01		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub Refresh()
		dim arCache			' local copy of page cache
		dim lCacheSize		' number of pages cached
		dim x, y			' loop counters
		
		lCacheSize = Application(g_sDB_NAME & "Settings")(g_CACHE_SIZE)
		'arCache = Application(g_sDB_NAME & "Cache")
		for x = 0 to lCacheSize
			if Page(x)(g_CACHE_START_DATE) < Date then
				' the page has expired
				for y = x to lCacheSize - 1
					' shift following pages up in the cache
					Page(y) = Page(y + 1)
				next
			end if
		next
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		trimQuery()
	'	Purpose: 	trim query to return only unique portion
	'	Return: 	string
	'Modifications:
	'	Date:		Name:	Description:
	'	3/7/01		JEA		Creation
	'-------------------------------------------------------------------------
	Private Function trimQuery(ByVal v_sQuery)
		v_sQuery = Right(v_sQuery, Len(v_sQuery) - 283)
		v_sQuery = Left(v_sQuery, Len(v_sQuery) - 22)
		trimQuery = v_sQuery
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		readCache()
	'	Purpose: 	build grid settings sing
	'	Return: 	string
	'Modifications:
	'	Date:		Name:	Description:
	'	3/7/01		JEA		Creation
	'-------------------------------------------------------------------------
	Private Function getGridString(ByVal v_lViewID)
		dim sWE
		sWE = "-0"	
		if Session(g_sDB_NAME & "Weekends") Or v_lViewID = g_MONTH then sWE = "-1"
		getGridsing = Join(Session(g_sDB_NAME & "Segments")(v_lViewID), "-") & sWE
	End Function
End Class
%>