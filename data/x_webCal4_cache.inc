<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' this allows entire calendar pages to be cached
' it should be included AFTER data.inc

dim g_arCache
dim x

g_arCache = Application(g_unique & "Cache")

' retrieve page HTML from application cache (updated 3/7/01)
' returns string ---------------------------------------------------------
Function readCache(ByVal v_strQuery, ByVal v_intView)
	dim strExpires
	dim strGrid
	dim strFirst
	dim strLast
	dim strHTML
	dim y
	
	strHTML = ""
	strGrid = getGridString(v_intView)
	v_strQuery = trimQuery(v_strQuery)

	for x = 0 to UBound(g_arCache)
		if g_arCache(x)(g_CACHE_SQL) = v_strQuery And _
		   g_arCache(x)(g_CACHE_GRID) = strGrid then
			
			strExpires = g_arCache(x)(g_CACHE_EXPIRE_DATE)
			strFirst = g_arCache(x)(g_CACHE_START_DATE)
			strLast = g_arCache(x)(g_CACHE_END_DATE)
			strHTML = g_arCache(x)(g_CACHE_HTML)
			' demote other cached pages
			for y = 1 to x
				g_arCache(y) = g_arCache(y - 1)
			next
			' move page to head of cache
			g_arCache(0) = Array(v_strQuery, strHTML, strGrid, strFirst, strLast, strExpires)
			Application.Lock
			Application(g_unique & "Cache") = g_arCache
			Application.Unlock
			exit for
		end if
	next
	readCache = strHTML
End Function

' save HTML to application cache (updated 3/7/01)
' updates application object ---------------------------------------------
Sub saveCache(ByVal v_strQuery, ByVal v_strHTML, ByVal v_strFirst, _
	ByVal v_strLast, ByVal v_intView)
	
	dim strExpires
	dim strGrid

	strGrid = getGridString(v_intView)
	v_strQuery = trimQuery(v_strQuery)
	
	for x = UBound(g_arCache) to 1 step - 1	' FIFO
		g_arCache(x) = g_arCache(x - 1)
	next
	
	if v_strLast < Date then
		' page never needs to automatically expire
		strExpires = "1/1/15"
	elseif v_strFirst > Date then
		' expire cache when date is reached
		strExpires = v_strFirst
	elseif v_strFirst <= Date AND v_strLast >= Date then
		' page is only good for today
		strExpires = Date
	end if
	
'	response.write strLast & ", " & strExpires
	
	g_arCache(0) = Array(v_strQuery, v_strHTML, strGrid, v_strFirst, v_strLast, strExpires)
	Application.Lock
	Application(g_unique & "Cache") = g_arCache
	Application.Unlock
End Sub

' update cached pages to show new event (updated 3/7/01)
' updates application array ----------------------------------------------
Sub updateCache(ByVal v_arEventDates)
	dim x, y		' loop counters
	dim bExpire
	
	for x = 0 to UBound(g_arCache)
		for	y = 0 to UBound(v_arEventDates)
			if v_arEventDates(y) >= g_arCache(x)(g_CACHE_START_DATE) And _
			   v_arEventDates(y) <= g_arCache(x)(g_CACHE_END_DATE) then
				' this cached page needs to be expired
				
				'bExpire = true
				exit for
			end if
		next
	next
End Sub

' trim query to return only unique portion (updated 3/7/01)
' returns string ---------------------------------------------------------
Function trimQuery(ByVal v_strQuery)
	v_strQuery = Right(v_strQuery, Len(v_strQuery) - 283)
	v_strQuery = Left(v_strQuery, Len(v_strQuery) - 22)
	trimQuery = v_strQuery
End Function

' build grid settings string (updated 3/7/01)
' returns string ---------------------------------------------------------
Function getGridString(ByVal v_intView)
	dim strWE
	strWE = "-0"	
	if Session(g_unique & "Weekends") Or v_intView = g_MONTH then strWE = "-1"
	getGridString = Join(Session(g_unique & "Segments")(v_intView), "-") & strWE
End Function
%>