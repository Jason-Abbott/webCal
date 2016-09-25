<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' Last updated 2/20/2001

' retrieve all values necessary to populate the edit form

dim m_strQuery			' query passed to database
dim m_strTitle			' event title
dim m_strDescription	' event description
dim m_strRecur			' event recurrence type
dim m_strEventClr		' event color
dim m_strNoTime			' does event occur at a particular time
dim m_strSkipWE			' if recurring, should weekend be skipped?
dim m_strShowTime		' string to disable time display
dim m_strStartDate		' event start date
dim	m_intStartHour		' event start hour
dim	m_strStartMin		' minutes past start hour
dim m_strEndDate		' event start date
dim	m_intEndHour		' event end hour
dim	m_strEndMin			' minutes past end hour
dim m_strGroups			' comma-delimited string of permitted groups
dim m_strJSGroups		' javascript array of groups
dim m_strEditType		' all/future/current/new etc.
dim m_strView			' calendar view (month/week) to return to
dim m_intEventID		' event id
dim x

' default values (use military time)
m_intStartHour = 8
m_strStartMin = "00"
m_intEndHour = 17
m_strEndMin = "00"
m_strNoTime = ""
m_strShowTime = ""
m_strSkipWE = ""
m_strEventClr = "black"

' get strings to populate fields (updated 2/23/01)
' updates page scope variables -------------------------------------------
Function getEventFields(ByVal v_bEdit, ByVal v_intEventID, _
	ByVal v_strTimeScope, ByVal v_strDate, ByRef r_strJSEventScopes)
	
	dim arEvent
	dim arGroups
	dim oConn
	dim strQuery
	dim strGroupIDs			' used in query in getGroupNames
	dim x, y				' loop counters

	' get a local copy of the groups this user has access to
		
	Set oConn = Server.CreateObject("ADODB.Connection")
	oConn.Open g_strDSN
	
	if v_bEdit then
		' retrieve event data
		strQuery = "SELECT e.event_id, e.event_title, e.event_recur, " _
			& "e.event_color, e.time_start, e.time_end, ed.event_date, " _
			& "e.event_description, e.skip_weekends " _
			& "FROM tblEvents e INNER JOIN tblEventDates ed " _
			& "ON (e.event_id = ed.event_id) " _
			& "WHERE (e.event_id)=" & v_intEventID _
			& " ORDER BY ed.event_date"
		arEvent = getArray(strQuery, oConn)
		m_strTitle = arEvent(g_EVENT_TITLE,0)
		m_strDescription = arEvent(g_EVENT_DESC,0)
		m_strEventClr = arEvent(g_EVENT_COLOR,0)
			
		' these need to be broken out for separate form fields
		if arEvent(g_TIME_START,0) <> "" then
			m_intStartHour = Hour(arEvent(g_TIME_START,0))
			m_strStartMin = Minute(arEvent(g_TIME_START,0))
			m_intEndHour = Hour(arEvent(g_TIME_END,0))
			m_strEndMin = Minute(arEvent(g_TIME_END,0))
		else
			m_strNoTime = " checked"
			m_strShowTime = " disabled"
		end if
			
		if arEvent(g_EVENT_SKIP_WE,0) then m_strSkipWE = " checked"
		
		' get recurrence information
		Select Case v_strTimeScope
			Case "future"
				m_strRecur = arEvent(c_Recur,0)
				m_strStartDate = v_strDate
				m_strEndDate = DateValue(arEvent(g_EVENT_DATE,UBound(arEvent,2)))
			Case "all"
				m_strRecur = arEvent(c_Recur,0)
				m_strStartDate = DateValue(arEvent(g_EVENT_DATE,0))
				m_strEndDate = DateValue(arEvent(g_EVENT_DATE,UBound(arEvent,2)))
			Case else
				m_strRecur = "none"
				m_strStartDate = v_strDate
				m_strEndDate = ""
				m_strSkipWE = ""
		End Select
		
		' if no scope was sent then we're editing an event that
		' doesn't recur, in which case we want to edit "all"
		' instances
		if v_strTimeScope <> "" then
			m_strEditType = v_strTimeScope
		else
			m_strEditType = "all"
		end if
		m_strView = Request.Form("view")

		r_strJSEventScopes = getEditScopes(oConn, v_intEventID)
	else
		m_strTitle = ""
		m_strDescription = ""
		m_strRecur = "none"
		m_strStartDate = Request.QueryString("date")
		m_strEndDate = ""
		m_strEditType = "new"
		m_strView = Request.QueryString("view")
			
		r_strJSEventScopes = getNewScopes()
	end if

	oConn.Close : Set oConn = nothing
End Function

' create string of groups allowed on this event (updated 2/24/01)
' returns string and updates page scope variable -------------------------
Private Function getEditScopes(ByRef r_oConn, ByVal v_intEventID)
	dim strQuery
	strQuery = "SELECT g.group_id, g.group_name, es.scope_id, " _
		& "g.allow_title_html, g.allow_desc_html, g.allow_loc_html " _
		& "FROM (tblGroups AS g LEFT OUTER JOIN " _
		& "(SELECT group_id, scope_id FROM tblEventGroupScopes WHERE event_id = " _
		& v_intEventID & ") AS es ON es.group_id = g.group_id)" _
		& " WHERE g.group_id IN (" & getUserGroups(g_ADD_ACCESS) _
		& ") ORDER BY g.group_name"
	getEditScopes = getJSArray(strQuery, r_oConn)
End Function

' create string of groups allowed on this event (updated 2/24/01)
' returns string and updates page scope variables ------------------------
Private Function getNewScopes()
	dim strQuery
	dim strGroupList
	' 0 is default user scope for new events
	strQuery = "SELECT group_id, group_name, 0, " _
		& "allow_title_html, allow_desc_html, allow_loc_html " _
		& "FROM tblGroups WHERE group_id IN (" _
		& getUserGroups(g_ADD_ACCESS) & ") ORDER BY group_name"
	getNewScopes = getJSArray(strQuery, "")
End Function

' get list of groups current user has at least given access (updated 2/27/01)
' returns string ---------------------------------------------------------
Function getUserGroups(ByVal v_lAccess)
	dim arGroups
	dim strList
	dim x
	arGroups = Session(g_unique & "Groups")
	for x = 0 to UBound(arGroups, 2)
		if arGroups(g_GROUP_ACCESS, x) >= v_lAccess then
			strList = strList & arGroups(g_GROUP_ID, x) & ","
		end if
	next
	if strList <> "" then strList = Left(strList, Len(strList) - 1)
	getUserGroups = strList
End Function
%>