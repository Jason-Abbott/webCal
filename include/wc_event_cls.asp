<%
' Copyright 1996-2004 Jason Abbott (webcal@webott.com)
' methods for event form

Class wcEvent

	Private m_oLayout
	Private m_aRecurType
	Private m_aRecurName
	Private m_aHourName
	Private m_EVENT_FIELDS
	
	Private Sub Class_Initialize()
		Set m_oLayout = New wcLayout
		m_EVENT_FIELDS = 14
		m_aRecurType = Array("none","daily","weekly","2weeks","monthly","yearly")
		m_aRecurName = Array("None","Daily","Weekly","Every other wk","Monthly","Yearly")
		m_aHourName = Array("12 AM","1 AM","2 AM","3 AM","4 AM","5 AM","6 AM","7 AM","8 AM","9 AM","10 AM","11 AM","12 PM","1 PM","2 PM","3 PM","4 PM","5 PM","6 PM","7 PM","8 PM","9 PM","10 PM","11 PM")
	End Sub
	
	Private Sub Class_Terminate()
		Set m_oLayout = nothing
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		writeEditFormHTML()
	'	Purpose: 	generates the body of the month calendar
	' Modifications:
	'	Date:		Name:	Description:
	'	10/20/03	JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub writeEditFormHTML()
		Const sFORM_NAME = "frmEdit"
		
		dim aColorHex(5)
		dim aColorName(5)
		dim aEvent
		dim x
		
		aColorHex = Array("000000","0000ff","aa00aa","ff0000","00cc00","ffbb00")
		aColorName = Array("black","blue","purple","red","green","orange")
		aEvent = getData()
		
		with response
			.write "<table border='0' cellpadding='4' cellspacing='0'>"
			.write "<form name='"
			.write sFORM_NAME
			.write "' method='post'><tr><td colspan='2'>Event Details</td>"
			.write "<tr><td>Title<br>"
			.write "<input name='fldTitle' type='text' size='35' max='50' value='"
			.write aEvent(g_EVENT_TITLE)
			.write "'></td><td rowspan='2'>Description<br>"
			.write "<textarea cols='24' name='fldDescription' rows='14' wrap='virtual'>"
			.write aEvent(g_EVENT_DESC)
			.write "</textarea></td><tr><td>"
			' timing table
			.write "<table cellpadding='2' cellspacing='2' border='0' width='100%'><tr>"
			.write "<td>Date<br>"
			.write "<input name='fldStartDate' type='text' size='10' value='"
			.write aEvent(START_DATE)
			.write "'><input type='button' value='&gt;' onClick=""calPop('frmEdit','fldStartDate');"">"
			.write "<br>Recurrence<br><select name='fldEventRecur' onChange='newRecur(this);'>"
			.write m_strRecurList
			.write "</select><br>until<br>"
			.write "<input name='fldEndDate' type='text' size='10' value='"
			.write aEvent(END_DATE)
			.write "'><input type='button' value='&gt;' onClick=""calPop('frmEdit','fldEndDate');"">"
			.write "<br><input type='checkbox' name='fldSkipWE'"
			if aEvent(g_EVENT_SKIP_WE) then .write " checked"
			.write ">Skip weekends</td>"
			
			.write "<td>Time<br><nobr><select name='fldStartHour'"
			.write m_strShowTime
			.write ">"
			.write m_strStartHrList
			.write "</select><select name='fldStartMin'"
			.write m_strShowTime
			.write ">"
			.write m_strStartMinList
			.write "</select></nobr><br>until<br><nobr><select name='fldEndHour'"
			.write m_strShowTime
			.write ">"
			.write m_strEndHourList
			.write "</select><select name='fldEndMin'"
			.write m_strShowTime
			.write ">"
			.write m_strEndMinList
			.write "</select></nobr><p><input type='checkbox' name='fldNoTime'"
			.write m_strNoTime
			.write "onClick='newTimeCheck(this);'>No Specific Time</td>"
			
			.write "<tr><td colspan='2'Display<br>"
			.write "<table cellspacing='1' cellpadding='0' border='0' width='100%'><tr>"
			
			for x = 0 to UBound(aColorHex)
				.write "<td style='background-color: #"
				.write aColorHex(x)
				.write ";'><input type='radio' name='fldEventColor' value='"
				.write aColorName(x)
				.write "'"
				if m_strEventClr = aColorName(x) then .write " checked"
				.write "></td>"
			next

			.write "</table><center>In"
			.write "<select name='fldGroup' onChange='newGroup(this);'></select>,"
			.write "visible to <select name='fldShowTo' onChange='newUserScope(this);'>"
			.write "<option value='0'>none"
			.write "<option value='1'>only me"
			.write "<option value='2'>group"
			.write "<option value='3'>public"
			.write "</select></center></td></table></td>"
			
			.write "<tr><td colspan='2'>"
			.write "<input type='button' value='Save' onClick=""saveEvent('frmEdit', false);"">"
			.write "<input type='button' value='Save & Add Another'	onClick=""saveEvent('frmEdit', true);"">"
			.write "<input type='button' value='Cancel' onClick=""goPage('"
			.write Request.Form("fldURL")
			.write "','frmEdit');""></td></table>"
		end with
	End Sub
	
	Private Sub writeRecurOption
	
	End Sub

	'-------------------------------------------------------------------------
	'	Name: 		getData()
	'	Purpose: 	retrieve event data
	'	Return:		array
	' Modifications:
	'	Date:		Name:	Description:
	'	2/23/01		JEA		Creation
	'	10/20/03	JEA		Updated
	'-------------------------------------------------------------------------
	Private Function getData(ByVal v_bEdit, ByVal v_lEventID, ByVal v_sTimeScope, ByVal v_strDate)
		dim aEvent			' event array to pass back
		dim aData			' data from db
		dim arGroups
		dim sQuery			' SQL
		dim strGroupIDs		' used in query in getGroupNames
		dim oData
		dim x, y			' loop counters
	
		' get a local copy of the groups this user has access to
			
		if v_bEdit then
			' retrieve event data
			sQuery = "SELECT e.event_id, e.event_title, e.event_recur, " _
				& "e.event_color, e.time_start, e.time_end, ed.event_date, " _
				& "e.event_description, e.skip_weekends " _
				& "FROM tblEvents e INNER JOIN tblEventDates ed " _
				& "ON (e.event_id = ed.event_id) " _
				& "WHERE (e.event_id)=" & v_lEventID _
				& " ORDER BY ed.event_date"
				
			Set oData = New wcData
			aData = oData.getArray(sQuery)
			aEvent = oData.dimDown(aData, 0)
			Set oData = nothing
			
			ReDim Preserve aEvent(m_EVENT_FIELDS)
			
			' these need to be broken out for separate form fields
			if Not IsVoid(aEvent(g_TIME_START)) then
				m_intStartHour = Hour(aEvent(g_TIME_START,0))
				m_strStartMin = Minute(aEvent(g_TIME_START,0))
				m_intEndHour = Hour(aEvent(g_TIME_END,0))
				m_strEndMin = Minute(aEvent(g_TIME_END,0))
			else
				m_strNoTime = " checked"
				m_strShowTime = " disabled"
			end if
				
			' get recurrence information
			Select Case v_sTimeScope
				Case "future"
					m_strRecur = aEvent(c_Recur,0)
					m_strStartDate = v_strDate
					m_strEndDate = DateValue(aEvent(g_EVENT_DATE,UBound(aEvent,2)))
				Case "all"
					m_strRecur = aEvent(c_Recur,0)
					m_strStartDate = DateValue(aEvent(g_EVENT_DATE,0))
					m_strEndDate = DateValue(aEvent(g_EVENT_DATE,UBound(aEvent,2)))
				Case else
					m_strRecur = "none"
					m_strStartDate = v_strDate
					m_strEndDate = ""
					m_strSkipWE = ""
			End Select
			
			' if no scope was sent then we're editing an event that
			' doesn't recur, in which case we want to edit "all"
			' instances
			if v_sTimeScope <> "" then
				m_strEditType = v_sTimeScope
			else
				m_strEditType = "all"
			end if
			m_strView = Request.Form("view")
	
			r_strJSEventScopes = getEditScopes(oConn, v_lEventID)
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

End Class







' retrieve all values necessary to populate the edit form

dim m_sQuery			' query passed to database
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



' create string of groups allowed on this event (updated 2/24/01)
' returns string and updates page scope variable -------------------------
Private Function getEditScopes(ByRef r_oConn, ByVal v_lEventID)
	dim sQuery
	sQuery = "SELECT g.group_id, g.group_name, es.scope_id, " _
		& "g.allow_title_html, g.allow_desc_html, g.allow_loc_html " _
		& "FROM (tblGroups AS g LEFT OUTER JOIN " _
		& "(SELECT group_id, scope_id FROM tblEventGroupScopes WHERE event_id = " _
		& v_lEventID & ") AS es ON es.group_id = g.group_id)" _
		& " WHERE g.group_id IN (" & getUserGroups(g_ADD_ACCESS) _
		& ") ORDER BY g.group_name"
	getEditScopes = getJSArray(sQuery, r_oConn)
End Function

' create string of groups allowed on this event (updated 2/24/01)
' returns string and updates page scope variables ------------------------
Private Function getNewScopes()
	dim sQuery
	dim strGroupList
	' 0 is default user scope for new events
	sQuery = "SELECT group_id, group_name, 0, " _
		& "allow_title_html, allow_desc_html, allow_loc_html " _
		& "FROM tblGroups WHERE group_id IN (" _
		& getUserGroups(g_ADD_ACCESS) & ") ORDER BY group_name"
	getNewScopes = getJSArray(sQuery, "")
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