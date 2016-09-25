<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' requires constants.inc

dim g_unique		' unique name for data in this copy of the calendar
dim g_strDelim		' date delimiter (' for Access, # for SQL)
dim g_strDSN		' connection string

g_unique = "newCal"
g_strDelim = "#"

' OLEDB connection string
g_strDSN = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
	& Server.Mappath("data/" & g_unique & ".mdb")
	
' OLEDB connection string for MSSQL
'g_strDSN = "Provider=SQLOLEDB.1;User ID=[login];Password=[password];" _
'	& "Initial Catalog=[database];Data Source=[server];Network=DBMSSOCN"

if Session(g_unique & "LCID") <> "" then Session.LCID = Session(g_unique & "LCID")

' generate array for given query string (updated 1/4/2001)
' returns array ----------------------------------------------------------
Function getArray(ByVal v_strQuery, ByRef r_oConn)
	dim oRS
	dim arData
	Set oRS = newRecordSet(v_strQuery, r_oConn)
	if not oRS.EOF then arData = oRS.GetRows
	oRS.Close : Set oRS = nothing
	getArray = arData
End Function

' generate array for single row returning query (updated 3/4/01)
' returns array ----------------------------------------------------------
Function getRowArray(ByVal v_strQuery, ByRef r_oConn)
	dim strRowList
	dim oRS
	
	Set oRS = newRecordSet(v_strQuery, r_oConn)
	if not oRS.EOF then
		strRowList = oRS.GetString(2, , ",", ",")
		' remove trailing comma
		strRowList = Left(strRowList, Len(strRowList) - 1)
		getRowArray = Split(strRowList, ",")
	else
		getRowArray = ""
	end if
	oRS.Close : Set oRS = nothing
End Function

' creates javascript array literal for given query (updated 2/24/01)
' returns string ---------------------------------------------------------
function getJSArray(ByVal v_strQuery, ByRef r_oConn)
	dim strJSArray
	dim oRS
	
	Set oRS = newRecordSet(v_strQuery, r_oConn)
	strJSArray = oRS.GetString(2, , """,""", """],[""")
	oRS.Close : Set oRS = nothing
	if strJSArray <> "" then strJSArray = "[""" & Left(strJSArray, Len(strJSArray) - 4) & "]"
	getJSArray = strJSArray 
end function

' create and open a recordset with the given query and connection (updated 2/24/01)
' returns open, disconnected recordset -----------------------------------
Private Function newRecordSet(ByVal v_strQuery, ByRef r_oConn)
	dim oRS
	dim bNewObject
	bNewObject = newConnection(r_oConn)
	Set oRS = Server.CreateObject("ADODB.Recordset")
	Set oRS.ActiveConnection = r_oConn
	oRS.CursorLocation = adUseClient
	oRS.Open v_strQuery, , adOpenForwardOnly, adLockReadOnly, adCmdText
	Set oRS.ActiveConnection = nothing
	if bNewObject then r_oConn.Close : Set r_oConn = nothing
	Set newRecordSet = oRS
End Function

' create a new connection object if needed (updated 2/22/01)
' returns boolean --------------------------------------------------------
Private Function newConnection(ByRef r_oConn)
	dim bNewObject
	bNewObject = false
	if Not IsObject(r_oConn) then
		' create and open new connection
		Set r_oConn = Server.CreateObject("ADODB.Connection")
		r_oConn.Open g_strDSN
		bNewObject = true
	end if
	newConnection = bNewObject
End Function

' Build the query string to retrieve events for view (updated 1/7/2001)
' returns string ---------------------------------------------------------
Function makeQuery(ByVal v_strFirst, ByVal v_strLast)
	if Session(g_unique & "Query") <> g_sNO_EVENTS then
		makeQuery = "SELECT e.event_id, e.event_title, e.event_recur, " _
			& "e.event_color, e.time_start, e.time_end, ed.event_date " _
			& "FROM ((tblEvents AS e INNER JOIN tblEventDates AS ed " _
			& "ON e.event_id = ed.event_id) INNER JOIN tblEventGroupScopes AS egs " _
			& "ON e.event_id = egs.event_id) WHERE (ed.event_date " _
			& "BETWEEN " & g_strDelim _
			& sqlDate(v_strFirst) _
			& " 12:00:00 AM" & g_strDelim & " AND " & g_strDelim _
			& sqlDate(v_strLast)  _
			& " 11:59:59 PM" & g_strDelim & ")" & Session(g_unique & "Query") _
			& " ORDER BY e.time_start"
	else
		makeQuery = g_sNO_EVENTS
	end if
End Function

' get the list of groups and scopes available to this user (updated 2/27/2001)
' updates three session variables ----------------------------------------
Sub initDataAccess(ByRef r_oConn)
	dim strQuery
	dim intUserID
	dim bNewObject
	dim arScopes			' array of scope settings
	dim arGroups			' array of group settings
	dim oRS					' recordset object
	dim intDefaultAccess	' default group access
	dim x					' loop counter
	
	intUserID = Session(g_unique & "UserID")

	bNewObject = newConnection(r_oConn)
	
	if intUserID then
		' logged in user
		strQuery = "SELECT us.scope_id, s.scope_name, us.visible" _
			& " FROM (tblUserScopes AS us RIGHT OUTER JOIN" _
			& " tblScopes AS s ON (us.scope_id = s.scope_id))" _
			& " WHERE user_id = " & intUserID _
			& " ORDER BY us.scope_id DESC"
		arScopes = getArray(strQuery, r_oConn)
		
		' get default access
		strQuery = "SELECT default_access FROM tblUsers WHERE (user_id) = " & intUserID
		Set oRS = Server.CreateObject("ADODB.Recordset")
		oRS.Open strQuery, r_oConn, adOpenForwardOnly, adLockReadOnly, adCmdText
		intDefaultAccess = oRS("default_access")
		oRS.Close : Set oRS = nothing
		
		if intDefaultAccess > g_NO_ACCESS then
			' assign default to groups with no explicit access
			strQuery = "SELECT g.group_id, g.group_name, p.visible, p.access_level" _
				& " FROM (tblGroups AS g LEFT OUTER JOIN" _
				& " (SELECT group_id, visible, access_level FROM" _
				& " tblPermissions WHERE user_id = " & intUserID _
				& ") AS p ON p.group_id = g.group_id)" _
				& " ORDER BY group_name"
			arGroups = getArray(strQuery, r_oConn)
			if intDefaultAccess = g_ADMIN_ACCESS then
				' admin access overrides all others
				for x = 0 to UBound(arGroups, 2)
					arGroups(g_GROUP_ACCESS, x) = g_ADMIN_ACCESS
				next
			else
				for x = 0 to UBound(arGroups, 2)
					if arGroups(g_GROUP_ACCESS, x) = "" or _
						arGroups(g_GROUP_ACCESS, x) = -1 then
						
						arGroups(g_GROUP_ACCESS, x) = intDefaultAccess
					end if
				next
			end if
		else		
			strQuery = "SELECT g.group_id, g.group_name, p.visible, p.access_level" _
				& " FROM (tblPermissions AS p INNER JOIN" _
				& " tblGroups AS g ON (p.group_id = g.group_id))" _
				& " WHERE p.user_id = " & intUserID _
				& " ORDER BY group_name"
			arGroups = getArray(strQuery, r_oConn)
		end if
	else
		' guest user
		' 0 is default group visibility for guests
		strQuery = "SELECT scope_id, scope_name, 0" _
			& " FROM tblScopes ORDER BY scope_id DESC"
		arScopes = getArray(strQuery, r_oConn)
		
		' only "public" scope is visible to guests
		arScopes(g_VISIBLE, 0) = 1
		
		' only return groups that have public events
		strQuery = "SELECT DISTINCT g.group_id, g.group_name, 1, 0 " _
			& "FROM tblGroups AS g INNER JOIN tblEventGroupScopes egs " _
			& "ON g.group_id = egs.group_id " _
			& "WHERE egs.scope_id = " & g_PUBLIC _
			& " ORDER BY g.group_name"
		arGroups = getArray(strQuery, r_oConn)
	end if
	Session(g_unique & "Scopes") = arScopes
	Session(g_unique & "Groups") = arGroups
	Call makeUserQuery(arGroups, arScopes)
	
	' clean up objects
	Erase arGroups : Erase arScopes
	if bNewObject then r_oConn.Close : Set r_oConn = nothing
End Sub

' build the user-specific portion of event query (updated 3/3/01)
' updates session variable -----------------------------------------------
Sub makeUserQuery(ByVal v_arGroups, ByVal v_arScopes)
	dim strGroups			' query condition for groups
	dim strScopes			' query condition for scopes
	dim strQuery

	strGroups = makeWHERE(v_arGroups, "group_id", g_GROUP_ID)
	strScopes = makeWHERE(v_arScopes, "scope_id", g_SCOPE_ID)
	strQuery = ""
	if strGroups <> g_sNO_EVENTS and strScopes <> g_sNO_EVENTS then
		if strGroups <> "" AND strScopes <> "" then	strQuery = " AND "
		strQuery = Trim(strGroups & strQuery & strScopes)
		if strQuery <> "" then
			Session(g_unique & "Query") = " AND " & strQuery
		else
			Session(g_unique & "Query") = ""
		end if
	else
		Session(g_unique & "Query") = g_sNO_EVENTS
	end if
End Sub

' generate conditions of all queries for this user (updated 3/3/2001)
' returns string ---------------------------------------------------------
Private Function makeWHERE(ByVal v_arParms, ByVal v_strField, ByVal v_intID)
	dim strWhere		' SQL WHERE condition
	dim intAllowCount	' one-based count of matching parms
	dim intParmCount	' zero-based count of parms
	dim x				' loop counter
	
	intAllowCount = 0
	intParmCount = UBound(v_arParms, 2)
	for x = 0 to intParmCount
		if v_arParms(g_VISIBLE, x) then
			' add scopes and groups that are visible to this user
			strWhere = strWhere & v_arParms(v_intID, x) & ","
			intAllowCount = intAllowCount + 1
		end if
	next
	' trim exra comma
	if strWhere <> "" then strWhere = Left(strWhere, Len(strWhere) - 1)
	
	if intAllowCount = 1 then
		' match single parameter
		strWhere = "egs." & v_strField & "=" & strWhere
	elseif intAllowCount <> (intParmCount + 1) AND intAllowCount > 0 then
		strWhere = "egs." & v_strField & " IN (" & strWhere & ")"
	elseif intAllowCount > 0 then
		' match all parameters, so no limitation
		strWhere = ""
	else
		' no parameters match
		strWhere = g_sNO_EVENTS
	end if
	
	makeWHERE = strWhere
End Function

' Access only accepts mm/dd/yy dates, so force that format
' returns string----------------------------------------------------------
Function sqlDate(ByVal v_strDate)
	sqlDate = Month(v_strDate) & "/" & Day(v_strDate) & "/" & Year(v_strDate)
End Function

' converts a binary string to decimal
' returns integer---------------------------------------------------------
Function BinToDec(bin)
	Dim x			' counter
	Dim intNum		' decimal number
	intNum = 0		' initialize
	For x = 1 to Len(bin)
		intNum = intNum + (Mid(StrReverse(bin), x, 1) * 2^(x - 1))
	Next
	BinToDec = intNum
End Function

' converts an integer to a binary string
' returns string----------------------------------------------------------
Function DecToBin(dec)
	dim strBin		' binary string
	strBin = ""		' initialize
	While dec >= 1
		strBin = Abs(CBool(dec Mod 2)) & strBin
		dec = Int(dec / 2)
	Wend
	DecToBin = strBin
End Function

' formats the milliseconds returned by the JScript function, below (updated 3/3/01)
' returns string----------------------------------------------------------
Function showElapsedTime(ByVal v_ms)
	dim strMil
	dim strSec
	dim strMin
	
	if v_ms = 0 then
		timer = g_sMSG_SUBSECOND
	else
		strSec = Int(v_ms / 1000)
		strMil = v_ms Mod 1000
		
		strMin = Int(strSec / 60)
		strSec = strSec Mod 60
		
		if strMin <> 0 then
			strMin = strMin & " " & g_sMSG_MINUTE & ", "
		else
			strMin = ""
		end if
		
		if Not Right(strMil,1) then
			strMil = Left(strMil, Len(strMil) - 1)
		end if
		
		showElapsedTime = strMin & strSec & "<b>.</b>" & strMil & " " & g_sMSG_SECONDS
	end if
End Function
%>

<script language=JScript runat=Server>
// time db operations to the ms--not possible with VBScript
// returns integer--------------------------------------------------------
function milliTime() {
	var d = new Date();
	return d.getTime()
}

// unused hex function
function getHex(num){
   hexStr = "0123456789ABCDEF";
   hex="";
   if (num>=16) {
      hex = hexStr.substr(parseInt(num/16),1);
      num = num%16;
   }
   hex += hexStr.substr(num,1);

   return hex;
}

</script>