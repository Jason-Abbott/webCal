<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="./data/webCal4_data.inc"-->
<!--#include file="./include/webCal4_constants.inc"-->
<!--#include file="./include/webCal4_settings.inc"-->
<!--#include file="./include/webCal4_functions.inc"-->
<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' Last updated 3/7/2001

dim m_arScopes		' array of visibility levels
dim m_arGroups		' array of groups
dim m_intChecked
dim m_strChanged	' track groups already changed to visible
dim m_strScopes
dim m_strView		' calendar view
dim m_strDate		' calendar date
dim x, y			' loop counters
dim m_intGroupID	' group id
dim m_bSuccess
dim m_bGridInfo

m_strView = Request.Form("fldView")
m_strDate = Request.Form("fldDate")
m_bSuccess = true

' update user settings
m_strChanged = ""

' update group visibility settings
m_arGroups = Session(g_unique & "Groups")
For Each m_intGroupID In Request.Form("fldGroups")
	For x = 0 to UBound(m_arGroups, 2)
		If m_arGroups(g_GROUP_ID, x) = CInt(m_intGroupID) Then
			m_arGroups(g_VISIBLE, x) = 1
			m_strChanged = m_strChanged & "[" & x & "]"
		ElseIf InStr(m_strChanged, "[" & x & "]") = 0 Then
			' disable groups not already changed
			m_arGroups(g_VISIBLE, x) = 0
		End If
	Next
Next
Session(g_unique & "Groups") = m_arGroups

' update scope values based on checkbox
m_arScopes = Session(g_unique & "Scopes")
For x = 0 to UBound(m_arScopes, 2)
	If Request.Form("fldScope" & CStr(x)) = "on" Then
		m_arScopes(g_VISIBLE, x) = 1
		m_intChecked = m_intChecked + 1
	Else
		m_arScopes(g_VISIBLE, x) = 0
	End If
Next
Session(g_unique & "Scopes") = m_arScopes

Call makeUserQuery(m_arGroups, m_arScopes)

m_bGridInfo = saveGridInfo()

If Request.Form("fldDefault") = "on" Then
	' update database If "make default" was checked
	m_bSuccess = saveUserSettings(m_arGroups, m_arScopes, _
		Session(g_unique & "UserID"), m_strView, m_strDate, m_bGridInfo)
End If

If m_bSuccess Then response.redirect "webCal4_" & m_strView & ".asp?date=" & m_strDate

' update grid segment information (updated 3/3/01)
' returns boolean --------------------------------------------------------
Function saveGridInfo()
	dim intSegMins
	dim intSegStart
	dim intSegEnd
	dim bGridInfo
	dim intView
	
	bGridInfo = (Request.Form("fldSegMins") <> "")
	If bGridInfo Then
		select case m_strView
			case "week" : intView = g_WEEK
			case "day" : intView = g_DAY
		end select
	
		' update segment information
		intSegMins = Request.Form("fldSegMins")
		intSegStart = (60 / intSegMins) * Request.Form("fldHourStart")
		intSegEnd = ((60 / intSegMins) * Request.Form("fldHourEnd")) - 1

		Session(g_unique & "Segments")(intView) = Array(intSegMins, intSegStart, intSegEnd)
		response.write intSegMins & "," & Session(g_unique & "Segments")(intView)(0) : response.flush

		If Request.Form("fldWeekend") = "on" Then
			Session(g_unique & "Weekends") = True
		Else
			Session(g_unique & "Weekends") = False
		End If
	End If
	saveGridInfo = bGridInfo
End Function

' update database with new options (updated 3/3/01)
' updates database and returns boolean -----------------------------------
Function saveUserSettings(ByVal v_arGroups, ByVal v_arScopes, _
	ByVal v_intUserID, ByVal v_strView, ByVal v_strDate, ByVal v_bGridInfo)
	
	dim oConn		' ADODB connection object
	dim oRS			' ADODB recordset object
	dim strQuery
	dim bEmail		' send e-mail
	dim bUpdate		' update visibility
	dim bWeekend
	dim x			' loop counter
	
	'On Error Resume Next
	Set oConn = Server.CreateObject("ADODB.Connection")
	oConn.Open g_strDSN : oConn.BeginTrans

	' update group visibility	
	strQuery = "SELECT user_id, group_id, access_level, visible, " _
		& "send_email FROM tblPermissions WHERE " _
		& "user_id = " & v_intUserID
		
	Set oRS = Server.CreateObject("ADODB.RecordSet")
	oRS.CursorLocation = adUseClient
	oRS.Open strQuery, oConn, adOpenStatic, adLockBatchOptimistic, adCmdText
	For x = 0 to UBound(v_arGroups, 2)
		oRS.Filter = "group_id = " & v_arGroups(g_GROUP_ID, x)
		bUpdate = Not oRS.EOF	' not ready for update if EOF
		If Not bUpdate And v_arGroups(g_VISIBLE, x) = 1 Then
			' didn't find match so add new record
			oRS.AddNew
			oRS.Fields("user_id") = v_intUserID
			oRS.Fields("group_id") = v_arGroups(g_GROUP_ID, x)
			oRS.Fields("access_level") = -1
			oRS.Fields("send_email") = getEmailDefault(oConn, v_intUserID, bEmail)
			bUpdate = true		' now ready for update
		ElseIf bUpDate And v_arGroups(g_VISIBLE, x) = 0 Then
			If oRS("access_level") = -1 And _
				oRS("send_email") = getEmailDefault(oConn, v_intUserID, bEmail) Then
				' delete unecessary record (all its values are defaults)
				oRS.Delete : bUpdate = false
			End If
		End If
		If bUpdate Then oRS.Fields("visible") = v_arGroups(g_VISIBLE, x)
		oRS.Filter = adFilterNone
	Next
	oRS.UpdateBatch : oRS.Close : Set oRS = nothing
	
	' update scope visibility
	strQuery = "SELECT user_id, scope_id, visible FROM tblUserScopes " _
		& "WHERE user_id = " & v_intUserID

	Set oRS = Server.CreateObject("ADODB.RecordSet")
	oRS.CursorLocation = adUseClient
	oRS.Open strQuery, oConn, adOpenStatic, adLockBatchOptimistic, adCmdText
	For x = 0 to UBound(v_arScopes, 2)
		oRS.Filter = "scope_id = " & v_arScopes(g_SCOPE_ID, x)
		If oRS.EOF Then
			' didn't find match so add new record
			oRS.AddNew
			oRS.Fields("user_id") = v_intUserID
			oRS.Fields("scope_id") = v_arScopes(g_SCOPE_ID, x)
		End If
		oRS.Fields("visible") = v_arScopes(g_VISIBLE, x)
		oRS.Filter = adFilterNone
	Next
	oRS.UpdateBatch : oRS.Close : Set oRS = nothing
	
	if v_bGridInfo then
		' update segment settings
		if Session(g_unique & "Weekends") then
			bWeekend = 1
		else
			bWeekend = 0
		end if
		strQuery = "UPDATE tblUsers SET" _
			& " seg_size = " & Session(g_unique & "Segments")(g_SEG_MINS) _
			& ", seg_start = " & Session(g_unique & "Segments")(g_SEG_START) _
			& ", seg_end = " & Session(g_unique & "Segments")(g_SEG_END) _
			& ", show_weekend = " & bWeekend _
			& " WHERE (user_id) = " & v_intUserID
		oConn.Execute strQuery,,adCmdText + adExecuteNoRecords
	end if
	
	saveUserSettings = CheckForErrors(oConn, v_strView, v_strDate)
End Function

' update cookies with new options (updated 3/3/01)
' updates cookies --------------------------------------------------------
Sub saveGuestSettings()


End Sub

' return users default e-mail preference to use with new groups (updated 3/3/01)
' returns integer --------------------------------------------------------
Private Function getEmailDefault(ByRef r_oConn, ByVal v_intUserID, ByVal v_bEmail)
	dim oRS
	dim strQuery
	If v_bEmail = "" Then
		strQuery = "SELECT send_email FROM tblUsers WHERE (user_id) = " & v_intUserID
		Set oRS = Server.CreateObject("ADODB.RecordSet")
		oRS.Open strQuery, r_oConn, adOpenForwardOnly, adLockReadOnly, adCmdText
		v_bEmail = oRS("send_email") : oRS.Close : Set oRS = nothing	
	End If
	getEmailDefault = v_bEmail
End Function

' handle any errors resulting from db updates (updated 3/3/01)
' returns boolean --------------------------------------------------------
Private Function CheckForErrors(ByRef r_oConn, ByVal v_strView, ByVal v_strDate)
	dim strMessage, strError, bStatus, x
	
	If r_oConn.Errors.Count = 0 AND Err.Number = 0 Then
		r_oConn.CommitTrans : bStatus = true
		r_oConn.Close : Set r_oConn = nothing
	Else
		r_oConn.RollbackTrans : bStatus = false
		strMessage = "An error was encountered while updating preferences"
		
		If r_oConn.Errors.Count > 0 Then
			strMessage = strMessage & "\n\n"
			For x = 0 to r_oConn.Errors.Count - 1
				strError = r_oConn.Errors(x).Description
				strMessage = strMessage & strError & "\n"
			Next
			strMessage = Left(strMessage, Len(strMessage) - 2)
		End If
		If Err.Number <> 0 and Err.Description <> strError Then
			' this will only return the most recent error
			strMessage = strMessage & "\n\n" & Err.Source & " " & Err.Number _
				& "\n  " & Err.Description
		End If
		r_oConn.Close : Set r_oConn = nothing
		Call Redirect("", strMessage)
	End If
	CheckForErrors = bStatus
End Function
%>