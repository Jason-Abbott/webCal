<% Option Explicit %>
<% Response.Buffer = true %>
<!--#include file="./data/webCal4_data.inc"-->
<!--#include file="./include/webCal4_constants.inc"-->
<!--#include file="./include/webCal4_functions.inc"-->
<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' Last updated 2/25/2001

dim m_strQuery		' query passed to db
dim m_strMessage	' data Errors to pass back
dim m_intUserID		' user id
dim m_strGroups		' javascript generated string
dim m_arGroups		' array of all group permissions
dim m_arUserScopes	' array of user scope ids
dim m_oConn			' connection object
dim m_oRS			' recordset object
dim x				' loop counter

Const c_UserGroupID = 0
Const c_UserGroupAccess = 1
Const c_UserGroupNew = 2

On Error Resume Next

' format (group id|access level|newness,[repeat])
m_strGroups = Trim(Request.Form("fldAccessList"))
Set m_oRS = Server.CreateObject("ADODB.Recordset")
Set m_oConn = Server.CreateObject("ADODB.Connection") : m_oConn.Open g_strDSN
m_oConn.BeginTrans

if m_strGroups <> "" then m_arGroups = ListToArray(m_strGroups, ",", "|")

if Request.Form("fldEditType") = "update" then
	' update permission settings for modified groups
	m_strMessage = "updated"
	m_intUserID = Request.Form("fldUserID")
	if m_strGroups <> "" then
		' retrieve recordset of groups for this user
		m_strQuery = "SELECT * FROM tblPermissions WHERE user_id=" & m_intUserID
		m_oRS.CursorLocation = adUseClient	' allows batch updates
		m_oRS.Open m_strQuery, m_oConn, adOpenStatic, adLockBatchOptimistic, adCmdText
		for x = 0 to UBound(m_arGroups, 2)
			if m_arGroup(c_UserGroupNew) then
				' add user to this new group
				m_oRS.AddNew
				m_oRS.Fields("group_id") = m_arGroups(c_UserGroupID, x)
				m_oRS.Fields("user_id") = m_intUserID
				m_oRS.Fields("access_level") = m_arGroups(c_UserGroupAccess, x)
			else
				' set cursor on the record for this group
				m_oRS.Filter = "group_id = " & m_arGroups(c_UserGroupID, x)
				if m_arGroups(c_UserGroupAccess, x) = g_NO_ACCESS then
					' remove user from groups to which they have no access
					m_oRS.Delete
				else
					' update group permissions
					m_oRS.Fields("access_level") = m_arGroups(c_UserGroupAccess, x)
				end if
			end if
		next
		m_oRS.UpdateBatch
		m_oRS.Close
	end if

	' update existing user
	m_strQuery = "UPDATE tblUsers SET " _
		& "name_first = '" & Request.Form("fldNameFirst") & "', " _
		& "name_last = '" & Request.Form("fldNameLast") & "', " _
		& "user_email = '" & Request.Form("fldEmail") & "', " _
		& "user_login = '" & Request.Form("fldLogin") & "', " _
		& "user_password = '" & Request.Form("fldPassword") & "', " _
		& "default_access = " & Request.Form("fldDefault") & ", " _
		& "user_lcid = '" & Request.Form("fldLCID") & "' " _
		& "WHERE (user_id)=" & m_intUserID
	m_oConn.Execute m_strQuery,,adCmdText + adExecuteNoRecords
else
	' add new user
	m_strMessage = "added"

	' use ADO methods so we can immediately retrieve new user id
	' ADD month / day segments, start page
	m_oRS.Open "tblUsers", m_oConn, adOpenStatic, adLockOptimistic, adCmdTable
	m_oRS.AddNew
	m_oRS.Fields("name_first") = Request.Form("fldNameFirst")
	m_oRS.Fields("name_last") = Request.Form("fldNameLast")
	m_oRS.Fields("user_email") = Request.Form("fldEmail")
	m_oRS.Fields("user_login") = Request.Form("fldLogin")
	m_oRS.Fields("user_password") = Request.Form("fldPassword")
	m_oRS.Fields("default_access") = Request.Form("fldDefault")
	m_oRS.Fields("user_lcid") = Request.Form("fldLCID")
	m_oRS.Update
	m_intUserID = m_oRS("user_id")
	m_oRS.Close

	' insert group permissions for the new user
	if m_strGroups <> "" then
		m_oRS.CursorLocation = adUseClient	' allows batch updates
		m_oRS.Open "tblPermissions", m_oConn, adOpenStatic, adLockBatchOptimistic, adCmdTable
		for x = 0 to UBound(m_arGroups, 2)
			if m_arGroups(c_UserGroupAccess, x) <> g_NO_ACCESS then
				m_oRS.AddNew
				m_oRS.Fields("group_id") = m_arGroups(c_UserGroupID, x)
				m_oRS.Fields("user_id") = m_intUserID
				m_oRS.Fields("access_level") = m_arGroups(c_UserGroupAccess, x)
			end if
		next
		m_oRS.UpdateBatch
		m_oRS.Close
	end if

	' set default user scope visibility
	m_strQuery = "SELECT scope_id FROM tblScopes"
	m_arUserScopes = getRowArray(m_strQuery, m_oConn)
	m_oRS.CursorLocation = adUseClient
	m_oRS.Open "tblUserScopes", m_oConn, adOpenStatic, adLockBatchOptimistic, adCmdTable
	for x = 0 to UBound(m_arUserScopes)
		m_oRS.AddNew
		m_oRS.Fields("user_id") = m_intUserID
		m_oRS.Fields("scope_id") = m_arUserScopes(x)
		m_oRS.Fields("visible") = 1
	next
	m_oRS.UpdateBatch
	m_oRS.Close
end if
Set m_oRS = nothing

Call HandleErrors(m_oConn, m_strMessage, "user-edit", "The user " _
	& Request.Form("fldNameFirst") & " " & Request.Form("fldNameLast"), _
	"saving user information", Request.Form("fldView"), Request.Form("fldStartDate"))
%>