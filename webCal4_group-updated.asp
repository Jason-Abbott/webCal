<% Option Explicit %>
<% Response.Buffer = true %>
<!--#include file="./data/webCal4_data.inc"-->
<!--#include file="./include/webCal4_constants.inc"-->
<!--#include file="./include/webCal4_functions.inc"-->
<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' Last updated 2/27/2001

dim m_strMessage	' feedback
dim m_intGroupID	' group id
dim bOverlap		' allow overlapping events
dim m_strQuery		' query passed to db
dim m_strUsers		' javascript generated list of user permissions
dim m_strGroupName	' group name
dim m_arGroupUsers	' array of users in this group
dim m_arGroups		' array of groups
dim m_oConn			' connection object
dim m_oRS			' recordset object
dim x				' loop counter

Const m_GROUP_USER_ID = 0
Const m_GROUP_USER_ACCESS = 1
Const m_GROUP_DEFAULT_ACCESS = 2
Const m_GROUP_USER_NEW = 3

On Error Resume Next

' format (user id|access level|default access|newness,[repeat])
m_strUsers = Trim(Request.Form("fldAccessList"))
m_strGroupName = Request.Form("fldGroupName")
m_arGroups = Session(g_unique & "Groups")
Set m_oRS = Server.CreateObject("ADODB.Recordset")
Set m_oConn = Server.CreateObject("ADODB.Connection")
m_oConn.Open g_strDSN : m_oConn.BeginTrans

if m_strUsers <> "" then m_arGroupUsers = ListToArray(m_strUsers, ",", "|")

if Request.Form("fldOverlap") = "on" then
	bOverlap = 1
else
	bOverlap = 0
end if

if Request.Form("fldEditType") = "update" then
	' update permission settings for modified users
	m_strMessage = "updated"
	m_intGroupID = Request.Form("fldGroupID")
	if m_strUsers <> "" then
		' retrieve recordset of users for this group
		m_strQuery = "SELECT * FROM tblPermissions WHERE group_id=" & m_intGroupID
		m_oRS.CursorLocation = adUseClient	' allows batch updates
		m_oRS.Open m_strQuery, m_oConn, adOpenStatic, adLockBatchOptimistic, adCmdText
		for x = 0 to UBound(m_arGroupUsers, 2)
			if CInt(m_arGroupUsers(m_GROUP_USER_NEW, x)) = 1 then
				' give user access to this group
				m_oRS.AddNew
				m_oRS.Fields("group_id") = m_intGroupID
				m_oRS.Fields("user_id") = m_arGroupUsers(m_GROUP_USER_ID, x)
				m_oRS.Fields("access_level") = m_arGroupUsers(m_GROUP_USER_ACCESS, x)
			else
				' set cursor on the record for this user
				m_oRS.Filter = "user_id = " & m_arGroupUsers(m_GROUP_USER_ID, x)
				if deletePermissions(CInt(m_arGroupUsers(m_GROUP_USER_ACCESS, x)), _
					CInt(m_arGroupUsers(m_GROUP_DEFAULT_ACCESS, x))) then
					' delete group entry for this user
					m_oRS.Delete
				else
					' update user permissions
					m_oRS.Fields("access_level") = m_arGroupUsers(m_GROUP_USER_ACCESS, x)
				end if
				oRS.Filter = adFilterNone
			end if
		next
		m_oRS.UpdateBatch : m_oRS.Close
	end if
	
	' update existing group
	m_strQuery = "UPDATE tblGroups SET " _
		& "group_name = '" & m_strGroupName & "', " _
		& "group_description = '" & Request.Form("fldGroupDescription") & "', " _
		& "allow_overlap = " & bOverlap & " " _
		& "WHERE (group_id)=" & m_intGroupID
	m_oConn.Execute m_strQuery,,adCmdText + adExecuteNoRecords
	
	' update group session array
	for x = 0 to UBound(m_arGroups)
		if m_arGroups(g_GROUP_ID, x) = CInt(m_intGroupID) then
			' the name is the only possible change
			m_arGroups(g_GROUP_NAME, x) = m_strGroupName
		end if
	next
else
	' add new group
	m_strMessage = "added"

	' use ADO methods so we can immediately retrieve new group id
	m_oRS.Open "tblGroups", m_oConn, adOpenStatic, adLockOptimistic, adCmdTable
	m_oRS.AddNew
	m_oRS.Fields("group_name") = m_strGroupName
	m_oRS.Fields("group_description") = Request.Form("fldGroupDescription")
	m_oRS.Fields("allow_overlap") = bOverlap
	m_oRS.Update
	m_intGroupID = m_oRS("group_id")
	m_oRS.Close

	' insert user permissions for the new group
	m_oRS.CursorLocation = adUseClient	' allows batch updates
	m_oRS.Open "tblPermissions", m_oConn, adOpenStatic, adLockBatchOptimistic, adCmdTable
	if m_strUsers <> "" then
		for x = 0 to UBound(m_arGroupUsers, 2)
'			if CInt(m_arUser(m_GROUP_USER_ACCESS)) <> g_NO_ACCESS and _
'			   CInt(m_arUser(m_GROUP_USER_ACCESS)) <> CInt(m_arUser(m_GROUP_DEFAULT_ACCESS)) then
			   
				m_oRS.AddNew
				m_oRS.Fields("group_id") = m_intGroupID
				m_oRS.Fields("user_id") = m_arGroupUsers(m_GROUP_USER_ID, x)
				m_oRS.Fields("access_level") = m_arGroupUsers(m_GROUP_USER_ACCESS, x)
'			end if
		next
	end if
	m_oRS.UpdateBatch
	m_oRS.Close
	
	' add the new group to the group session array
	' only admin should be here so add admin access
	x = UBound(m_arGroups, 2) + 1
	ReDim Preserve m_arGroups(UBound(m_arGroups), x)	
	m_arGroups(g_GROUP_ID, x) = m_intGroupID
	m_arGroups(g_GROUP_NAME, x) = m_strGroupName
	m_arGroups(g_VISIBLE, x) = 1
	m_arGroups(g_GROUP_ACCESS, x) = g_ADMIN_ACCESS
end if

Session(g_unique & "Groups") = m_arGroups

Call HandleErrors(m_oConn, m_strMessage, "group-edit", "The group " & m_strGroupName, _
	"saving group information", Request.Form("fldView"), Request.Form("fldStartDate"))
	
' process business rules for deleting permissions row (updated 2/27/01)
' returns boolean --------------------------------------------------------
Function deletePermissions(ByVal v_lAccess, ByVal v_lDefaultAccess)
	dim bDelete
	if v_lAccess = g_NO_ACCESS and v_lDefaultAccess <> g_NO_ACCESS then
		' revert to default when group-specific access removed
		bDelete = true
	elseif v_lAccess = v_lDefaultAccess then
		' revert to default when same as group-specific access
		bDelete = true
	else
		bDelete = false
	end if
	deletePermissions = bDelete
End Function
%>