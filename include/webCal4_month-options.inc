<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' Last updated 3/3/2001

dim m_arGroups
dim m_strGroupList
dim m_arScopes
dim m_strScopeList

' get arrays of groups and scopes available to current user
m_arGroups = Session(g_unique & "Groups")
m_arScopes = Session(g_unique & "Scopes")

for x = 0 to UBound(m_arGroups, 2)
	' go through each group visible to the current user
	'if m_arGroups(g_GROUP_ACCESS, x) then
		m_strGroupList = m_strGroupList & "<option value='" _
			& m_arGroups(g_GROUP_ID, x) & "'"
		if m_arGroups(g_VISIBLE, x) then
			m_strGroupList = m_strGroupList & " selected"
		end if
		m_strGroupList = m_strGroupList & ">" & m_arGroups(g_GROUP_NAME, x)
	'end if
next

for x = UBound(m_arScopes, 2) to 0 step -1
	m_strScopeList = m_strScopeList & "<input type='checkbox' name='fldScope" & x & "'"
	if m_arScopes(g_VISIBLE, x) then
		m_strScopeList = m_strScopeList & " checked"
	end if
	m_strScopeList = m_strScopeList & ">" & m_arScopes(g_SCOPE_NAME, x) & "<br>"
next
' remove extra "<br>"
m_strScopeList = Left(m_strScopeList, Len(m_strScopeList) - 4)
%>
<div id="divSettings">
<table cellspacing='0' cellpadding='1' border='0' class='tblSettings'>
<form name="frmOptions" action="webCal4_prep.asp" method="post">
<tr>
	<td class='titleBar'><%=g_sTITLE_SHOW_EVENTS%></td>
	<td align='right' class='titleBar'>
		<a href="javascript:hideSettings();" class='SymbolButton'>
		<%=showSymbol(g_CHAR_CLOSE, 12)%></a>
	</td>
<tr>
	<td valign='top' align='center'><b>From</b><br>
		<select name="fldGroups" multiple size=3><%=m_strGroupList%></select>
	</td>
	<td valign='top' align='left'>
		<center><b><%=g_sFRM_THAT_ARE%></b></center><%=m_strScopeList%>
	</td>
<tr>
	<td><input type="checkbox" name="fldDefault"><%=g_sFRM_MAKE_DEFAULT%></td>
	<td align='center'>
	<%=makeButton(g_sBTN_SAVE,"document.frmOptions.submit();",12,60)%>
	</td>
	
<input type="hidden" name="fldTotalGroups" value="<%=UBound(m_arGroups, 2) + 1%>">
<input type="hidden" name="fldView" value="<%=m_strView%>">
<input type="hidden" name="fldDate" value="<%=m_arDates(g_THIS_DATE)%>">
</form>
</table>
</div>
