<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' Last updated 3/8/2001

dim m_arGroups
dim m_strGroupList
dim m_arScopes
dim m_strScopeList
dim m_strSegMinsList	' option list for number of minutes per segment
dim m_strAMHourList
dim m_strPMHourList
dim m_strSelect
dim m_intHour

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

' make list for selecting minutes per segment
for each x in Array(5,15,30,60)
	m_strSegMinsList = m_strSegMinsList & "<option value='" & x & "'>" & x
next
m_strSegMinsList = makeSelected(m_strSegMinsList, m_arSegments(g_SEG_MINS))

' make hour lists
for x = 0 to 11
	m_intHour = x : if m_intHour = 0 then m_intHour = 12
	m_strAMHourList = m_strAMHourList & "<option value='" & m_intHour & "'>" & m_intHour & ":00"
	m_strPMHourList = m_strPMHourList & "<option value='" & m_intHour + 12 & "'>" & m_intHour & ":00"
next
m_strSelect = CStr(m_arSegments(g_SEG_START) / m_arSegments(m_SEG_PER_HOUR))
m_strAMHourList = makeSelected(m_strAMHourList, m_strSelect)
m_strSelect = (m_arSegments(g_SEG_END) + 1) / m_arSegments(m_SEG_PER_HOUR)
m_strPMHourList = makeSelected(m_strPMHourList,  m_strSelect)
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
<tr><td bgcolor="#<%=g_arColor(5)%>" height='1' colspan='2'></td>
<tr>
	
	<td bgcolor="#<%=g_arColor(3)%>" colspan='2' align='center'><font face="<%=g_arFont(0)%>" size=2>
		<input type="checkbox" name="fldWeekend"<%if Session(g_unique & "Weekends") then%> checked<%end if%>> Show weekends
		<br>
		<select name="fldSegMins"><%=m_strSegMinsList%></select> minute intervals from<br>
		<select name="fldHourStart"><%=m_strAMHourList%></select> AM to
		<select name="fldHourEnd"><%=m_strPMHourList%></select> PM</font>
	</td>	
<tr><td bgcolor="#<%=g_arColor(5)%>" height='1' colspan='2'></td>
<tr>
	<td><input type="checkbox" name="fldDefault"><%=g_sFRM_MAKE_DEFAULT%></td>
	<td align='center'>
	<%=makeButton(g_sBTN_SAVE,"javascript:document.frmOptions.submit();","update",15,60)%>
	</td>
<input type="hidden" name="fldView" value="<%=m_strView%>">
<input type="hidden" name="fldDate" value="<%=m_arDates(g_THIS_DATE)%>">
<input type="hidden" name="fldTotalGroups" value="<%=UBound(m_arGroups, 2) + 1%>">
</form>
</table>
</div>