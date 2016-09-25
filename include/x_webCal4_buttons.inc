<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' Last updated 4/1/2000

dim c		' loop counter
%>

<tr>
	<td><%=m_strTitle%></td>
	<form method="post" action="webCal4_<%=m_strView%>.asp">
	<td align="right" valign="bottom"><nobr>

		<a href="webCal4_<%=m_strView%>.asp?date=<%=m_arDates(g_PREV_DATE)%>"
		<%=switchIcon("Prev","", "Previous " & m_strView)%>><img name="Prev" src="./images/icon_calprev_grey.gif"
		 width=15 height=16 alt="Previous <%=m_strView%>" border=0></a>
		&nbsp;
		<a href="webCal4_find.asp?view=<%=m_strView%>"
		<%=switchIcon("Search","", "Find a scheduled event")%>><img name="Search" src="./images/icon_search_grey.gif"
		 width=17 height=16 alt="Find events" border=0></a>
		&nbsp;
		<a href="webCal4_<%=m_strView%>.asp?date=<%=m_arDates(g_NEXT_DATE)%>"
		<%=switchIcon("Next","", "Next " & m_strView)%>><img name="Next" src="./images/icon_calnext_grey.gif"
		 width=15 height=16 alt="Next <%=m_strView%>" border=0></a>
		&nbsp;
<%
if Session(g_sDB_NAME & "UserID") = 0 or Session(g_sDB_NAME & "UserID") = "" then
	response.write "<a href=""webCal4_login.asp?url=" _
		& Request.ServerVariables("URL") & "?" _
		& Server.URLEncode(Request.ServerVariables("QUERY_STRING")) & """"
%>
		<%=switchIcon("Key","","Login")%>><img name="Key" src="./images/icon_key_grey.gif"
		 width=16 height=15 alt="Login" border=0></a>
		&nbsp;
<% else %>
		<a href="webCal4_admin.asp?view=<%=m_strView%>"
		<%=switchIcon("Users","","Manage Users and Groups")%>><img name="Users" src="./images/icon_users_grey.gif"
		 width=12 height=15 alt="Manage Users and Groups" border=0></a>
		&nbsp;
<% end if%>
		<a href="webCal4_<%=m_strView%>-print.asp?date=<%=m_arDates(g_THIS_DATE)%>" target="_top"
		<%=switchIcon("Print","", "Make printable")%>><img name="Print" src="./images/icon_print_grey.gif"
		width=16 height=14 border=0 alt="Make printable"></a>
		&nbsp;
		<a href="javascript:document.forms[0].submit();" 
		<%=switchIcon("Goto","", "Goto the selected date")%>><img name="Goto" src="./images/icon_goto_grey.gif"
		 width=18 height=15 alt="Goto the selected date" border=0></a>
		
		<select name="month" class="flat">
<%
' this creates the form list of month names

for c = 1 to 12
	response.write "<option value='" & c & "'"
	if c = Month(Date) then response.write " selected"
	response.write ">" & MonthName(c,1) & VbCrLf
next
%>
		</select>
		<select name="year">
<%
' this creates the form list of 20 years

for c = Year(Date) - 10 to Year(Date) + 10
	response.write "<option"
	if c = Year(Now) then response.write " selected"
	response.write ">" & c & VbCrLf
next
%>
		</select>
		
		<a href="#" onClick="javascript:showHelp('month');"
		<%=switchIcon("Help","", "Display help")%>><img name="Help" src="./images/icon_help_grey.gif"
		width=8 height=10 border=0 alt="Display help"></a>
		
	</nobr>
	</td>
	</form>	