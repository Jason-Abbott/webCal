<%
' Copyright 1996-2004 Jason Abbott (webcal@webott.com)

Class wcLayout

	'-------------------------------------------------------------------------
	'	Name: 		writeButtons()
	'	Purpose: 	write standard calendar buttons
	' Modifications:
	'	Date:		Name:	Description:
	'	3/14/01		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub writeButtons(ByVal v_sTitle, ByVal v_sView, ByVal v_aDates)
		dim x
	
		with response
			.write "<table width='100%'><tr><td>"
			.write v_sTitle
			.write "</td><form method='post' action='"
			.write g_sFILE_PREFIX
			.write v_sView
			.write ".asp'><td align='right' valign='bottom'><nobr><a href='"
			.write g_sFILE_PREFIX
			.write v_sView
			.write ".asp?date="
			.write v_aDates(g_PREV_DATE)
			.write "' "
			Call writeIcon("Prev","calprev", "Previous " & v_sView, 15, 16)
			.write "<a href='"
			.write g_sFILE_PREFIX
			.write "find.asp?view="
			.write v_sView
			.write "' "
			Call writeIcon("Search","search", "Find a scheduled event", 17, 16)
			.write "&nbsp;<a href='"
			.write g_sFILE_PREFIX
			.write v_sView
			.write ".asp?date="
			.write v_aDates(g_NEXT_DATE)
			.write "' "
			Call writeIcon("Next", "calnext", "Next " & v_sView, 17, 16)
			
			if Session(g_sDB_NAME & "UserID") = 0 Or Session(g_sDB_NAME & "UserID") = "" then
				.write "<a href='"
				.write g_sFILE_PREFIX
				.write "login.asp?url="
				.write Request.ServerVariables("URL")
				.write "?"
				.write Server.URLEncode(Request.ServerVariables("QUERY_STRING"))
				.write "' "
				Call writeIcon("Key", "key", "Login", 16, 15)
			else
				.write "<a href='"
				.write g_sFILE_PREFIX
				.write "admin.asp?view="
				.write v_sView
				.write "' "
				Call writeIcon("Users", "users", "Manage Users and Groups", 12, 15)
			end if

			.write "<a href='"
			.write g_sFILE_PREFIX
			.write v_sView
			.write "-print.asp?date="
			.write v_aDates(g_THIS_DATE)
			.write "' target='_top' "
			Call writeIcon("Print", "print", "Make printable", 16, 14)
			.write "<a href='javascript:document.forms[0].submit();' "
			Call writeIcon("Goto", "goto", "Goto the selected date", 18, 15)
			.write "<select name='month' class='flat'>"
			for x = 1 to 12
				.write "<option value='"
				.write x
				.write "'"
				if x = Month(Date) then .write " selected"
				.write ">"
				.write MonthName(x, 1)
			next
			.write "</select><select name='year'>"
			for x = Year(Date) - 10 to Year(Date) + 10
				.write "<option"
				if x = Year(Now) then .write " selected"
				.write ">"
				.write x
			next
			.write "</select><a href='#' onClick=""javascript:showHelp('month');"" "
			Call writeIcon("Help", "help", "Display help", 8, 10)
			.write "</nobr></td></form></table>"
		end with
	End Sub

	'-------------------------------------------------------------------------
	'	Name: 		writeIcon()
	'	Purpose: 	write image icon
	' Modifications:
	'	Date:		Name:	Description:
	'	9/23/03		JEA		Creation
	'-------------------------------------------------------------------------
	Public Sub writeIcon(ByVal v_sName, ByVal v_sSource, ByVal v_sMsg, ByVal v_lWidth, ByVal v_lHeight)
		with response
			Call writeRolloverJS(v_sName, v_sName, v_sMsg)
			.write "><img name='"
			.write v_sName
			.write "' src='./images/icon_"
			.write v_sSource
			.write "_grey.gif' "
			.write "width='"
			.write v_lWidth
			.write "' height='"
			.write v_lHeight
			.write "' alt='"
			.write v_sMsg
			.write "' border='0' class='Icon'></a>"
		end with
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		writeRollover()
	'	Purpose: 	write javascript for mouse rollover
	'				depends on inclusion of /script/wc_functions.js
	' Modifications:
	'	Date:		Name:	Description:
	'	3/2/01		JEA		Creation
	'	9/23/03		JEA		Use .writes
	'-------------------------------------------------------------------------
	Public Sub writeRolloverJS(ByVal v_sName, ByVal v_sSource, ByVal v_sMsg)
		if IsVoid(v_sSource) then v_sSource = v_sName
		with response
			.write "onMouseOver=""iconOver('"
			.write v_sName
			.write "','"
			.write v_sSource
			.write "','"
			.write v_sMsg
			.write "'); return true;"" onMouseOut=""iconOut('"
			.write v_sName
			.write "','"
			.write v_sSource
			.write "'); return true;"""
		end with
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		writeStatusJS()
	'	Purpose: 	write javascript to update status bar with message
	' Modifications:
	'	Date:		Name:	Description:
	'	3/2/01		JEA		Creation
	'	9/23/03		JEA		Use .writes
	'-------------------------------------------------------------------------
	Public Sub writeStatusJS(ByVal v_sStatus)
		with response
			.write "onMouseOver=""status='"
			.write Replace(v_sStatus, "'", "\'")
			.write "'; return true;"" onMouseOut=""status=''; return true;"""
		end with
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		writeButton()
	'	Purpose: 	write button
	' Modifications:
	'	Date:		Name:	Description:
	'	3/2/01		JEA		Creation
	'	9/23/03		JEA		Use .writes
	'-------------------------------------------------------------------------
	Public Sub writeButton(ByVal v_sName, ByVal v_sLink, ByVal v_lHeight, ByVal v_lWidth)
		with response
			select case Session(g_sDB_NAME & "Browser")(g_BROWSER_ID)
				case g_BROWSER_IE
					.write "<div class='button' style='height: "
					.write v_lHeight
					.write "px; width: "
					.write v_lWidth
					.write "px;' onMouseOver=""this.className='buttonOn';"" "
					.write "onMouseOut=""this.className='button';"" onClick="""
					.write v_sLink
					.write ";"">"
					.write v_sName
					.write "</div>"
				case else
					if InStr(v_sLink, "javascript:") then
						v_sLink = Right(v_sLink, Len(v_sLink) - 11)
					else
						v_sLink = "javascript:location.href=" & v_sLink
					end if
					.write "<input type='button' value='"
					.write v_strName
					.write "' onClick='"
					.write v_strLink
					.write "'>"	
			end select
		end with
	End Sub
	
	' show the requested symbol (updated 3/4/01)
	' returns string ---------------------------------------------------------
	
	'-------------------------------------------------------------------------
	'	Name: 		writeSymbol()
	'	Purpose: 	write symbolic character
	' Modifications:
	'	Date:		Name:	Description:
	'	3/4/01		JEA		Creation
	'	10/15/03	JEA		Use .writes
	'-------------------------------------------------------------------------
	Public Sub writeSymbol(ByVal v_lCharID, ByVal v_sSize)
		dim aSymbol
		
		aSymbol = Session(g_sDB_NAME & "Symbols")
		
		with response
			.write "<span style='font: "
			.write v_sSize
			.write " "
			.write aSymbol(g_FONT_FACE, v_lCharID)
			.write "; color: #aa0000; float: right;'>"
			.write Chr(aSymbol(g_FONT_CHAR, v_lCharID))
			.write "</span>"
		end with
	End Sub
End Class
%>