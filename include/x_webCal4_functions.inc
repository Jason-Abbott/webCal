<%

Function makeButton(ByVal v_strName, ByVal v_strLink, _
	ByVal v_intHeight, ByVal v_intWidth)
	
	dim strHTML
	select case Session(g_unique & "Browser")
		case "ie"
			strHTML = "<div class='button' style='height: " & v_intHeight _
				& "px; width: " & v_intWidth & "px;' " _
				& "onMouseOver=""this.className='buttonOn';"" " _
				& "onMouseOut=""this.className='button';"" " _
				& "onClick=""" & v_strLink & ";"">" & v_strName & "</div>"
		case else
			if InStr(v_strLink, "javascript:") then
				v_strLink = Right(v_strLink, Len(v_strLink) - 11)
			else
				v_strLink = "javascript:location.href=" & v_strLink
			end if
			strHTML = "<input type='button' value='" & v_strName _
				& "' onClick='" & v_strLink & "'>"	
	end select
	makeButton = strHTML
End Function

' show the requested symbol (updated 3/4/01)
' returns string ---------------------------------------------------------
Function showSymbol(ByVal v_intCharID, ByVal v_intSize)
	dim strHTML
	dim arSymbol
	arSymbol = Session(g_unique & "Symbols")
	select case Session(g_unique & "Browser")
		case "ie"
			strHTML = "<font style='font: " & v_intSize & "pt " _
				& arSymbol(g_FONT_FACE, v_intCharID) & "'>"
		case else
			' convert point size to approximate HTML size
			v_intSize = Round((v_intSize - 2) / 2)
			strHTML = "<font size='" & v_intSize & "' " _
				& "face='" & arSymbol(g_FONT_FACE, v_intCharID) & "'>"
	end select
	strHTML = strHTML & Chr(arSymbol(g_FONT_CHAR, v_intCharID)) & "</font>"
	showSymbol = strHTML
End Function

' show load time from cache or database (updated 3/3/01)
' returns string ---------------------------------------------------------
Function showLoadTime(ByVal v_strQuery, ByVal v_strLoadFrom)
	if v_strQuery <> g_sNO_EVENTS then
		v_strLoadFrom = "<font color='#" & g_arColor(5) & "'>" _
			& g_sMSG_LOAD_FROM & " <b>" & v_strLoadFrom & "</b> in " _
			& showElapsedTime(milliTime() - m_intStartTime) & "</font><br>"
	else
		v_strLoadFrom = ""
	end if
	showLoadTime = v_strLoadFrom
End Function

' write javascript to update status bar with message (updated 3/2/01)
' returns string ---------------------------------------------------------
function showStatus(ByVal v_strMsg)
	v_strMsg = Replace(v_strMsg, "'", "\'")
	showStatus = "onMouseOver=""status='" & v_strMsg & "'; " _
		& "return true;"" " & VbCrLf _
		& "onMouseOut=""status=''; return true;"""
end function

' write javascript for mouse rollover (updated 3/2/01)
'   depends on inclusion of /script/webCal4_functions.js
' returns string ---------------------------------------------------------
function switchIcon(ByVal v_strName, ByVal v_strSource, ByVal v_strMsg)
	if v_strSource = "" then v_strSource = v_strName
	switchIcon = "onMouseOver=""" _
		& "iconOver('" & v_strName & "','" & v_strSource & "','" & v_strMsg & "'); return true;"" " _
		& "onMouseOut=""" _
		& "iconOut('" & v_strName & "','" & v_strSource & "'); return true;"""
end function

' handle any errors that may have occured on page (updated 3/1/01)
' redirects or raises error message --------------------------------------
Sub HandleErrors(ByRef r_oConn, ByVal v_strMessage, ByVal v_strPage, _
	ByVal v_strMsgSuccess, ByVal v_strMsgFail, ByVal v_strView, ByVal v_strDate)
	
	dim strError
	dim strPage
	dim x
	
	if r_oConn.Errors.Count = 0 AND Err.Number = 0 then
		r_oConn.CommitTrans
		r_oConn.Close : Set r_oConn = nothing
		v_strMessage = v_strMsgSuccess & " was successfully " & v_strMessage

		if Request.QueryString("again") then
			' send back to edit page
			strPage = "webCal4_" & v_strPage & ".asp?date=" _
				& v_strDate	& "&view=" & v_strView
		else
			if v_strPage = "event-edit" then
				' send back to calendar
				strPage = "webCal4_" & v_strView & ".asp?date=" & v_strDate
			else
				' send to admin page
				strPage = "webCal4_admin.asp?view=" & v_strView
			end if
		end if
	else
		r_oConn.RollbackTrans
		v_strPage = ""
		v_strMessage = "An error was encountered while " & v_strMsgFail
		
		if r_oConn.Errors.Count > 0 then
			v_strMessage = v_strMessage & "\n\n"
			for x = 0 to r_oConn.Errors.Count - 1
				strError = r_oConn.Errors(x).Description
				v_strMessage = v_strMessage & strError & "\n"
			next
			v_strMessage = Left(v_strMessage, Len(v_strMessage) - 2)
		end if
		if Err.Number <> 0 and Err.Description <> strError then
			' this will only return the most recent error
			v_strMessage = v_strMessage & "\n\n" & Err.Source & " " & Err.Number _
				& "\n  " & Err.Description
		end if
		r_oConn.Close : Set r_oConn = nothing
	end if
	Call Redirect(strPage, v_strMessage)
End Sub

' write HTML to call JS to save message and redirect (updated 3/1/01)
'   cookies written with VBScript can't be written in JavaScript
' writes string ----------------------------------------------------------
Sub Redirect(ByVal v_strPage, ByVal v_strMessage)
	v_strMessage = Replace(v_strMessage, "'", "\'")
	response.write "<html><head><script language='javascript' " _
		& "src='./script/webCal4_functions.js'></script></head>" _
		& "<body onLoad=""goPageMessage('" & v_strMessage _
		& "','" & v_strPage & "');"" bgcolor='#" & g_arColor(1) _
		& "'></body></html>"
End Sub

' convert a string list to an array (updated 2/27/01)
' returns array ----------------------------------------------------------
Function ListToArray(ByVal v_strList, ByVal v_strSplit1, ByVal v_strSplit2)
	dim arData
	dim arData1		' temporary 1D array per row
	dim arData2()	' temporary 2D array
	dim intBound2	' upper bound of second dimension
	dim x, y		' loop counters
	
	arData = Split(v_strList, v_strSplit1)
	if v_strSplit2 <> "" then
		' two dimensional array--assumes rectangular data set
		intBound2 = UBound(Split(arData(0), v_strSplit2))
		' reverse dimension order to match rs.getrows
		ReDim arData2(intBound2, UBound(arData))
		for x = 0 to UBound(arData)
			arData1 = Split(arData(x), v_strSplit2)
			for y = 0 to UBound(arData1)
				arData2(y,x) = arData1(y)			
			next
		next
		arData = arData2 : Erase arData1 : Erase arData2
	end if
	ListToArray = arData
End Function

' none of FormatDateTime() options are quite right (updated 8/21/00)
' returns string ---------------------------------------------------------
Function formatDate(ByVal v_sDate)
	if IsDate(v_sDate) then
		' don't try to format if it's not a date
		formatDate = MonthName(Month(v_sDate)) & " " & Day(v_sDate) & ", " & Year(v_sDate)
	else
		formatDate = v_sDate
	end if
End Function

' build LCID list for drop-down (updated 2/25/01)
' returns string ---------------------------------------------------------
Function makeLCIDList(ByRef r_oConn)
	dim strQuery
	strQuery = "SELECT lcid, lcid_name FROM tblLCIDs ORDER BY lcid_name"
	makeLCIDList = makeList(strQuery, Session.LCID, r_oConn)
End Function

' build list from recordset (updated 12/4/00)
'   rs must have only id and description fields (two fields)
' returns string ---------------------------------------------------------
Function makeList(ByVal v_strQuery, ByVal v_strSelect, ByRef r_oConn)
	dim bNewObject
	dim strHtml
	dim oRS
	
	bNewObject = newConnection(r_oConn)
	Set oRS = Server.CreateObject("ADODB.Recordset")
	oRS.Open v_strQuery, r_oConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	if not oRS.EOF then
		strHTML = oRS.GetString(2,,"'>",vbCrLf & "<option value='")
		strHTML = "<option value='" & Left(strHTML, Len(strHTML) - 15)
		makeList = makeSelected(strHTML, v_strSelect)
	else
		makeList = ""
	end if
	oRS.Close : Set oRS = nothing
	if bNewObject then Set r_oConn = nothing
End Function

' selects item in HTML option list (updated 8/30/00)
' returns string ---------------------------------------------------------
Function makeSelected(ByVal v_strList, ByVal v_strSelect)
	' assumes <option value='[v_strSelect]'> with (') delimiters
	if v_strSelect <> "" then
		makeSelected = Replace(v_strList, "'" & v_strSelect & "'", "'" _
			& v_strSelect & "' selected", 1, 1)
	else
		makeSelected = v_strList
	end if
End Function
%>