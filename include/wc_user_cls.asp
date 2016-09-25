<%
' Copyright 1996-2004 Jason Abbott (webcal@webott.com)

Class wcUser

	Private Sub Class_Initialize()
	End Sub
	
	Private Sub Class_Terminate()
	End Sub

	'-------------------------------------------------------------------------
	'	Name: 		Login()
	'	Purpose: 	login user
	' Modifications:
	'	Date:		Name:	Description:
	'	9/24/03		JEA		Created
	'-------------------------------------------------------------------------
	Public Function Login(ByVal v_sUserName, ByVal v_sPassword)
		Const USER_ID = 0
		Const PASSWORD = 1
		Const LCID = 2
		Const SHOW_WEEKEND = 3
		Const WEEK_SEG_MINS = 4
		Const WEEK_SEG_START = 5
		Const WEEK_SEG_END = 6
		Const DAY_SEG_MINS = 7
		Const DAY_SEG_START = 8
		Const DAY_SEG_END = 9
		Const START_PAGE = 10
		
		dim aSegments
		dim oSession
		dim sResult
		dim sError
		dim aUser
		dim sQuery
		dim x
		
		x = 0
		sError = g_sMSG_BAD_LOGIN
		sResult = ""
		
		if v_sUserName <> "" then
			' login was attempted--validate
			aUser = getUser(v_sUserName)
			if Not IsArray(aUser) then
				' login not found
				sResult = sError
			else
				if aUser(PASSWORD,x) = v_sPassword then
					' Array of arrays--dummy array of 0s for month values
					aSegments = Array( _
						Array(aUser(WEEK_SEG_MINS,x), aUser(WEEK_SEG_START,x), aUser(WEEK_SEG_END,x)), _
						Array(aUser(DAY_SEG_MINS,x), aUser(DAY_SEG_START,x), aUser(DAY_SEG_END,x)), _
						Array(0,0,0))

					Set oSession = New wcSession
					Call oSession.initUser(aUser(USER_ID,x), CBool(aUser(SHOW_WEEKEND,x) <> 0), _
						aSegments, aUser(START_PAGE,x), aUser(LCID,x))
					Set oSession = Nothing
				else
					' password doesn't match
					sResult = sError
				end if
			end if
		end if
		Login = sResult
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		getUser()
	'	Purpose: 	get user info
	'	Return: 	array
	' Modifications:
	'	Date:		Name:	Description:
	'	9/24/03		JEA		Created
	'-------------------------------------------------------------------------
	Private Function getUser(ByVal v_sUserName)
		dim oData
		dim sQuery
		sQuery = "SELECT user_id, user_password, user_lcid, show_weekend, " _
			& "week_seg_mins, week_seg_start, week_seg_end, " _
			& "day_seg_mins, day_seg_start, day_seg_end, start_page " _
			& "FROM tblUsers WHERE user_login = '" & v_sUserName & "'"
		Set oData = New wcData
		getUser = oData.GetArray(sQuery)
		Set oData = nothing
	End Function
End Class
%>
