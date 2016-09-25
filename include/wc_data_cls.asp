<%
' Copyright 1996-2004 Jason Abbott (webcal@webott.com)

'-------------------------------------------------------------------------
'	Name: 		wcData()
'	Purpose: 	encapsulate data functionality
'Modifications:
'	Date:		Name:	Description:
'	12/30/02	JEA		Created
'-------------------------------------------------------------------------
Class wcData
	Public Connection
	Private m_bInTransaction

	Private Sub Class_Initialize()
		Set Connection = Server.CreateObject("ADODB.Connection")
		Connection.Open g_sDB_CONNECT & Server.Mappath(g_sDB_PATH & g_sDB_NAME & ".mdb")
		m_bInTransaction = false
	End Sub
	
	Private Sub Class_Terminate()
		Connection.Close
		Set Connection = nothing
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		*Trans()
	'	Purpose: 	manage ADO transactions
	'Modifications:
	'	Date:		Name:	Description:
	'	1/2/02		JEA		Created
	'-------------------------------------------------------------------------
	Public Sub BeginTrans()
	    Connection.beginTrans
	    m_bInTransaction = True
	End Sub
	Public Sub CommitTrans()
	    If m_bInTransaction Then Connection.commitTrans
		m_bInTransaction = False
	End Sub
	Public Sub RollbackTrans()
    	If m_bInTransaction Then Connection.rollbackTrans
	    m_bInTransaction = False
	End Sub

	'-------------------------------------------------------------------------
	'	Name: 		ExecuteOnly()
	'	Purpose: 	execute query without results
	'Modifications:
	'	Date:		Name:	Description:
	'	12/24/02	JEA		Created
	'-------------------------------------------------------------------------
	Public Sub ExecuteOnly(ByVal v_sQuery)
		Connection.Execute v_sQuery, , adExecuteNoRecords
	End Sub
	
	'-------------------------------------------------------------------------
	'	Name: 		getArray()
	'	Purpose: 	get array from recordset
	'	Return:		array
	'Modifications:
	'	Date:		Name:	Description:
	'	12/23/02	JEA		Created
	'-------------------------------------------------------------------------
	Public Function getArray(ByVal v_sQuery)
		dim oRS
		dim aData
		Set oRS = newRecordSet(v_sQuery)
		If Not oRS.EOF Then aData = oRS.GetRows
		oRS.Close : Set oRS = nothing
		getArray = aData
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		dimDown()
	'	Purpose: 	reduce 2D array to 1D
	'	Return:		array
	' Modifications:
	'	Date:		Name:	Description:
	'	10/8/03		JEA		Created
	'-------------------------------------------------------------------------
	Public Function dimDown(ByVal v_aArray, ByVal x)
		dim aNewArray()
		dim lUBound
		dim y
		
		if IsArray(v_aArray) then
			lUBound = UBound(v_aArray)
			ReDim aNewArray(lUBound)
		
			for y = 0 to lUBound
				aNewArray(y) = v_aArray(y, x)
			next
		else
			aNewArray = v_aArray
		end if
		dimDown = aNewArray
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		getString()
	'	Purpose: 	get delimited string from query
	'	Return: 	string
	'Modifications:
	'	Date:		Name:	Description:
	'	12/30/02	JEA		Created
	'	3/22/03		JEA		Check for empty recordset
	'-------------------------------------------------------------------------
	Public Function getString(ByVal v_sQuery, ByVal v_sColDelim, ByVal v_sRowDelim)
		dim oRS
		Set oRS = NewRecordSet(v_sQuery)
		If Not oRS.EOF Then
			GetString = oRS.GetString(adClipString, , v_sColDelim, v_sRowDelim)
			oRS.Close
		Else
			GetString = ""
		End If
		Set oRS = Nothing
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		getJSArray()
	'	Purpose: 	create JavaScript array literal from array
	'	Return: 	string
	'Modifications:
	'	Date:		Name:	Description:
	'	12/23/02	JEA		Created
	'-------------------------------------------------------------------------
	Public Function getJSArray(ByVal v_sQuery)
		dim sJSArray
		sJSArray = GetString(v_sQuery, """,""", """],[""")
		if sJSArray <> "" then sJSArray = "[""" & Left(sJSArray, Len(sJSArray) - 4) & "]"
		GetJSArray = sJSArray 
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		newRecordSet()
	'	Purpose: 	return open, disconnected recordset
	'Modifications:
	'	Date:		Name:	Description:
	'	12/23/02	JEA		Created
	'-------------------------------------------------------------------------
	Private Function newRecordSet(ByVal v_sQuery)
		dim oRS
		Set oRS = Server.CreateObject("ADODB.Recordset")
		Set oRS.ActiveConnection = Connection
		oRS.CursorLocation = adUseClient
		oRS.Open v_sQuery, , adOpenForwardOnly, adLockReadOnly, adCmdText
		Set oRS.ActiveConnection = nothing
		Set newRecordSet = oRS
		Set oRS = nothing
	End Function
	
	'-------------------------------------------------------------------------
	'	Name: 		funQuery()
	'	Purpose: 	run raw query
	'Modifications:
	'	Date:		Name:	Description:
	'	1/8/03		JEA		Created
	'-------------------------------------------------------------------------
	Public Function funQuery(ByVal v_sQuery)
		dim lAffected
		dim oRS
		If LCase(Left(v_sQuery, 6)) = "select" Then
			' get recordset
			Set RunQuery = NewRecordSet(v_sQuery)
		Else
			' indicate rows affected
			Call Connection.Execute(v_sQuery, lAffected, adExecuteNoRecords)
			Set oRS = Server.CreateObject("ADODB.Recordset")
			With oRS
				.Fields.Append "Affected Rows", adInteger	', , , lAffected
				.Open
				.AddNew
				.Fields("Affected Rows") = lAffected
				'.Update
			End With
			Set funQuery = oRS
			Set oRS = Nothing
		End If
	End Function
End Class
%>