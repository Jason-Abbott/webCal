<%
' Copyright 1999 Jason Abbott (jason@webott.com)
' Last updated 05/27/1999

dim dataName
dim strDelim
dim db

' To prevent conflicting Session variables with multiple
' copies of this program, specify a unique data name for
' this copy:

dataName = "webCal"
strDelim = "#"			' use ' for SQL, # for Access

' see _themes.inc for possible LCID values
Session.LCID = 1033

' This opens the database stored in the
' the directory /data under the main webCal files

Set db = Server.CreateObject("ADODB.Connection")

' OLEDB Access
DSN = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
	& Server.Mappath("data/" & dataName & ".mdb")

' ODBC Access
'DSN = "DRIVER={Microsoft Access Driver (*.mdb)}; " _
'	& "DBQ=" & Server.Mappath("data/" & dataName & ".mdb")

' OLEDB SQL (add information in brackets)
'DSN = "Provider=SQLOLEDB.1;Password=[input];Persist Security Info=True;" _
'	& "User ID=[input];Initial Catalog=[input];Data Source=[input];"

' ODBC SQL
'DSN = "[input DSN name from ODBC Control Panel]"

' Access only accepts mm/dd/yy dates, so force that format
' returns string----------------------------------------------------------
Function sqlDate(strDate)
	sqlDate = Month(strDate) & "/" & Day(strDate) & "/" & Year(strDate)
End Function

db.Open DSN

' make sure the containing file does
' db.Close
' Set db = nothing
%>