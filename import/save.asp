<html>
<head>
<title>testing import</title>
</head>
<body>

<!--#include file="../data/webCal4_data_inc.asp"-->
<%
dim intBytes
dim strHead
dim arData
dim x

strFile = "c:\temp.mdb"

intBytes = Request.TotalBytes
arData = Request.BinaryRead(intBytes)

intStart = 1
Do Until Right(strHead, 8) = "13101310"
	strHead = strHead & AscB(MidB(arData, intStart, 1))
	intStart = intStart + 1
Loop

intEnd = intBytes - 48

for x = intStart to intEnd
	strData = strData & Chr(AscB(MidB(arData, x, 1)))
next

'response.write "'" & strHead & "'<p>" & x & "<p>"

'For x = 1 to intBytes
'	strThing = Chr(AscB(MidB(arData, x, 1)))
'	response.write MidB(arData, x, 1)
'	response.write CStr(strThing)
'	response.write "<font size=1 color='#999999'>" & AscB(strThing) & "</font>"
'Next

Set oFS = Server.CreateObject("Scripting.FileSystemObject")
Set oFile = oFS.CreateTextFile(strFile,true)
oFile.Write(strData)

oFile.Close
Set oFile = nothing

g_strDSN2 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strFile

strQuery = "SELECT * FROM Calendar ORDER BY StartDate, StartTime"
oRS.Open strQuery, g_strDSN2, adOpenForwardOnly, adLockReadOnly, adCmdText

do while not oRS.EOF
	response.write oRS("Subject") & "<br>"
	oRS.MoveNext
loop

oFS.DeleteFile(strFile)
Set oFS = nothing
oRS.Close
Set oRS = nothing
%>

file written
</body>
</html>