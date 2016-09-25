<html>
<head>
<title>testing import</title>

<%
dim oFS
dim oFile
dim strFile
dim strText
dim arHeader
dim arRecords()
dim intCount
dim bString
dim x
dim y

x = 0
strFile = "c:\temp\webCal\webcal.csv"

Set oFS = Server.CreateObject("Scripting.FileSystemObject")
Set oFile = oFS.OpenTextFile(strFile)

' read header into array
arHeader = Split(Replace(oFile.ReadLine,"""",""),",")
intCount = UBound(arHeader)

Do While oFile.AtEndOfStream <> True
	strText = oFile.Readline

	if UBound(Split(strText,",")) > intCount then
		' this record must span multiple lines
		if Right(Trim(strText),2) = ",""" then
			bString = 1
		else
			bString = 0
		end if
		response.write bString
	
	
	end if
	
'	Do Until UBound(Split(strText,",")) >= intCount - 2
'		response.write UBound(Split(strText,",")) & ","
'		' the record must span another line
'		strText = strText & oFile.ReadLine
'	Loop
	
	ReDim Preserve arRecords(x)
	arRecords(x) = Split(strText,",")
	x = x + 1
Loop
  
'oFS.DeleteFile(strFile)

oFile.Close
Set oFile = nothing
Set oFS = nothing

%>

</head>
<body>

<table border=1>
<tr>
<%
for x = 0 to intCount
	response.write "<td>" & arHeader(x) & "</td>" & vbCrLf
next

for x = 0 to UBound(arRecords)
	response.write "<tr>"
	for y = 0 to intCount
		response.write "<td>" & arRecords(x)(y) & "</td>" & vbCrLf
	next
next
%>
</table>

</body>
</html>