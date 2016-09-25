<% Option Explicit %>
<% Response.Buffer = True %>
<%
dim m_strDSN
dim m_strQueryWrite
dim m_strQueryRead
dim m_oConn
dim m_oRS
dim m_oMail
dim m_strMDAC
dim m_strVBScript
dim m_strFont
dim m_intRandom
dim m_intNumber
dim m_intRecord
dim m_strStatus

m_strFont = "<font face='Tahoma, Arial, Helvetica' size='2'>"
m_strDSN = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
	& Server.Mappath("testdb.mdb")
Randomize
m_intRandom = Int(Rnd * 2000)
m_strQueryWrite = "INSERT INTO tblTest (random) values (" & m_intRandom & ")"
m_strQueryRead = "SELECT id, random FROM tblTest WHERE random = " & m_intRandom
Set m_oConn = Server.CreateObject("ADODB.Connection")
m_oConn.open m_strDSN
m_strMDAC = m_oConn.Version
m_strVBScript = ScriptEngineMajorVersion & "." & ScriptEngineMinorVersion & "." & ScriptEngineBuildVersion
%>
<html>
<head>
<title>webCal Database test</title>
</head>
<body bgcolor="#ffffff">
<%=m_strFont%><font size='4'>
Your current settings are
<table cellspacing='0' cellpadding='2' border='0'>
<tr>
	<td align="center" bgcolor='#bbccee'><%=m_strFont%><b>Component</b></font></td>
	<td align="center" bgcolor='#bbccee'><%=m_strFont%><b>Version</b></font></td>
	<td align="center" bgcolor='#bbccee'><%=m_strFont%><b>For Upgrades</b></font></td>
<tr>
	<td align="right" bgcolor='#ddddbb'><%=m_strFont%>MDAC:</font></td>
	<td><%=m_strFont%><%=m_strMDAC%></font></td>
	<td><%=m_strFont%><a href="http://www.microsoft.com/data">http://www.microsoft.com/data</a></font></td>
<tr>
	<td align="right" bgcolor='#ddddbb'><%=m_strFont%>VBScript:</font></td>
	<td><%=m_strFont%><%=m_strVBScript%></font></td>
	<td><%=m_strFont%><a href="http://msdn.microsoft.com/scripting">http://msdn.microsoft.com/scripting</a></font></td>
</table>
<p>
Connecting to
<table><tr><td bgcolor='#bbccee'><%=m_strFont%><%=m_strDSN%></font></td></tr></table>
<p>
Now attempting to write the random number <%=m_intRandom%> to testdb.mdb ...
<%
response.flush
m_oConn.execute m_strQueryWrite
%>
<p>
Now attempting to read the number from the database ...
<%
response.flush
Set m_oRS = m_oConn.execute(m_strQueryRead)
m_intNumber = m_oRS("random")
m_intRecord = m_oRS("id")
m_oRS.Close : Set m_oRS = nothing
m_oConn.Close : Set m_oConn = nothing
if m_intNumber = m_intRandom then
	m_strStatus = "succeeded"
else
	m_strStatus = "failed"
end if
%>
<p>
Found <%=m_intNumber%> in record <%=m_intRecord%>.  The database test <u><%=m_strStatus%></u>.

<%
'Set m_oMail = Server.CreateObject("CDONTS.NewMail")
%>

</font></font>
</body>
</html>