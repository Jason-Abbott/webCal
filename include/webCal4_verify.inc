<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' Last updated 03/25/2000

' check to see if the user has already logged in
' if not, send them to the login page

dim m_strPath

' this is necessary to accomodate virtual hosts
m_strPath = Request.ServerVariables("PATH_TRANSLATED")
m_strPath = Right(m_strPath, Len(m_strPath) - InStrRev(m_strPath,"\"))

if Session(g_unique & "UserID") = 0 then
	response.redirect "webCal4_login.asp?url=" _
		& m_strPath & "?" _
		& Server.URLEncode(Request.ServerVariables("QUERY_STRING"))
end if

' if the user does nothing for five minutes, clear this secured page
%>
<script language="javascript">
	//window.setTimeout("top.location = 'default.asp';", 300000);
</script>