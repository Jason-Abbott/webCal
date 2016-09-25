<%
' Copyright 1999 Jason Abbott (jason@webott.com)
' Last updated 05/27/1999

' check to see if the user has already logged in
' if not, send them to the login page

dim strPath
' accomodate virtual host (jea:3/8/00)
strPath = Request.ServerVariables("PATH_TRANSLATED")
strPath = Right(strPath, Len(strPath) - InStrRev(strPath,"\"))

if Session(dataName & "User") = "" OR Session(dataName & "User") = 0 then
	response.redirect "webCal3_login.asp?url=" _
		& strPath & "?" _
		& Server.URLEncode(Request.ServerVariables("QUERY_STRING"))
end if
%>