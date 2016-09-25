<%
' Copyright 2001 Jason Abbott (webcal@webott.com)
' Last updated 2/26/2000

dim arLCID			' array of location ids

' create HTML selection list of location ids
' returns string----------------------------------------------------------
Function listlcids(intLCID)
	dim x			' loop counter
	dim strHTML		' build select list
	
	for x = 0 to UBound(arLCID)
		strHTML = strHTML & "<option value=" & arLCID(x)(1)
		if arLCID(x)(1) = intLCID then strHTML = strHTML & " selected"
		strHTML = strHTML & ">" & arLCID(x)(0) & vbCrLF
	next
	listlcids = strHTML
End Function

arLCID = Array( _
	Array("Afrikaans",1078), _
	Array("Albanian",1052), _
	Array("Arabic",5121), _
	Array("Basque",1069), _
	Array("Bulgarian",1026), _
	Array("Catalan",1027), _
	Array("Chinese",3076), _
	Array("Croatian",1050), _
	Array("Czech",1029), _
	Array("Danish",1030), _
	Array("Dutch",1043), _
	Array("English (U.S.)",1033), _
	Array("English (non-U.S.)",2057), _
	Array("Estonian",1061), _
	Array("Faeroese",1080), _
	Array("Farsi",1065), _
	Array("Finnish",1035), _
	Array("French",1036), _
	Array("German",1031), _
	Array("Greek",1032), _
	Array("Hebrew",1037), _
	Array("Hindi",1081), _
	Array("Hungarian",1038), _
	Array("Icelandic",1039), _
	Array("Indonesian",1057), _
	Array("Italian",1040), _
	Array("Japanese",1041), _
	Array("Korean (Ext. Wansung)",1042), _
	Array("Korean (Johab)",2066), _
	Array("Latvian",1062), _
	Array("Lithuanian (Classic)",2087), _
	Array("Lithuanian",1063), _
	Array("Macedonian",1071), _
	Array("Malay",1086), _
	Array("Norwegian",1044), _
	Array("Polish",1045), _
	Array("Portuguese",2070), _
	Array("Romanian",1048), _
	Array("Russian",1049), _
	Array("Serbian (Cyrillic)",3098), _
	Array("Serbian (Latin)",2074), _
	Array("Slovak",1051), _
	Array("Slovene",1060), _
	Array("Spanish",1034), _
	Array("Swahili",1089), _
	Array("Swedish",1053), _
	Array("Thai",1054), _
	Array("Turkish",1055), _
	Array("Ukrainian",1058), _
	Array("Urdu",1056), _
	Array("Vietnamese",1066) _
)
%>