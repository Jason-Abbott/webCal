<% Option Explicit %>
<!--#INCLUDE VIRTUAL="shop/include/sitesettings.asp"-->
<%	
' Constants
Const sPAGE_FORM = "pf"
	
' Declare variables
Dim lNavAction
Dim sPageMessage
Dim oList
Dim oCart
Dim oRemind
Dim oShopper
Dim oPromotion
Dim oBundle
Dim lQty
Dim lProductID
Dim lListID
Dim sNewName
Dim lTypeID
Dim lDay
Dim lDayOfMonth
Dim lDayOfWeek
Dim lMonth
Dim rsData
Dim sWelcome
Dim bUpdateListHeader
Dim lMaxQty
Dim sBrowser
Dim sVersion
Dim oDate
Dim sDate
Dim lBundleID
Dim lMenuType
Dim lSelectDay
Dim lSelectMealType
Dim sBundleIDs
Dim sRefresh		' should we auto-refresh page?

' Login processing (jea:3/8/00)-------------------------------------------
dim strLoginForm	' login form
dim oAuthenticate	' AuthenticationUI object
dim oErrors			' dictionary object of validation errors
dim oLogin			' used to assign guest id
dim sUserID
dim sPassword
dim sUserDef		' default login name
dim sPassDef		' default password
dim sErrors			' login errors returned from VB
dim sOpeningPage	' not used

sUserID = Trim(Request.Form(sFLD_USER_ID))
sShopperID = Request.Cookies(sCOOKIE_SHOPPER_ID)
sPassword = Trim(Request.Form(sFLD_PASSWORD))
sBrowser = Request.Cookies(sCOOKIE_BROWSER_TYPE)

If sUserID <> "" then
	' login form was submitted
	Set oAuthenticate = Server.CreateObject("StoreUI.IAuthenticateUI")
	Set oErrors	= oAuthenticate.NoZipValidate(lSTORE_ID,sUserID,sShopperID,sPassword)
	Set oAuthenticate = nothing
	sErrors = oErrors(sFLD_USER_ID) & oErrors(sFLD_PASSWORD)
	Set oErrors = nothing
	
	if sErrors <> "" then
		' show form again with error information
		sRefresh = ""
		sPageMessage = "<font color='990000'>" & sErrors & "</font>"
		sWelcome = makeLogin(sUserID)
	else
		' reload nav_top so lists and checkout are allowed
		sRefresh = "window.setTimeout(""top.message.location = 'message.asp';"", 5000);" & vbCrLf _
			& "top.nav_top.location = 'nav_top.asp';"
		sWelcome = makeWelcome(Request.Cookies(sCOOKIE_SHOPPER_ID))
		sPageMessage =  "<font color='990000'>Successful Login</font>"
	end if
Elseif Request.QueryString("login") = 1 then
	' user clicked "signin" button--reload nav_top and main to disallow lists and checkout
	sRefresh = "top.nav_top.location = 'nav_top.asp';" & vbCrLf _
		& "top.main.location = 'welcome.asp';"
	sUserID = ""
	sShopperID = ""
	Response.Cookies(sCOOKIE_USER_ID) = ""
	Set oLogin = Server.CreateObject("StoreUI.IAuthenticateUI")
	sOpeningPage = oLogin.AutoLogin(sUserID,sShopperID,lSTORE_ID)
	Set oLogin = nothing
	sWelcome = makeLogin("")
Else
	sRefresh = "window.setTimeout(""top.message.location = 'message.asp';"", 10000);"
	sWelcome = makeWelcome(sShopperID)

	' Create variables
	lNavAction			= Trim(Request.Form(sFLD_NAV_ACTION))
	sVersion			= Request.Cookies(sCOOKIE_BROWSER_VERSION)
	bUpdateListHeader	= false
		
	' Process the request
	If lNavAction <> vbNullString then lNavAction = CLng(lNavAction)
	Select Case lNavAction
		Case lACTION_ADD_MARKED
			' Add marked items to cart
			lProductID	 = Request.Form(sFLD_NAV_PRODUCT)
			lQty		 = Request.Form(sFLD_NAV_QTY)
			Set oCart	 = Server.CreateObject("StoreBU.ICart")	
			sPageMessage = oCart.AddMultipleItemsToCart(sShopperID, lSTORE_ID, lProductID, lQty, false)
			Set oCart	 = nothing
		Case lACTION_ADD_TO_LIST
			' Add item to spanning list
			Set oList	 = Server.CreateObject("StoreBU.IList")
			lProductID	 = Request.Form(sFLD_NAV_PRODUCT)
			lQty		 = Request.Form(sFLD_NAV_QTY)
			lListID		 = Request.Form(sFLD_SHOPPING_LIST_ID)
			If lQty		 = vbNullString then lQty = 1
			sPageMessage = oList.AddOneItemToList(sShopperID, lSTORE_ID, lProductID, lQty, lListID)
			Set oList = nothing
		Case lACTION_BUY_NOW
			' Add buy now to cart
			Set oCart	 = Server.CreateObject("StoreBU.ICart")
			lProductID	 = Request.Form(sFLD_NAV_PRODUCT)
			lQty		 = Request.Form(sFLD_NAV_QTY)
			If lQty		 = vbNullString then lQty = 1
			sPageMessage = oCart.AddOneItemToCart(sShopperID, lSTORE_ID, lProductID, lQty)
			Set oCart = nothing
		Case lACTION_UPDATE_CART
			' Update quantity for this product in the cart
			lProductID	 = Request.Form(sFLD_NAV_PRODUCT)
			lQty		 = Request.Form(sFLD_NAV_QTY)
			Set oCart	 = Server.CreateObject("StoreBU.ICart")
			sPageMessage = oCart.AddMultipleItemsToCart(sShopperID, lSTORE_ID, lProductID, lQty, true)
			Set oCart	 = nothing
		Case lACTION_UPDATE_LIST
			' Update quantity for this product in the cart
			lProductID	 = Request.Form(sFLD_NAV_PRODUCT)
			lQty		 = Request.Form(sFLD_NAV_QTY)
			lListID		 = Request.Form(sFLD_SHOPPING_LIST_ID)
			Set oList	 = Server.CreateObject("StoreBU.IList")
			sPageMessage = oList.AddMultipleItemsToList(sShopperID, lSTORE_ID, lListID, lProductID, lQty, true)
			Set oCart	 = nothing
		Case lACTION_REMOVE_FROM_CART
			Set oCart	 = Server.CreateObject("StoreBU.ICart")
			' Remove marked products in the cart
			lProductID	 = Request.Form(sFLD_NAV_PRODUCT)
			sPageMessage = oCart.RemoveItemsFromCart(sShopperID, lSTORE_ID, lProductID)
			Set oCart	 = nothing
		Case lACTION_REMOVE_FROM_LIST
			Set oList	 = Server.CreateObject("StoreBU.IList")
			lListID		 = Request.Form(sFLD_SHOPPING_LIST_ID)
			' Remove marked products in the cart
			lProductID	 = Request.Form(sFLD_NAV_PRODUCT)
			sPageMessage = oList.RemoveItemsFromList(sShopperID, lSTORE_ID, lProductID, lListID)
			Set oList	 = nothing
		Case lACTION_REMOVE_ALL_FROM_CART
			Set oCart	 = Server.CreateObject("StoreBU.ICart")
			' Remove all products in the cart
			sPageMessage = oCart.RemoveAllItemsFromCart(sShopperID, lSTORE_ID)
			Set oCart = nothing
		Case lACTION_REMOVE_ALL_FROM_LIST
			Set oList	 = Server.CreateObject("StoreBU.IList")
			lListID		 = Request.Form(sFLD_SHOPPING_LIST_ID)
			' Remove all products in the cart
			sPageMessage = oList.RemoveAllListItems(sShopperID, lSTORE_ID, lListID)
			Set oList    = nothing
		Case lACTION_SAVE_CART_AS_NEW_LIST
			Set oCart	 = Server.CreateObject("StoreBU.ICart")
			' Save current contents of shopping cart to a list				
			sNewName			= Request.Form(sFLD_NEW_SHOPPINGLIST_NAME)
			lListID				= Request.Form(sFLD_SHOPPING_LIST_ID)
			sPageMessage		= oCart.CopyShoppingCartToNewList(sShopperID, lSTORE_ID, _
																  lListID, sNewName)
			bUpdateListHeader	= true
			Set oCart = nothing
		Case lACTION_DELETE_LIST
			' Remove the specified list 
			Set oList			= Server.CreateObject("StoreBU.IList")
			lListID				= Request.Form(sFLD_SHOPPING_LIST_ID)
			sPageMessage		= oList.RemoveList(lListID)
			bUpdateListHeader	= true
			Set oList	 = nothing
		Case lACTION_RENAME_LIST
			' Rename an existing list
			Set oList			= Server.CreateObject("StoreBU.IList")
			lListID				= Request.Form(sFLD_SHOPPING_LIST_ID)
			sNewName			= Request.Form(sFLD_NEW_SHOPPINGLIST_NAME)
			sPageMessage		= oList.RenameList(sNewName, lListID)
			bUpdateListHeader	= true
			Set oList	 = nothing
		Case lACTION_CHANGE_REMINDERS
			Set oRemind	 = Server.CreateObject("StoreBU.IReminders")
			If Request.Form(sFLD_ALL_REMINDERS_OFF) = "on" then
				sPageMessage = oRemind.TurnRemindersOff(sShopperID, lSTORE_ID)
			Else
				' Change the current reminder settings
				lListID		 = Request.Form(sFLD_SHOPPING_LIST_ID)
				lTypeID		 = CInt(Request.Form(sFLD_NEW_REMINDER_TYPE))
				lDay		 = CInt(Request.Form(sFLD_DAY))
				lDayOfMonth	 = CInt(Request.Form(sFLD_DAY_OF_MONTH))
				lDayOfWeek	 = CInt(Request.Form(sFLD_DAY_OF_WEEK_ID))
				lMonth	     = CInt(Request.Form(sFLD_MONTH_ID))
				sPageMessage = oRemind.UpdateReminder(lListID, lTypeID, lDay, lDayOfMonth, lDayOfWeek, lMonth)
			End if 
			Set oRemind	 = nothing
		Case lACTION_ADD_ALL
			' Add the contents of the currently viewd list to the shopper's cart
			Set oList	 = Server.CreateObject("StoreBU.IList")
			lListID		 = Request.Form(sFLD_SHOPPING_LIST_ID)
			sPageMessage = oList.AddShoppingListToCart(sShopperID, lSTORE_ID, lListID)		
			Set oList	 = nothing
		Case lACTION_ADD_TO_PLANNER
			' Add a recipe to menuplanner
			lBundleID		= Clng(Request(sFLD_BUNDLE_ID))
			lSelectDay		= Clng(Request.Form(sFLD_MENU_DAY))
			lSelectMealType	= Clng(Request.Form(sFLD_MENU_TYPE))
			Set oBundle		= Server.CreateObject("StoreBU.IMenuPlanner")
			sPageMessage	= oBundle.AddRecipeToMenu(sShopperID, lSTORE_ID, lBundleID, lSelectDay, lSelectMealType)			
			Set oBundle		= nothing
		Case lACTION_REMOVE_RECIPE
			' Remove marked recipes from the shopper's menuplanner
			Set oBundle		= Server.CreateObject("StoreBU.IMenuPlanner")
			sBundleIDs		= Request.Form(sFLD_MENU_IDS)
			sPageMessage	= oBundle.RemoveRecipesFromMenu(sBundleIDs)	
			Set oBundle		= nothing			
		Case lACTION_RECIPE_ADD_TO_CART, lACTION_RECIPE_ADD_ALL_TO_CART
			' Remove all recipes from the shopper's menuplanner
			Set oBundle		= Server.CreateObject("StoreBU.IRecipes")
			sBundleIDs		= Request.Form(sFLD_MENU_IDS)
			sPageMessage	= oBundle.AddRecipesToCart(sShopperID, lSTORE_ID, sBundleIDs)
			Set oBundle		= nothing		
		Case lACTION_LOGIN, lACTION_CREATE_ACCOUNT 
			' Update the list header
			bUpdateListHeader	= true
			sPageMessage		= Request.Form(sFLD_MESSAGE)
		Case lACTION_LOGIN_REMINDER, lACTION_UPDATE, lACTION_CHECKOUT
			' Just display the page message
			sPageMessage = Request.Form(sFLD_MESSAGE)
	End Select
End If

' Set the timeout for the message
If IsNumeric(lNavAction) then
%>	
	<script language="Javascript">
		<%=sRefresh%>
	</script>
<%
End if 
If bUpdateListHeader then
%>	
	<script language="Javascript">
		top.nav_top.document.<%=sFRM_MY_LIST%>.submit();
	</script>
<%
End if 
If Request.Form("sPage") = "product.asp" then
	lMaxQty = Request.Form("mxQty")
	If CLng(lMaxQty) < CLng(lQty) then
%>
		<script language="Javascript">
			var oMain = top.main.document.pf;
			oMain.L_<%=lProductID%>.value = <%=lMaxQty%>;
		</script>
<%
	End if 
End if 

If IsNumeric(sVersion) then
	If Len(sPageMessage) > 0 and (sBrowser = "IE" and CInt(sVersion) >= 4) then
		Response.Write "<body bgcolor='#ffffff' leftmargin='0' topmargin='0' onLoad='transitionHead()'>"
	Else
		Response.Write "<body bgcolor='#ffffff' leftmargin='0' topmargin='0'>"
	End if 
Else
	Response.Write "<body bgcolor='#ffffff' leftmargin='0' topmargin='0'>"
End if
%>
<table valign="top" bgcolor="#ffffff" border="0" width="621" cellpadding="0" cellspacing="0">
<form name='login' action="message.asp" method='post'>
<tr>
	<td valign="center" height="25"><font face='Arial, Sans-serif, Helvetica' size=2 color='#0000A5'><b><%=sWelcome%></b></font></td>
<%
If Len(sPageMessage) > 0 then
	' there's a lot of unecessary code here (jea:3/10/00)
	If IsNumeric(sVersion) then
		If (sBrowser = "IE" and CInt(sVersion >= 4)) then
%>
	<td id="idHead" align="right" valign="center">
		<DIV ID="idTransDiv" STYLE="position:relative; top:0; left:0; height:0;
		filter:revealTrans(duration=3.0, transition=0);" align='right' valign='bottom'>
		<font face="Arial,helvetica,sans-serif" Size=-1 color="0000A5"><nobr><b><%=sPageMessage%></b>
		</nobr></font></DIV>
	</td>

<%		ElseIf LCase(sBrowser) = "netscape" then %>

	<td align="right" valign="center">
		<font face="Arial,helvetica,sans-serif" Size=-1 color="0000A5">
		<blink><nobr><b><%=sPageMessage%></b></nobr></blink></font>
	</td>

<%		Else %>

	<td align="right" valign="center">
		<font face="Arial,helvetica,sans-serif" Size=-1 color="0000A5">
		<nobr><b><%=sPageMessage%></b></nobr></font>
	</td>
<%
		End if 
	Else
%>
	<td align="right" valign="center">
		<font face="Arial,helvetica,sans-serif" Size=-1 color="0000A5">
		<nobr><b><%=sPageMessage%></b></nobr></font>
	</td>
<%
	End if 
Else
	Response.Write "<td align='right' valign='center'>"
	Response.Write "<font face='Arial,helvetica,sans-serif' Size=-1 color='0000A5'><nobr><b>"
	Set oDate = Server.CreateObject("StoreUI.IAuthenticateUI")
	sDate = oDate.GetCurrentDate()
	Set oDate = nothing		' added jea:3/10/00
	Response.Write WeekDayName(Weekday(sDate)) & ",&nbsp;" & sDate & "</b></nobr></font></td>"
End if
%>
	</tr>
	<tr>
		<td colspan="2" valign="bottom" align="center" height="6"><img src="./images/line_black.gif" width="621" height="1"></td>
	</tr>
</form>
</table>
</body>
</html>

<%
' Create login form (jea:3/9/00)
Private Function makeLogin(sUserDef)
	dim intWidth		' input field width

	if sBrowser <> "IE" then
		intWidth = 7
	else
		intWidth = 13
	end if
 
	makeLogin = "<table cellspacing='0' cellpadding='0' border='0'><tr>" _
		& "<td rowspan=3 height='25' width='6'><img src='./images/buttons_left.gif'></td>" _
		& "<td bgcolor='#0070AF' height=2><img src='./images/blank.gif' height='2'></td>" _
		& "<td width=56 height=25 rowspan=3>" _
		& "<a href='JavaScript:doLogin();'>" _
		& "<img src='./images/buttons_go_login.gif' width='46' height='25' border='0'></a>" _
		& "</td></tr>" & vbCrLf & "<tr><td bgcolor='#CAD7ED'><nobr>" _
		& "<font face='arial, helvetica' size='2'>Username " _
		& "<input class='Login' type='text' name='" & sFLD_USER_ID & "' size='" & intWidth & "' value='" _
		& sUserDef & "'> Password " _
		& "<input class='Login' type='password' name='" & sFLD_PASSWORD & "' size='" & intWidth & "' value='" _
		& sPassDef & "'></nobr></td></tr>" & vbCrLf & "<tr>" _
		& "<td bgcolor='#0070AF' height=2><img src='./images/blank.gif' height='2'></td></tr></table>"

'CAD7ED
'"<input name='submit' type='image' src='./images/buttons_go_login.gif' width='46' height='25' border='0'>"
End Function

' Create the welcome message 
Private Function makeWelcome(sShopperID)
	dim sHTML		' text returned by function
	
	Set oShopper		= Server.CreateObject("StoreBU.ICommon")
	Set rsData			= oShopper.GetShopperData(sShopperID, lSTORE_ID)
	If not rsData.EOF then
		If Not IsNull(rsData("vsFirstName")) then
			sHTML = "Welcome&nbsp;back&nbsp;<i>" & rsData("vsFirstName") & "</i>"
		Else
			' if not user data then show login (jea:3/10/00)
			sHTML = makeLogin("")
		End if 
	End if
	rsData.Close
	Set rsData = nothing
	Set oShopper = nothing
	
	makeWelcome = sHTML
End Function

%>

<script LANGUAGE="JavaScript"><!--
function transitionHead()
{
	idHead.style.visibility = "hidden";
	idTransDiv.filters.item(0).apply();
	idTransDiv.filters.item(0).transition = 12;
	idHead.style.visibility = "visible";
	idTransDiv.filters(0).play(1.000);
}

function doLogin() {
	var oMain = document.forms[0];
	var errors = "";
	if (oMain.<%=sFLD_USER_ID%>.value.length == 0) {
		errors = "Please enter your username";
	}
	if (oMain.<%=sFLD_PASSWORD%>.value.length == 0) {
		if (errors.length == 0) {
			errors = "Please enter your password";
		} else {
			errors = errors + " and password";
		}
	}
	if (errors.length == 0) {
		oMain.action = "<%=Request.ServerVariables("SCRIPT_NAME")%>";
		oMain.submit();
	} else {
		alert(errors);
	}
}

//-->
</script>