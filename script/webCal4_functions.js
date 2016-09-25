var c_ckMessage = "Message";

// Access constants
var g_NO_ACCESS = 0;
var g_READ_ACCESS = 1;
var g_ADD_ACCESS = 2;
var g_EDIT_ACCESS = 3;
var g_MGR_ACCESS = 4;
var g_ADMIN_ACCESS = 5;

var m_ie = document.all;
var m_ns = document.layers;
var x = 0;
var y = 0;

// logout
function logout() {
	alert("future logout");
}

// change image and display status bar message (updated 3/2/01)
// updates objects -------------------------------------------------------
function iconOver(v_strName, v_strSource, v_strMessage) {
	if (document.images) {
  		document.images[v_strName].src = eval("icon" + v_strSource + "On.src");
		document.status = v_strMessage;
	}
}
// restore image and clear status bar (updated 3/2/01)
// updates objects -------------------------------------------------------
function iconOut(v_strName, v_strSource) {
	if (document.images) {
  		document.images[v_strName].src = eval("icon" + v_strSource + ".src");
		document.status = "";
	}
}
// read cookie with given name (updated 7/31/00)
// returns string --------------------------------------------------------
function getCookie(v_sName) {
	re = new RegExp(v_sName + "=(\\S+)[\\;\\b]");
	if (re.test(document.cookie)) {
		var aMatch = re.exec(document.cookie);
		var sValue = unescape(aMatch[1])
		sValue = sValue.replace(/\+/gi, " ");
		return(sValue);
	}
	return "";
}
// save a cookie (updated 8/14/00)
// updates collection ----------------------------------------------------
function setCookie(v_sName, v_sValue) {
	document.cookie = v_sName + "=" + escape(v_sValue);
}
// erase cookie (updated 8/22/00)
// updates collection ----------------------------------------------------
function delCookie(v_sName) {
    document.cookie = v_sName + "=; expires=Thu, 01-Jan-80 00:00:01 GMT";
}
// display any pending messages (updated 2/22/01)
// shows alert, returns boolean ------------------------------------------
function showMessage() {
	var sMessage = getCookie(c_ckMessage);
	if (sMessage != "") { alert(sMessage); delCookie(c_ckMessage); return true; }
	return false;
}
// queue message (updated 2/20/01)
function saveMessage(v_sValue) { setCookie(c_ckMessage, v_sValue); }

// queue error message and go to given page (updated 3/1/01)
function goPageMessage(v_sMessage, v_sPage) {
	saveMessage(v_sMessage);
	(v_sPage != "") ? location.replace(v_sPage) : history.back();
}
// track mouse position to pop window by click (updated 2/24/01)
// updates page scope variables ------------------------------------------
function trackMouse() {
	if (m_ns) { document.captureEvents(Event.MOUSEMOVE); }
	document.onmousemove = mouseMove;
}
// update screen coordinates from mouse position (updated 4/20/98)
// updates pages scope variables -----------------------------------------
function mouseMove(e) {
	if (m_ie) { x = event.screenX; y = event.screenY; }
	if (m_ns) { x = e.screenX; y = e.screenY; }
}
// pop miniature calendar (updated 2/24/01)
// invokes browser instance ----------------------------------------------
function calPop(v_sForm, v_sField) {
	var bOpen = false;
	if (typeof(calWin) == "object") {
		if (!(calWin.closed)) { bOpen = true; }
	}
	if (bOpen) {
		// window already open--put in focus
		calWin.focus();
	} else {
		// create new window
		x = x + 25; y = y - 10;
		var oField = eval("document." + v_sForm + "." + v_sField);
		var sDate = (isDate(oField)) ? oField.value : "";
		var url = "webCal4_mini.asp?field=" + v_sField + "&form=" + v_sForm + "&date=" + sDate
		calWin = window.open(url,"calendar","height=140,width=140,scrollbars=no,titlebar=no,resizable,screenX="+x+",left="+x+",screenY="+y+",top="+y);
	}
}
// submit form to given page (updated 2/24/01)
// submits form ----------------------------------------------------------
function goPage(v_sAction, v_sFormName) {
	var oForm = eval("document." + v_sFormName);
	oForm.action = v_sAction;
	oForm.submit();
}
// read query string value (updated 8/29/00)
// returns string --------------------------------------------------------
function getQS(v_sName) {
// [&\\b]
	re = new RegExp("[&?]" + v_sName + "=(\\w+)\\b");
	if (re.test(unescape(location.search))) {
		var aMatch = re.exec(unescape(location.search));
		return(aMatch[1]);
	}
	return "";
}