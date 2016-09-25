// Copyright 2001 Jason Abbott (webcal@webott.com)
// Last updated 2/27/2001

var m_oFields = { 
	fldGroupName:{desc:"Group Name",type:"String",req:1}
	};

var m_oForm;					// current form object
var m_oGroupUser = new Object;	// group object
var m_sPadList = "";

function initPage(v_sType) {
	m_oForm = document.frmEditGroup;
	buildGroupUsers();
	buildUserLists();
	showMessage();
	if (m_ns) { window.outerWidth += 1; }		// force NS to redraw option lists
}
// object constructor for group users (updated 2/27/01)
// returns new object ----------------------------------------------------
function GroupUser(v_sUserName, v_lDefaultAccess, v_lNewAccess, v_lOldAccess) {
	this.name = v_sUserName;
	this.old_access = (v_lOldAccess != "") ? v_lOldAccess * 1 : -1;
	this.new_access = v_lNewAccess;
	this.default_access = v_lDefaultAccess * 1;
}
// build group user object using array returned from query (updated 2/27/01)
// updates page scope group user object ----------------------------------
function buildGroupUsers() {
	var c_UserID = 0, c_UserName = 1, c_UserDefault = 2, c_UserAccess = 3;
	var sUserName;
	var lNewAccess;
	var lMaxLength = 0;
	for (var x = 0; x < m_aGroupUsers.length; x++) {
		sUserName = m_aGroupUsers[x][c_UserName];
		lNewAccess = m_aGroupUsers[x][c_UserAccess];
		if (lNewAccess == 0 || lNewAccess == "") {
			// use default access if none specific to group
			lNewAccess = m_aGroupUsers[x][c_UserDefault] * 1;
		}
		if (lMaxLength < sUserName.length) { lMaxLength = sUserName.length; }
		m_oGroupUser[m_aGroupUsers[x][c_UserID]] = new GroupUser(
			sUserName,
			m_aGroupUsers[x][c_UserDefault],
			lNewAccess,
			m_aGroupUsers[x][c_UserAccess])
	}
	delete m_arGroupUsers;
	// build string to pad lists so their width remains fixed
	for (x = 0; x < lMaxLength + 4; x++) { m_sPadList += "  "; }
}
// build users drop-downs using group users object (updated 2/27/01)
// updates form objects --------------------------------------------------
function buildUserLists() {
	var oFldMembers = m_oForm.fldMembers;
	var oFldNonMembers = m_oForm.fldNonMembers;
	var m = 0, n = 0;
	var sUserName;
	for (var lUserID in m_oGroupUser) {
		if (m_oGroupUser[lUserID].new_access != g_NO_ACCESS) {
			// add to members list
			sUserName = m_oGroupUser[lUserID].name + showAccess(m_oGroupUser[lUserID].new_access);
			oFldMembers.options[m] = new Option(sUserName, lUserID);
			m++;
		} else {
			// add to non-members list
			oFldNonMembers.options[n] = new Option(m_oGroupUser[lUserID].name, lUserID);
			n++;
		}
	}
	// add pad for fixed width
	oFldMembers.options[m] = new Option(m_sPadList, 0);
	oFldNonMembers.options[n] = new Option(m_sPadList, 0);
}
// show abbreviation representing user access in this group (updated 2/27/01)
// returns string --------------------------------------------------------
function showAccess(v_lAccess) {
	var sAbbrev;
	v_lAccess = v_lAccess * 1;		// ensure numeric
	switch (v_lAccess) {
		case g_READ_ACCESS: sAbbrev = "R"; break;
		case g_ADD_ACCESS: sAbbrev = "A"; break;
		case g_EDIT_ACCESS: sAbbrev = "E"; break;
		case g_MGR_ACCESS: sAbbrev = "M"; break;
		default: return "";
	}
	return " [" + sAbbrev + "]";
}
// submit form to save group (updated 2/25/01)
// validates fields and submits form -------------------------------------
function saveGroup(v_sForm, v_bAgain) {
	var sQS = "?again=" + ((v_bAgain) ? "1" : "0");
	if (isValid(v_sForm, m_oFields)) {
		m_oForm.action = "webCal4_group-updated.asp" + sQS;
		saveAccess();
		m_oForm.submit();
	}
}
// move selected user to member or non-member (updated 2/27/01)
// updates form objects --------------------------------------------------
function moveUser(v_sDir) {
	var oFldFrom, oFldTo;
	var lIndex, sUserName, lUserID;
	if (v_sDir == "remove") {
		oFldFrom = m_oForm.fldMembers;
		oFldTo = m_oForm.fldNonMembers;
		lIndex = oFldFrom.selectedIndex;
		sUserName = oFldFrom.options[lIndex].text;
		// remove access abbreviation
		sUserName = sUserName.substr(0, sUserName.length - 4);
	} else {
		oFldFrom = m_oForm.fldNonMembers;
		oFldTo = m_oForm.fldMembers;
		lIndex = oFldFrom.selectedIndex;
		sUserName = oFldFrom.options[lIndex].text;
		// default access is read-only
		sUserName += " [R]";
	}
	lUserID = oFldFrom.options[lIndex].value;
	
	// fail if a user wasn't selected
	if (lUserID < 1) {
		alert("You must select a user to " + v_sDir);
		return false;
	}
	addToList(oFldTo, sUserName, lUserID);
	removeFromList(oFldFrom, lIndex)
	
	// set default permissions and update display
	if (v_sDir == "add") {
		m_oGroupUser[lUserID].new_access = 1;
		m_oForm.fldAccessLevel.options[0].selected = true;
	} else {
		m_oGroupUser[lUserID].new_access = 0;
	}
	//saveList();
	return true;
}
// add new entry to top of existing list (updated 2/27/01)
// updates field object --------------------------------------------------
function addToList(r_oField, v_sText, v_sValue) {
	for (var x = r_oField.length; x > 0; x--) {
		// move old entries down so new entry can be at top
		r_oField.options[x] = new Option(r_oField.options[x-1].text, r_oField.options[x-1].value);
	}
	r_oField.options[0] = new Option(v_sText, v_sValue);
	r_oField.options[0].selected = true;
}
// remove entry from existing list (updated 2/27/01)
// updates field object --------------------------------------------------
function removeFromList(r_oField, v_lIndex) {
	for (x = v_lIndex; x < r_oField.length - 1; x++) {
		r_oField.options[x] = new Option(r_oField.options[x+1].text, r_oField.options[x+1].value);
	}
	r_oField.length -= 1;
}
// update access for selected user (updated 2/27/01)
// updates form objects --------------------------------------------------
function newAccess(r_oField) {
	var lAccess = r_oField.options[r_oField.selectedIndex].value;
	var oOptUser = m_oForm.fldMembers.options[m_oForm.fldMembers.selectedIndex];
	var lUserID = oOptUser.value;
	var sUserName = oOptUser.text;
	if (lUserID < 1) {
		alert("You must select a user before changing permissions");
		return false;
	}
	m_oGroupUser[lUserID].new_access = lAccess;
	// remove old access abbreviation and add new
	sUserName = sUserName.substr(0, sUserName.length - 4)
	sUserName += showAccess(lAccess);
	oOptUser.text = sUserName;
}	
// save user access settings to string for posting (updated 2/27/01)
//   format (user id|access level|default access|newness,[repeat]) 
// updates form element --------------------------------------------------
function saveAccess() {
	var sLevels = "";
	var bUpdate = false;
	var oUser;
	for (var lUserID in m_oGroupUser) {
		oUser = m_oGroupUser[lUserID];
		if (needUpdate(oUser.old_access, oUser.new_access, oUser.default_access)) {
			// save only if access has changed
			sLevels += "," + lUserID + "|" + oUser.new_access +
				"|" + oUser.default_access + "|";
			sLevels += (oUser.old_access < 0) ? "1" : "0";
		}
	}
	m_oForm.fldAccessList.value = sLevels.substr(1);	// remove extra comma
}
// process business rules for updating permissions (updated 2/27/01)
// returns boolean -------------------------------------------------------
function needUpdate(v_lOldAccess, v_lNewAccess, v_lDefaultAccess) {
	if (v_lOldAccess < 0) {
		// no previous entry in tblPermissions
		if (v_lNewAccess != v_lDefaultAccess) { return true; }
	} else {
		// previous entry exists
		if (v_lNewAccess != v_lOldAccess) { return true; }
	}
	return false;
}
// update the access level to match the selected member (updated 2/27/01)
// updates form object ---------------------------------------------------
function newUser(r_oField) {
	var lUserID = r_oField.options[r_oField.selectedIndex].value;
	if (lUserID < 1) { return false; }
	var lAccess = m_oGroupUser[lUserID].new_access;
	var lIndex = lAccess - 1;
	m_oForm.fldAccessLevel.options[lIndex].selected = true;
}