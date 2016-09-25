// Copyright 2001 Jason Abbott (webcal@webott.com)
// Last updated 2/24/2001

var m_oFields = { 
	fldNameFirst:{desc:"First Name",type:"String",req:1},
	fldNameLast:{desc:"Last Name",type:"String",req:1},
	fldEmail:{desc:"E-mail Address",type:"Email",req:0},
	fldPassword:{desc:"Password (they don't match)",type:"Password",req:1}
	};

var m_oForm;

function initPage(v_sType) {
	m_oForm = document.frmEdit;
	if (v_sType == "update") { showAccess(); }
	showMessage();
}
// submit form to save user (updated 2/25/01)
// validates fields and submits form -------------------------------------
function saveUser(v_sForm, v_bAgain) {
	var sQS = "?again=" + ((v_bAgain) ? "1" : "0");
	if (isValid(v_sForm, m_oFields)) {
		m_oForm.action = "webCal4_user-updated.asp" + sQS;
		saveAccess();
		m_oForm.submit();
	}
}
// update the access list when the group selection is changed
// updates form element --------------------------------------------------
function newGroup(r_oField) {
	var lGroupID = r_oField.options[r_oField.selectedIndex].value;
	m_oForm.fldAccess.options[m_oNewGroup[lGroupID]].selected = true;
}
// save access selection whenever it is changed (updated 2/24/01)
// displays alert or updates form element---------------------------------
function newAccess(r_oField) {
	var lAccessID = r_oField.options[r_oField.selectedIndex].value;
	var lGroupID = m_oForm.fldGroup.options[m_oForm.fldGroup.selectedIndex].value;
	var lDefault = m_oForm.fldDefault.options[m_oForm.fldDefault.selectedIndex].value;
	if (lDefault == g_ADMIN_ACCESS) {
		// this should only apply to NS since IE option lists are disabled
		alert("Administrators cannot have their group access limited. " +
			"If you want to limit this user's access, change their " +
			"default permission setting to a level other than administrator.");
		lAccessID = g_MGR_ACCESS;
		showAccess();		
	} else {
		m_oNewGroup[lGroupID] = lAccessID;
		//saveAccess(lDefault);
	}
}
// update per group access when default is changed
// updates form elements -------------------------------------------------
function newDefault(r_oField) {
	var lDefault = r_oField.options[r_oField.selectedIndex].value;
	var lMaxAccess, sUserType;
	if (lDefault == g_ADMIN_ACCESS) {
		lMaxAccess = g_MGR_ACCESS;
		// admin access overwrites per-group so treat as new user
		sUserType = "new";
		m_oForm.fldAccess.disabled=1;
		m_oForm.fldGroup.disabled=1;
	} else {
		lMaxAccess = lDefault;
		m_oForm.fldAccess.disabled=0;
		m_oForm.fldGroup.disabled=0;
	}
	if (sUserType == "new") {
		// only overwrite per group settings for new users
		for (var lGroupID in m_oNewGroup) { m_oNewGroup[lGroupID] = lMaxAccess; }
		saveAccess(lMaxAccess);
		showAccess();
	}
}
// save access settings to string for posting (updated 2/24/01)
//   format (group id|access level|newness,[repeat]) 
// updates form element --------------------------------------------------
function saveAccess(v_lDefault) {
	var sLevels = "";
	var bUpdate = false;
	for (var lGroupID in m_oNewGroup) {
		if (m_oNewGroup[lGroupID] == v_lDefault ||
			m_oNewGroup[lGroupID] != m_oOldGroup[lGroupID]) {
			
			sLevels += "," + lGroupID + "|";
			if (m_oNewGroup[lGroupID] == v_lDefault) {
				// remove groups with access same as default
				sLevels += "0|"
			} else {
				// insert selected group access
				sLevels += m_oNewGroup[lGroupID] + "|";
			}
			// append 1 if group is new to this user
			sLevels += ((m_oOldGroup[lGroupID] == 0) ? "1" : "0");
		}
	}
	m_oForm.fldAccessList.value = sLevels.substr(1);	// remove extra comma
}
// initialize group access level
// updates form element --------------------------------------------------
function showAccess() {
	var lGroupID = m_oForm.fldGroup.options[0].value;
	m_oForm.fldAccess.options[m_oNewGroup[lGroupID]].selected=true;
}