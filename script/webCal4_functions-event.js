// instantiate validation objects
var m_oFields = { 
	fldTitle:{desc:"Event Title",type:"String",req:1},
	fldShowTo:{desc:"Visibility (some visibility must be granted to at least one group)",type:"Scoped",req:1} };

var m_oForm;				// current form object
var m_oGroup = new Object;	// group object

// object constructor for group scopes (updated 2/24/01)
// returns new object ----------------------------------------------------
function Group(v_sGroupName, v_lScopeID, v_lHTMLTitle, v_lHTMLDesc, v_lHTMLLoc) {
	this.name = v_sGroupName;
	this.scope = (v_lScopeID != "") ? v_lScopeID : 0;
	this.html_title = (v_lHTMLTitle == "1");
	this.html_desc = (v_lHTMLDesc == "1");
	this.html_loc = (v_lHTMLLoc == "1");
}
// build group object using group array returned from query (updated 2/24/01)
// updates page scope group object ---------------------------------------
function buildGroups() {
	// constants to index array of group properties
	var g_GROUP_ID = 0, g_GROUP_NAME = 1, c_GroupScope = 2;
	var c_GroupTitleHtml = 3, c_GroupDescHtml = 4, c_GroupLocHtml = 5;
	for (var x = 0; x < m_arEditScopes.length; x++) {
		m_oGroup[m_arEditScopes[x][g_GROUP_ID]] = new Group(
			m_arEditScopes[x][g_GROUP_NAME],
			m_arEditScopes[x][c_GroupScope],
			m_arEditScopes[x][c_GroupTitleHtml],
			m_arEditScopes[x][c_GroupDescHtml],
			m_arEditScopes[x][c_GroupLocHtml])
	}
	delete m_arEditScopes;
}
// build group drop-down using group object (updated 2/24/01)
// updates form objects --------------------------------------------------
function buildGroupList() {
	var oFldGroup = m_oForm.fldGroup;
	var x = 0;
	for (var lGroupID in m_oGroup) {
		oFldGroup.options[x] = new Option(m_oGroup[lGroupID].name, lGroupID);
		x++;
	}
	oFldGroup.options[0].selected = true;
}
// set default page values (updated 2/21/01)
// updates form objects --------------------------------------------------
function initPage() {
	m_oForm = document.frmEdit;
	m_oForm.fldTitle.focus();
	buildGroups();
	buildGroupList();
	var oFldGroup = m_oForm.fldGroup;
	var lGroupID = oFldGroup.options[oFldGroup.selectedIndex].value;
	m_oForm.fldShowTo.options[m_oGroup[lGroupID].scope].selected = true;
	updateUserScopeList();
	trackMouse();	// track mouse position to pop calendar
	showMessage();	// display any pending messages
	if (m_ns) { window.outerWidth += 1; }
}
// submit form to save event (updated 2/21/01)
// validates fields and submits form -------------------------------------
function saveEvent(v_sForm, v_bAgain) {
	var sQS = "?again=" + ((v_bAgain) ? "1" : "0");
	if (isValid(v_sForm, m_oFields)) {
		m_oForm.action = "webCal4_event-updated.asp" + sQS;
		updateUserScopeList();
		m_oForm.submit();
	}
}
// process "No Specific Time" checkbox (updated 2/21/01)
// updates form objects --------------------------------------------------
function newTimeCheck(r_oField) {
	// Netscape doesn't support the disabled property
	var aFields = ["fldStartHour","fldStartMin","fldEndHour","fldEndMin"]
	var bEnable = (r_oField.checked) ? 1 : 0;
	for (var x = 0; x < aFields.length; x++) {
		eval("m_oForm." + aFields[x] + ".disabled = " + bEnable);
	}
}
// process new recurrence selection (updated 2/24/01)
// updates form objects --------------------------------------------------
function newRecur(r_oField) {
	var sRecur = r_oField.options[r_oField.selectedIndex].value;
	if (sRecur == "none") { m_oForm.fldEndDate.value = ""; }
}
// update the scope option list when the group selection is changed (updated 2/21/01)
// updates form objects --------------------------------------------------
function newGroup(r_oField) {
	var lGroupID = r_oField.options[r_oField.selectedIndex].value;
	m_oForm.fldShowTo.options[m_oGroup[lGroupID].scope].selected = true;
}
// save scope settings whenever it is changed (updated 2/21/01)
// updates form objects --------------------------------------------------
function newUserScope(r_oField) {
	var lGroupID = m_oForm.fldGroup.options[m_oForm.fldGroup.selectedIndex].value;
	var lScopeID = r_oField.options[r_oField.selectedIndex].value;
	m_oGroup[lGroupID].scope = lScopeID;
}
// save scope selections to hidden field (updated 2/21/01)
// updates form objects --------------------------------------------------
function updateUserScopeList() {
	var sViews = "", lScopeID;
	for (var lGroupID in m_oGroup) {
		if (m_oGroup[lGroupID].scope != 0) {
			lScopeID = m_oGroup[lGroupID].scope;
			sViews += "," + lGroupID + "|" + lScopeID;
		}
	}
	sViews = sViews.substr(1);	// remove leading comma
	m_oForm.fldUserScopes.value = sViews;
}