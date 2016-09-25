// form and display options specific to IE
var m_oFrmMain;
var m_oFrmPickup;
var m_oFrmCCNumber;
var m_oFrmPINumber;
var m_oFrmDeliver;
var m_oDivDeliver;
var m_oDivPickup;
var m_oDivInstruct;
var m_oDivPickupTitle;
var m_oDivDeliverTitle;
var m_oDivCCNumber;
var m_oDivPINumber;
var m_oDivPayment;

// submit given form (jea:1/11/01)
// updates form objects --------------------------------------------------
function submit(r_oForm) {
	// only change field availability for pickup
	if (m_lPickOrDlvr == c_lPickup) { disableShippingFields(false); }
	r_oForm.submit();
	if (m_lPickOrDlvr == c_lPickup) { disableShippingFields(true); }
}
// display delivery fields (jea:1/8/01)
// updates field objects -------------------------------------------------
function showDeliver(r_oField) {
	if (initDeliver(r_oField)) { disableShippingFields(false); }
}
// display pickup fields (jea:1/8/01)
// updates field objects -------------------------------------------------
function showPickup(r_oField) {
	if (initPickup(r_oField)) { disableShippingFields(true); }
}
// hide pickup and delivery fields (jea:1/24/01)
// updates field objects -------------------------------------------------
function showNeither() {
	disableShippingFields(false);
	initNeither();
}
// initialize field and form objects (jea:1/8/01)
// updates global variables ----------------------------------------------
function initDivs() {
	m_oDivDeliver = document.all.deliver;
	m_oDivPickup = document.all.pickup;
	m_oDivInstruct = document.all.instructions;
	m_oDivPickupTitle = document.all.pickuptitle;
	m_oDivDeliverTitle = document.all.delivertitle;
	m_oDivCCNumber = document.all.ccn;
	m_oDivPINumber = document.all.pin;
	m_oDivPayment = document.all.payment;

	// build abstracted form references
	m_oFrmMain = document.frmCheckout;
	m_oFrmPickup = m_oFrmMain;
	m_oFrmDeliver = m_oFrmMain;
	m_oFrmCCNumber = m_oFrmMain;
	m_oFrmPINumber = m_oFrmMain;
	
	hideDivs([m_oDivDeliver, m_oDivPickup, m_oDivPickupTitle, m_oDivCCNumber, m_oDivPINumber])
	initFields();
}
// disallow or allow editing of address fields for pickup stores (jea:1/9/01)
// updates form objects --------------------------------------------------
function disableShippingFields(v_bHide) {
	for (var x = 0; x < m_aFldAddress.length; x++) {
		eval("m_oFrmMain.fldShip" + m_aFldAddress[x] + ".disabled = " + v_bHide);
	}
}

// show or hide divisions (jea:1/10/01)
// updates form objects --------------------------------------------------
function showDivs(v_aDivs) { for (var x in v_aDivs) { v_aDivs[x].style.display = "block"; } }
function hideDivs(v_aDivs) { for (var x in v_aDivs) { v_aDivs[x].style.display = "none"; } }