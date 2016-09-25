// form and display options specific to NS
var m_oFrmMain;
var m_oFrmPickup;
var m_oFrmDeliver;
var m_oFrmCCNumber;
var m_oFrmPINumber;
var m_oDivPickup;
var m_oDivDeliver;
var m_oDivInstruct;				// used for place-holder
var m_oDivPickupTitle;
var m_oDivDeliverTitle;
var m_oDivCCNumber;
var m_oDivPINumber;
var m_oDivPayment;
var m_bRefreshed = false;			// has screen already been refreshed

// submit given form (jea:1/11/01)
function submit(r_oForm) { r_oForm.submit(); }

// display delivery fields (jea:1/8/01)
// updates field objects -------------------------------------------------
function showDeliver(r_oField) {
	initDeliver(r_oField);
	if (!(m_bRefreshed)) {
		// resetting window dimension is necessary
		// to have NS redraw fields (jea:1/5/01)
		self.outerHeight += 1;
		m_bRefreshed = true;
	}
}
// display pickup fields (jea:1/8/01)
// updates field objects -------------------------------------------------
function showPickup(r_oField) {
	initPickup(r_oField);
	if (!(m_bRefreshed)) { self.outerHeight += 1;	m_bRefreshed = true; }
}
// hide pickup and delivery fields (jea:1/24/01)
// updates field objects -------------------------------------------------
function showNeither() {
	initNeither();
}
// initialize field and form objects (jea:1/8/01)
// updates global variables ----------------------------------------------
function initDivs() {
	// build references to layers
	m_oDivInstruct = document.layers["instructions"]
	m_oDivPickup = document.layers["pickup"];
	m_oDivDeliver = document.layers["deliver"];
	m_oDivPickupTitle = document.layers["pickuptitle"];
	m_oDivDeliverTitle = document.layers["delivertitle"];
	m_oDivCCNumber = document.layers["ccn"];
	m_oDivPINumber = document.layers["pin"];
	m_oDivPayment = document.layers["payment"];
	
	// make instructions visible
	// move pickup and delivery tables to instructions position
	showDivs([m_oDivInstruct, m_oDivPayment, m_oDivDeliverTitle]);
	matchPosition([m_oDivPickup, m_oDivDeliver], m_oDivInstruct, 10);
	matchPosition([m_oDivPickupTitle], m_oDivDeliverTitle, 30);
	matchPosition([m_oDivCCNumber, m_oDivPINumber], m_oDivPayment, 10);
	
	// build abstracted form references
	m_oFrmMain = document.frmCheckout;
	m_oFrmPickup = m_oDivPickup.document.frmPickup;
	m_oFrmDeliver = m_oDivDeliver.document.frmDeliver;
	m_oFrmCCNumber = m_oDivCCNumber.document.frmCCNumber;
	m_oFrmPINumber = m_oDivPINumber.document.frmPINumber;
	initFields();
}
// prevent changes to store information (jea:1/9/01)
// forces blur away from field -------------------------------------------
function allow(r_oField) {
	if (m_oFrmMain.fldPickOrDlvr[c_lPickup].checked) {
		// applies only to pickup, not delivery
		alert("Store information cannot be changed");
		var oStore = m_oCity[m_sCity].store[m_lFSiteID];
		updateDeliveryFields("Albertsons", "", "", oStore.address, "", sCity, oStore.state, oStore.zip, "", oStore.phone);
	}
}
// show or hide divisions (jea:1/10/01)
// updates form objects --------------------------------------------------
function showDivs(v_aDivs) { for (var x in v_aDivs) { v_aDivs[x].visibility = "show"; } }
function hideDivs(v_aDivs) { for (var x in v_aDivs) { v_aDivs[x].visibility = "hide"; } }

// move given divs to mach position of div plus offset
//updates form objects ---------------------------------------------------
function matchPosition(v_aDivs, r_oDiv, v_lOffset) {
	var lFromTop = r_oDiv.pageY;
	var lFromLeft = r_oDiv.pageX;
	for (var x in v_aDivs) {
		v_aDivs[x].pageY = lFromTop; v_aDivs[x].pageX = lFromLeft + v_lOffset;		
	}
}