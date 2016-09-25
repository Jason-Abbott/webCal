// dependencies: m_oFrmMain, m_oFrmPickup, m_oFrmDelivery, m_sURL
var m_bInitDone = false;	// initialization status
var c_lPickup = 0;			// constant for radio button index
var c_lDeliver = 1;			// constant for radio button index
var m_lPickOrDlvr = -1;		// track user selection
var c_sDlvrCity = "delivery";
var m_oCity;				// data object
var m_oCustomer;			// data object
var m_oState;				// data object
var m_lPickupCount = 0;		// number of pickup ranges
var m_lDeliveryCount = 0;	// number of delivery ranges
var m_sCity;				// selected city
var m_lFSiteID;				// selected store id
var m_lPickupFSiteID;		// selected pickup store
var m_lAreaID;				// assigned area id
var m_sDay;					// selected day
var m_lRangeID				// selected range id
var m_oForm;				// form object
var m_lTries = 0;			// retry counter
var m_aFldAddress = ["First","MI","Last","Address1","Address2","City","State","Zip","Zip4","Phone"];
var m_oFldDate;				// date field for pickup or delivery
var m_oFldTime;				// time field for pickup or delivery
var m_oFldCity;				// pickup city field
var m_oFldStore;			// pickup store field
var m_sZips;				// string list of all zips in market region
var m_sForm = 'frmCheckout';
var m_sLine = '-------------------------------';
var m_sJoke = '';			// secret
var m_oDataFrame = top.codata;
var m_oTotalFrame = parent.cototal;
var m_oMsgFrame = top.message;
var m_bWait = false;		// waiting for validation from totals frame
var m_bPosted = false;		// waiting for page post

// retrieve data and populate city list (jea:1/9/01)
// updates variables form object -----------------------------------------
function initData() {
	var bDataReady = m_oDataFrame.m_bDataReady;
	if (typeof(bDataReady) == "boolean") {
		if (bDataReady) {
			m_oCity = m_oDataFrame.m_oCity;
			m_oCustomer = m_oDataFrame.m_oCustomer;
			m_oState = m_oDataFrame.m_oState;
			countRanges();
			// build list of zips for entire market region
			var aZips = new Array(0);	// all deliverable zip codes in region
			m_sCity = c_sDlvrCity;		// default city
			for (m_lFSiteID in m_oCity[m_sCity].store) {
				for (m_lAreaID in m_oCity[m_sCity].store[m_lFSiteID].area) {
					aZips = aZips.concat(m_oCity[m_sCity].store[m_lFSiteID].area[m_lAreaID].zips);
				}
			}
			m_sZips = "," + aZips.toString() + ",";
			makeCityList();				// for pickup
			if (m_oCustomer.fsiteid == 0 || m_sZips.search("," + m_oCustomer.zip + ",") == -1) {
				// customer zip is not valid for market region
				updateDelivery("badzip");
				m_lFSiteID = 0;
				m_lAreaID = "";
			} else {
				// initialize data for customer
				m_lFSiteID = m_oCustomer.fsiteid;
				m_lAreaID = getArea(m_oCustomer.zip);
				makeDateList(c_lDeliver);
			}
			m_oFrmPickup.fldRecipientTemp.value = m_oCustomer.firstname + ' ' + m_oCustomer.lastname;
			m_bInitDone = true;
			if (m_lPickupCount == 0) {
				// no pickup stores available
				m_oFrmMain.fldPickOrDlvr[c_lPickup].disabled = true;	// ignored by ns
				m_oFrmMain.fldPickOrDlvr[c_lDeliver].checked = true;
				showDeliver(m_oFrmMain.fldPickOrDlvr[c_lDeliver]);
			}
			if (m_lDeliveryCount == 0) {
				// no delivery stores available
				m_oFrmMain.fldPickOrDlvr[c_lDeliver].disabled = true;	// ignored by ns
				m_oFrmMain.fldPickOrDlvr[c_lPickup].checked = true;
				showPickup(m_oFrmMain.fldPickOrDlvr[c_lPickup]);
			}
			//alert(m_lPickupCount + ', ' + m_lDeliveryCount);
			stopProgressBar();
			return;
		}
	}
	// try again in .2 seconds--retry for 10 seconds
	if (m_lTries > 50) { 
		alert("There was a problem loading the data\nfor delivery and pickup options.\n\nPlease contact customer service or try\nchecking out again in a few minutes.");
		top.location.href = m_sURL + 'default.asp?page=welcome.asp';
		return;
	}
	m_lTries++;
	m_oFldCity.options[0].text += " .";
	m_oFrmDeliver.fldDeliverDate.options[0].text += " .";
	setTimeout("initData()", 200);
}
// retrieve area appropriate for given zip code (jea:1/19/01)
// returns number --------------------------------------------------------
function getArea(v_lZip) {
	var oStore = m_oCity[m_sCity].store[m_lFSiteID];
	var sAreaZips;
	for (var lTempAreaID in oStore.area) {
		sAreaZips = "," + oStore.area[lTempAreaID].zips.toString() + ",";
		if (sAreaZips.search("," + v_lZip + ",") != -1) {
			// zip matches this area
			return lTempAreaID;
		}
	}
}
// count the ranges to ascertain the availability of pickup or delivery (jea:1/17/01)
// updates variables -----------------------------------------------------
function countRanges() {
	var sCount, lCount;
	for (var sCity in m_oCity) {
		sCount = (sCity == c_sDlvrCity) ? "lDelivery" : "lPickup";
		for (var lFSiteID in m_oCity[sCity].store) {
			// get the range count for each store
			lCount = m_oCity[sCity].store[lFSiteID].range_total();
			eval("m_" + sCount + "Count += " + lCount);
		}
	}
}
// create list of cities for given market region (jea:1/4/01)
// updates form object ---------------------------------------------------
function makeCityList() {
	var x = 1;
	m_oFldCity.options[0].value = -1;
	m_oFldCity.options[0].text = "--Select a City--";
	for (var sCity in m_oCity) {
		if (sCity != c_sDlvrCity) {
			// exclude delivery cities
			m_oFldCity.options[x] = new Option(sCity, sCity);
			x++;
		}
	}
}
// create list of stores for given city (jea:1/4/01)
// updates form object ---------------------------------------------------
function makeStoreList() {
	var x = 1;
	var sAddress;
	var oStore;
	clearStoreList();
	m_oFldStore.options[0].text = "--Select a Store--";
	for (m_lFSiteID in m_oCity[m_sCity].store) {
		oStore = m_oCity[m_sCity].store[m_lFSiteID];
		//sAddress = oStore.address + " (" + oStore.zip + ")";
		m_oFldStore.options[x] = new Option(oStore.address, m_lFSiteID);
		x++;
	}
	m_oFldStore.selectedIndex = 0;
	updateDelivery("");			// erase delivery information
	self.outerHeight -= 1;		// for ns to redraw resized fields
}
// create list of delivery days for given store (jea:1/9/01)
// updates form object ---------------------------------------------------
function makeDateList(v_lPickOrDlvr) {
	if (v_lPickOrDlvr == c_lPickup) {
		// use pickup store site id
		m_lFSiteID = m_oFldStore.options[m_oFldStore.selectedIndex].value;
		m_lAreaID = c_lPickup;	// constant is based on radio button index
	} else {
		// use customer's delivery site id
		m_lFSiteID = m_oCustomer.fsiteid;
		//m_lAreaID = getArea(m_oCustomer.zip);
		//alert(m_oCustomer.zip);
	}
	var x = 1;
	var oStore = m_oCity[m_sCity].store[m_lFSiteID];
	var oArea = oStore.area[m_lAreaID];
	var sDescription, sValue;
	
	clearDateList();
	m_oFldDate.options[0].text = "--Select a Day--";
	for (var sDay in oArea.day) {
		// create option for each day
		sDescription = oArea.day[sDay].description;
		if (oArea.day[sDay].order_count() > oStore.order_max) {
			// day is maxed out
			sDescription += ' (full)';
			sValue = 0;
		} else {
			sValue = sDay;
		}
		m_oFldDate.options[x] = new Option(sDescription, sValue)
		x++;
	}
	m_oFldDate.selectedIndex = 0;
	clearRangeList();
	if (v_lPickOrDlvr == c_lPickup) {
		// update fields with store information (jea:1/5/01)
		updateDelivery("store");
		self.outerHeight += 1;		// for ns to redraw resized fields
	}
}
// create list of delivery ranges for given day (jea:1/4/01)
// updates form object ---------------------------------------------------
function makeRangeList(v_lPickOrDlvr) {
	var x = 1;
	var oRange;
	m_sDay = m_oFldDate.options[m_oFldDate.selectedIndex].value;
	clearRangeList();
	m_oFldTime.options[0].text = "--Select a Time--";
	var oCat = m_oCity[m_sCity].store[m_lFSiteID].area[m_lAreaID].day[m_sDay].cat;
	var sDescription, sValue;
	for (var lCatID in oCat) {
		if (v_lPickOrDlvr == c_lDeliver) {
			// only display categories for delivery ranges, not pickup
			m_oFldTime.options[x] = new Option(m_sLine, 0); x++;
			m_oFldTime.options[x] = new Option(oCat[lCatID].description, 0); x++;
			m_oFldTime.options[x] = new Option(m_sLine, 0); x++;
		}
		for (var lRangeID in oCat[lCatID].range) {
			oRange = oCat[lCatID].range[lRangeID];
			sDescription = oRange.description;
			if (oRange.order_count > oRange.order_max) {
				// range is maxed out
				sDescription += ' (full)';
				sValue = 0;
			} else {
				sValue = lCatID + ',' + lRangeID;
			}
			m_oFldTime.options[x] = new Option(sDescription, sValue);
			x++;
		}
	}
	m_oFldTime.selectedIndex = 0;
	self.outerHeight -= 1;
	// may want to convert this to date readable by SQL (jea:1/9/00)
	m_oFrmMain.fldRangeDate.value = m_sDay;					// hidden field
}
// reset the store list (jea:2/14/01)
// updates form object ---------------------------------------------------
function clearStoreList() {
	if (m_oFldStore.length > 1) {
		m_oFldStore.length = 1;
		m_oFldStore.options[0].value = -1;
		m_oFldStore.options[0].text = "First Select a City";
		clearDateList();
	}
}
// reset the date list (jea:1/17/01)
// updates form object ---------------------------------------------------
function clearDateList() {
	if (m_oFldDate.length > 1) {
		m_oFldDate.length = 1;
		m_oFldDate.options[0].value = -1;
		m_oFldDate.options[0].text = "First Select a Store";
		clearRangeList();
	}
}
// reset the range list (jea:1/17/01)
// updates form object ---------------------------------------------------
function clearRangeList() {
	if (m_oFldTime.length > 1) {
		m_oFldTime.length = 1;
		m_oFldTime.options[0].value = -1;
		m_oFldTime.options[0].text = "First Select a Date";
	}
}
// populate delivery fields with specified data (jea:1/9/01)
// updates form objects --------------------------------------------------
function updateDelivery(v_sType) {
	switch (v_sType) {
		case "customer":
			updateDeliveryFields(m_oCustomer.firstname, m_oCustomer.mi, m_oCustomer.lastname,
				m_oCustomer.address1, m_oCustomer.address2, m_oCustomer.city, m_oCustomer.state,
				m_oCustomer.zip, m_oCustomer.zip4, m_oCustomer.phone);
			break;
		case "store":
			if (typeof(m_oCity[m_sCity]) == "object") {
				var oStore = m_oCity[m_sCity].store[m_lFSiteID];
				updateDeliveryFields("Albertsons", "", m_lFSiteID, oStore.address, "", m_sCity,
					oStore.state, oStore.zip, "", oStore.phone);
			}
			break;
		case "badzip":
			alert("Sorry, we cannot deliver to your registered shipping address.\nPlease enter a new address or choose a pickup store.");
			updateDeliveryFields(m_oCustomer.firstname, m_oCustomer.mi, m_oCustomer.lastname,
				"", "", "", "", "", "", m_oCustomer.phone);
			m_oFrmMain.fldShipAddress1.focus();
			break;
		default:
			updateDeliveryFields("", "", "", "", "", "", "", "", "", "");
			
	}
}
// populate delivery fields with given values (jea:1/5/01)
// updates form objects --------------------------------------------------
function updateDeliveryFields(v_sNameFirst, v_sNameMI, v_sNameLast, v_sAddress1, v_sAddress2,
	v_sCity, v_sStateCode, v_lZip, v_lZip4, v_sPhone) {
	
	m_oFrmMain.fldShipFirst.value = v_sNameFirst;
	m_oFrmMain.fldShipMI.value = v_sNameMI;
	m_oFrmMain.fldShipLast.value = v_sNameLast;
	m_oFrmMain.fldShipAddress1.value = v_sAddress1;
	m_oFrmMain.fldShipAddress2.value = v_sAddress2;
	m_oFrmMain.fldShipCity.value = v_sCity;
	m_oFrmMain.fldShipZip.value = v_lZip;
	m_oFrmMain.fldShipZip4.value = v_lZip4;
	m_oFrmMain.fldShipPhone.value = v_sPhone;
	selectState(v_sStateCode);
}

// Initialize form elements ==============================================

// initialize field objects for delivery (jea:1/10/01)
// updates form objects --------------------------------------------------
function initDeliver(r_oDeliver) {
	if (!(m_bInitDone)) {
		alert("Page has not finished initializing");
		// restore user's original selection:
		m_oFrmMain.fldPickOrDlvr[c_lPickup].checked = (m_lPickOrDlvr == c_lPickup);
		return;
	}
	if (m_lDeliveryCount == 0) {
		alert("Home delivery is not currently available in your area" + m_sJoke);
		r_oDeliver.checked = false;
		m_oFrmMain.fldPickOrDlvr[c_lPickup].checked = (m_lPickOrDlvr == c_lPickup);
		return false;
	}
	if (m_lPickOrDlvr != c_lPickup) { updateCustomer(); }
	if (m_sZips.search("," + m_oCustomer.zip + ",") == -1) {
		// customer zip is not valid for market region
		// mention option to pickup only if pickup is available
		alert("Delivery to your shipping address is not available.\nPlease enter another address" + ((m_lPickupCount != 0) ? " or choose from the pickup options." : ".") + m_sJoke);
		showNeither();
		return true;
	}
	m_sCity = c_sDlvrCity;			// city for delivery-only stores
	m_lFSiteID = m_oCustomer.fsiteid;
	m_lAreaID = getArea(m_oCustomer.zip);
	hideDivs([m_oDivPickup, m_oDivInstruct, m_oDivPickupTitle]);
	showDivs([m_oDivDeliver, m_oDivDeliverTitle]);
	m_oFldDate = m_oFrmDeliver.fldDeliverDate;
	m_oFldTime = m_oFrmDeliver.fldDeliverTime;
	// only pre-populate customer info if zip is in market
	updateDelivery((m_sZips.search("," + m_oCustomer.zip + ",") != -1) ? "customer" : "");
	disablePickupValidation();
	enableDeliveryValidation();
	m_lPickOrDlvr = c_lDeliver;
	return true;
}
// initialize field objects for pickup (jea:1/9/01)
// updates form objects --------------------------------------------------
function initPickup(r_oPickup) {
	if (!(m_bInitDone)) {
		alert("Page has not finished initializing");
		m_oFrmMain.fldPickOrDlvr[c_lDeliver].checked = (m_lPickOrDlvr == c_lDeliver);
		return;
	}
	if (m_lPickupCount == 0) {
		alert("There are currently no pickup stores available in your area" + m_sJoke);
		r_oPickup.checked = false;
		m_oFrmMain.fldPickOrDlvr[c_lDeliver].checked = (m_lPickOrDlvr == c_lDeliver);
		return false;
	} 
	hideDivs([m_oDivDeliver, m_oDivInstruct, m_oDivDeliverTitle]);
	showDivs([m_oDivPickup, m_oDivPickupTitle]);
	m_oFldDate = m_oFrmPickup.fldPickupDate;
	m_oFldTime = m_oFrmPickup.fldPickupTime;
	m_sCity = m_oFldCity.options[m_oFldCity.selectedIndex].value;
	m_lFSiteID = m_oFldStore.options[m_oFldStore.selectedIndex].value;
	updateCustomer();
	updateDelivery("");
	disableDeliveryValidation();
	enablePickupValidation();
	m_lPickOrDlvr = c_lPickup;
	return true;
}
// disable both pickup and delivery so user can enter valid address (jea:1/22/01)
// updaets form objects --------------------------------------------------
function initNeither() {
	m_sCity = c_sDlvrCity;			// city for delivery-only stores
	if (m_sZips.search("," + m_oCustomer.zip + ",") == -1) {
		// customer zip is not valid for market region
		updateDelivery("");
		m_oFrmMain.fldShipFirst.focus();
		m_lFSiteID = 0;
		m_lAreaID = "";
	} else {
		// display customer address
		updateDelivery("customer");
		m_lFSiteID = m_oCustomer.fsiteid;
		m_lAreaID = getArea(m_oCustomer.zip);
		makeDateList(c_lDeliver);
	}
	m_oFrmMain.fldPickOrDlvr[c_lPickup].checked = false;
	m_oFrmMain.fldPickOrDlvr[c_lDeliver].checked = false;
	hideDivs([m_oDivDeliver, m_oDivPickup, m_oDivPickupTitle]);
	showDivs([m_oDivInstruct, m_oDivDeliverTitle]);
}
// initialize field objects for credit card (jea:1/9/01)
// updates form objects --------------------------------------------------
function initCCN() {
	hideDivs([m_oDivPINumber, m_oDivPayment]);
	showDivs([m_oDivCCNumber]);
	enableCCValidation();
	disablePINValidation();
}
// initialize field objects for credit card (jea:1/9/01)
// updates form objects --------------------------------------------------
function initPIN() {
	hideDivs([m_oDivCCNumber, m_oDivPayment]);
	showDivs([m_oDivPINumber]);
	disableCCValidation();
	enablePINValidation();
}
// initialize common field objects (jea:1/9/01)
// updates form objects --------------------------------------------------
function initFields() {
	// field shortcuts abstracted for NS or IE
	m_oFldDate = m_oFrmDeliver.fldDeliverDate;
	m_oFldTime = m_oFrmDeliver.fldDeliverTime;
	m_oFldCity = m_oFrmPickup.fldCity;
	m_oFldStore = m_oFrmPickup.fldPickupStoreTemp;
}
// process city selection (jea:1/20/01)
// updates form objects --------------------------------------------------
function newCity(r_oFldCity) {
	m_sCity = r_oFldCity.options[r_oFldCity.selectedIndex].value;
	if (m_sCity < 1) {
		alert("Please select a city");
		r_oFldCity.selectedIndex = 0;
	} else if (m_oCity[m_sCity].range_total() > 0) {
		makeStoreList();
	} else {
		alert("Sorry, there are presently no delivery times available in " + m_sCity);
	}
}

// Process user selections ===============================================

// process store selection (jea:1/20/01)
// updates form objects --------------------------------------------------
function newStore(r_oFldStore) {
	m_lFSiteID = r_oFldStore.options[r_oFldStore.selectedIndex].value;
	//oFrmMain.fldPickupStore.value = m_lFSiteID;
	//oFrmMain.fldStore.value = m_oCity[m_sCity].store[m_lFSiteID].supplier;
	if (m_lFSiteID < 1) {
		updateDelivery("");
		alert("Please select a store");
		r_oFldStore.selectedIndex = 0;
	} else if (m_oCity[m_sCity].store[m_lFSiteID].range_total() > 0) {
		makeDateList(m_lPickOrDlvr);
	} else {
		alert("Sorry, there are no pickup times available\nfor the store at " + r_oFldStore.options[r_oFldStore.selectedIndex].text);
	}
}
// process date selection (jea:1/20/01)
// updates form objects --------------------------------------------------
function newDate(r_oFldDate) {
	sDate = r_oFldDate.options[r_oFldDate.selectedIndex].value;
	if (sDate < 1) {
		alert("Please select a day");
		r_oFldDate.selectedIndex = 0;
	} else if (m_oCity[m_sCity].store[m_lFSiteID].area[m_lAreaID].day[sDate].range_total() > 0) {
		makeRangeList(m_lPickOrDlvr);
	} else {
		alert("Sorry, no delivery times are available on this day");
	}
}
// process delivery / pickup window selection (jea:1/9/00)
// updates form objects --------------------------------------------------
function newRange(r_oFldRange) {
	// range selection contains both category and range id
	var sSelect = r_oFldRange.options[r_oFldRange.selectedIndex].value;
	if (sSelect < 1) {
		alert("Please select a " + ((m_lPickOrDlvr == c_lPickup) ? "pickup" : "delivery") + " option");
		r_oFldRange.selectedIndex = 0;
		return;
	}
	var aIndex = sSelect.split(',');
	var lCatID = aIndex[0];	m_lRangeID = aIndex[1];
	var oRange = m_oCity[m_sCity].store[m_lFSiteID].area[m_lAreaID].day[m_sDay].cat[lCatID].range[m_lRangeID];
	var lFeeThreshold = oRange.fee_threshold;
	var lFee = (m_oDataFrame.m_lSubTotal < lFeeThreshold || lFeeThreshold == 0) ? oRange.fee : 0;
	m_oFrmMain.fldRangeID.value = m_lRangeID;
	if (m_oDataFrame.m_lServiceFee != lFee || m_oDataFrame.m_lFeeThreshold != lFeeThreshold) {
		m_oDataFrame.m_lServiceFee = lFee;
		//oFrmMain.fldFee.value = oRange.fee;
		m_oDataFrame.m_lFeeThreshold = lFeeThreshold;
		//oFrmMain.fldFeeThreshold.value = oRange.fee_threshold;
		updatePost("newrange");
	}
}
// is state allowed in region and does it match zip code (jea:1/22/01)
// updates form objects --------------------------------------------------
function newState(r_oFldState) {
	var sStateCode = r_oFldState.options[r_oFldState.selectedIndex].value;
	if (typeof(m_oState[sStateCode]) == "object") {
		var sStateZips = "," + m_oState[sStateCode].zips.toString() + ",";
		var lZipCode = m_oFrmMain.fldShipZip.value;
		if (lZipCode != "") {
			// user has already entered a zip code--does it match?
			if (sStateZips.search("," + lZipCode + ",") == -1) {
				// erase unmatched zip code
				m_oFrmMain.fldShipZip.value = "";
			}
		}
	} else {
		// state is not allowed
		var sStateName = r_oFldState.options[r_oFldState.selectedIndex].text;
		alert("Shipping to " + sStateName + " is not available");
		r_oFldState.selectedIndex = 0;
	}
}
// check to see if new zip is allowed in market region (jea:1/10/01)
// updates form objects --------------------------------------------------
function newZip(r_oZip) {
	if (!(isPostal(r_oZip))) {
		// invalid zip code
		alert("Please enter a valid zip code");
		//oZip.focus();		// doesn't seem to be working
		document.frmCheckout.fldShipZip.focus();
		return;
	}
	if (m_sZips.search("," + r_oZip.value + ",") == -1) {
		// non-market region zip code
		alert("Based on your zip code, your address is not part of " +
			"the market region you selected at login" + m_sJoke);
		return;
	}
	selectState(findState(r_oZip.value));

	// update m_lFSiteID to match delivery area based on zip
	var sAreaZips;		// zips specific to delivery areas
	for (var lFSiteID in m_oCity[m_sCity].store) {
		for (var lAreaID in m_oCity[m_sCity].store[lFSiteID].area) {
			sAreaZips = "," + m_oCity[m_sCity].store[lFSiteID].area[lAreaID].zips.toString() + ",";
			if (sAreaZips.search("," + r_oZip.value + ",") != -1) {
				// entered zip matches this area
				m_oCustomer.fsiteid = lFSiteID;
				m_oCustomer.zip = r_oZip.value;
				if (lAreaID != m_lAreaID) {
					// refresh delivery range selections to display new options
					m_lAreaID = lAreaID;
					m_lFSiteID = lFSiteID;
					if (m_oFldDate.selectedIndex > 0) {
						// user had already made a selection
						alert("The address you have entered has new delivery options.\nYou will need to re-select the date and time.");
					}
					makeDateList(c_lDeliver);
				}
				return;
			}
		}
	}
}
// process payment type selection (jea:1/10/01)
// updates form and validation objects -----------------------------------
function newPayType(r_oPayType) {
	var sPayType = r_oPayType.options[r_oPayType.selectedIndex].value;
	if (sPayType == "1") {
		initCCN();
	} else if (sPayType == "3") {
		initPIN();
	}
}
// process PIN selection (jea:1/10/01)
// post form to totals frame for validation of PIN -----------------------
function newPIN(r_oPIN) {
	m_oFrmMain.fldPINumber.value = m_oFrmPINumber.fldPINumberTemp.value;
	updatePost("newpin");
}

// Update fields for checkout ============================================

// update customer object with changed field values (jea:1/9/01)
// updates data object ---------------------------------------------------
function updateCustomer() {
	m_oCustomer.firstname = m_oFrmMain.fldShipFirst.value;
	m_oCustomer.mi = m_oFrmMain.fldShipMI.value;
	m_oCustomer.lastname = m_oFrmMain.fldShipLast.value;
	m_oCustomer.address1 = m_oFrmMain.fldShipAddress1.value;
	m_oCustomer.address2 = m_oFrmMain.fldShipAddress2.value;
	m_oCustomer.city = m_oFrmMain.fldShipCity.value;
	m_oCustomer.state = m_oFrmMain.fldShipState.options[m_oFrmMain.fldShipState.selectedIndex].value;
	m_oCustomer.zip = m_oFrmMain.fldShipZip.value;
	m_oCustomer.zip4 = m_oFrmMain.fldShipZip4.value;
	m_oCustomer.phone = m_oFrmMain.fldShipPhone.value;
}
// update and validate fields and submit form (jea:1/9/00)
// updates objects and submits -------------------------------------------
function purchasePost() {
	if (m_bPosted) {
		alert("Your cart is already being processed.  There is no need to click again.");
		return;
	}
	if (!(m_bWait)) {
		// waiting for totals frame to validate
		if (m_oTotalFrame.m_bSuccess) {
			// validation successful
			updateFields();
			if (isValid(m_sForm, m_oFields)) {
				m_oFrmMain.target = 'cototal';
				m_oFrmMain.action = 'alb_checkout-total.asp?action=purchase';
				updateFieldsForPickup(false);
				m_bPosted = true;
				top.message.startBar('Processing your order');
				submit(m_oFrmMain);
				updateFieldsForPickup(true);
			}
		} else {
			m_oTotalFrame.m_bSuccess = true;
		}
	} else {
		setTimeout("purchasePost()", 200);
	}
}
// submit form for totals update (jea:1/10/01)
// updates objects and submits -------------------------------------------
function updatePost(v_sAction) {
	m_bWait = true;
	updateFields();
	m_oFrmMain.target = 'cototal';
	m_oFrmMain.action = 'alb_checkout-total.asp?action=' + v_sAction;
	submit(m_oFrmMain);
}
// populate hidden fields
// updates form objects --------------------------------------------------
function updateFields() {
	// clean values
	var sMonth = m_oFrmCCNumber.fldCCMonth.options[m_oFrmCCNumber.fldCCMonth.selectedIndex].value;
	var sYear = m_oFrmCCNumber.fldCCYear.options[m_oFrmCCNumber.fldCCYear.selectedIndex].value;
	with (m_oFrmMain) {
		fldCCExpireMonth.value = sMonth;
		fldCCExpireYear.value = sYear;
		var sInstructions = (m_lPickOrDlvr == c_lPickup) ? m_oFrmPickup.fldPickupInstructions.value : m_oFrmDeliver.fldDeliverInstructions.value;
		fldInstructions.value = sInstructions.substr(0,249);	// only allow 250 characters
		// assign values from layered forms
		fldCCExpire.value = sMonth + "/1/" + sYear;
		fldCCType.value = m_oFrmCCNumber.fldCCTypeTemp.options[m_oFrmCCNumber.fldCCTypeTemp.selectedIndex].value;
		fldCCNumber.value = m_oFrmCCNumber.fldCCNumberTemp.value;
		fldPINumber.value = m_oFrmPINumber.fldPINumberTemp.value;
		fldRecipient.value = m_oFrmPickup.fldRecipientTemp.value;
		fldSubTotal.value = m_oDataFrame.m_lSubTotal;
		// assign values data from variables
		fldStore.value = m_lFSiteID;
		if (m_lPickOrDlvr == c_lPickup) {
			fldStore.value = m_oCity[m_sCity].store[m_lFSiteID].supplier;
			fldPickupStore.value = m_lFSiteID;
		} else {
			fldStore.value = m_lFSiteID;
		}
		fldFee.value = m_oDataFrame.m_lServiceFee;
		fldTax.value = m_oDataFrame.m_lTax;
		fldDiscount.value = m_oDataFrame.m_lDiscount;
		fldTotal.value = m_oDataFrame.m_lTotal;
		fldOrderID.value = m_oDataFrame.m_lOrderID;
		fldGiftAmount.value = m_oDataFrame.m_lGiftAmount;
		fldGiftBalance.value = m_oDataFrame.m_lGiftBalance;
		fldVerisignID.value = m_oDataFrame.m_lVerisignID;
		fldEmail.value = m_oCustomer.email;
	}
}
// populate pickup store id at last second so user doesn't see (jea:1/31/01)
// updates form objects --------------------------------------------------
function updateFieldsForPickup(v_bClear) {
	if (m_lPickOrDlvr == c_lPickup) {
		if (v_bClear) {
			m_oFrmMain.fldShipAddress2.value = '';
			m_oFrmMain.fldShipLast.value = '';
		} else {
			var lStoreID = m_oCity[m_sCity].store[m_lFSiteID].storeid;
			m_oFrmMain.fldShipAddress2.value = lStoreID;
			m_oFrmMain.fldShipLast.value = '- Store #: ' + lStoreID;
		}
	}
}
// match state to zip code (jea:1/22/01)
// returns boolean -------------------------------------------------------
function isZipInState(v_lZipCode) {
	var sStateCode = m_oFrmMain.fldShipState.options[m_oFrmMain.fldShipState.selectedIndex].value;
	var sStateZips = "," + m_oState[sStateCode].zips.toString() + ",";
	return (sStateZips.search("," + v_lZipCode + ",") != -1) ? true : false;
}
// find state with given zip code (jea:1/22/01)
// returns string --------------------------------------------------------
function findState(v_lZipCode) {
	// assumes valid zip code
	var sStateZips;
	for (var sStateCode in m_oState) {
		sStateZips = "," + m_oState[sStateCode].zips.toString() + ",";
		if (sStateZips.search("," + v_lZipCode + ",") != -1) {
			return sStateCode;
		}
	}
}
// select given state in option list (jea:1/22/01)
// updates form object ---------------------------------------------------
function selectState(v_sStateCode) {
	var oFldState = m_oFrmMain.fldShipState;
	var re = /^[a-zA-Z]{2}$/;
	if (re.test(v_sStateCode)) {
		// process valid state code
		// loop through state list and select given state
		for (var x = 0; x < oFldState.options.length; x++) {
			if (oFldState.options[x].value == v_sStateCode) {
				oFldState.options[x].selected = true;
				return;	// our work is done here
			}
		}		
	} else {
		oFldState.options[0].selected = true;
	}
}
// don't try to stop progress bar until finished loading (jea:2/9/01)
// updates objects -------------------------------------------------------
function stopProgressBar() {
	var bBarReady = m_oMsgFrame.m_bBarReady;
	if (typeof(bBarReady) == "boolean") {
		if (bBarReady) { m_oMsgFrame.stopBar('Checkout is ready'); return; }
	}
	setTimeout("stopProgressBar()", 200);
}
