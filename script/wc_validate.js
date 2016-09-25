// field object constructor (updated 7/28/00)
function field(v_desc, v_type, v_req) {
	this.desc = v_desc;
	this.type = v_type;
	this.req = v_req;
}
// master validation function (updated 7/27/00)
function isValid(v_sForm, r_oFields) {
	var oForm = eval("document." + v_sForm);
	var oErrors = new Array
	var lCount = 0;
	
	for (var i in r_oFields) {
		// pass every field to the appropriate validation function
		var r_oField = eval("oForm." + i);
		if (r_oFields[i].req || r_oField.value.length > 0) {
			// only check fields that have a value or are required
			if (!(eval("is" + r_oFields[i].type + "(oForm, r_oField)"))) {
				oErrors[lCount] = r_oFields[i].desc;
				lCount++;
			}
		}
	}
	if (lCount > 0) {
		// must be some errors
		var sMessage = "Please enter valid\n";
		for (i = 0; i < lCount; i++) {
			sMessage += "- " + oErrors[i] + "\n";
		}
		alert(sMessage);
		return false;
	} else {
		return true;
	}
}
// check for empty strings (updated 7/27/00)
function isString(r_oForm, r_oField) {
	return (r_oField.value == "") ? false : true;
}
// check for valid e-mail (updated 7/27/00)
function isEmail(r_oForm, r_oField) {
	var re = /^(\w+@\w+\.\w+)$/;
	return re.test(r_oField.value);
}
// confirm that field is all numerics (updated 7/20/00)
function isNumeric(r_oForm, r_oField) {
	var re = /^(\d+)$/;
	return re.test(r_oField.value);
}
// make sure they made some selection (updated 7/27/00)
// this assumes that layout options, like lines, are <= 0
function isSelect(r_oForm, r_oField) {
	var re = /[a-zA-Z:]/;	// \D gives bad result
	var val = r_oField.options[r_oField.selectedIndex].value;
	// true if option value > 0 or non-numeric
	return (val > 0 || re.test(val)) ? true : false;
}
// make sure one radio button was checked (updated 7/27/00)
function isRadio(r_oForm, r_oField) {
	for (var i = 0; i < r_oField.length; i++) {
		// cycle through each item in the radio collection
		if (r_oField[i].checked) { return true; }
	}
	// if we made it here then no radio is checked
	return false;
}
// make sure that a date has been entered (updated 7/27/00)
// allows dashes or slashes, m/d/yy or mm/dd/yyyy
function isDate(r_oForm, r_oField) {
	var re = /^((\d{1,2})[\/-\\](\d{1,2})[\/-\\](\d{2,4}))$/;
	if (re.test(r_oField.value)) {
		// format is right--now check each date value
		var arMatch = re.exec(r_oField.value);
		var iMonth = arMatch[2];
		var iDay = arMatch[3];
		var iYear = cleanYear(arMatch[4]);
		
		if (iMonth <= 12 && iMonth >= 1 && iDay <= 31 && iDay >= 1 && iYear <= 2010 && iYear >= 1850) {
			return true;
		}
	} 
	// invalid date format
	return false;
}
// make years four digits (updated 8/23/00)
function cleanYear(v_sYear) {
	if (v_sYear.length == 2) {
		// assume century break on xx10
		sYear = (v_sYear > 10 ? "19" : "20") + v_sYear;
	}
	return v_sYear;
}
// validate scope settings for event (updated 2/21/01)
// hardcoded for values on event edit page
function isScoped(r_oForm, r_oField) {
	var bVisible = false;
	for (var lGroupID in m_oGroup) {
		if (m_oGroup[lGroupID].scope != 0) { bVisible = true; break; }
	}
	return bVisible;
}
// validate pasword (updated 2/24/01)
//   assumes field 'fldConfirm' in same form
function isPassword(r_oForm, r_oField) {
	if (isString(r_oForm, r_oField)) {
		if (r_oField.value == r_oForm.fldConfirm.value) { return true; }
	}
	return false;
}
// converts a field to all numbers (updated 7/27/00)
function toNumeric(v_sField) {
	var sNum = "";
	for (var i = 0; i < v_sField.length; i++) {
		var sChar = v_sField.substring(i, i + 1);
		if (sChar >= "0" && sChar <= "9") {
			sNum += sChar;
		}
	}
	return sNum;
}