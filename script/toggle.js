// define values used by toggle function
var sHide; var sShow; var sElement; var sMethod;
if (document.all) {
	// Probably IE browser
	sElement = "document.all.";
	sMethod = ".style.display";
	sShow = "block";
	sHide = "none";
} else if (document.layers) {
	// Probably Netscape
	sElement = "document.layers['"
	sMethod = "'].visibility";
	sShow = "show";
	sHide = "hide";
}

// toggle tag visibility (jea:8/7/00)
// updates tag style property --------------------------------------------
function toggle(sTag) {
	sAction = sElement + sTag + sMethod;
	var sVis = (eval(sAction) == sHide) ? sShow : sHide;
	eval(sAction + "='" + sVis + "'");
}

// indicates whether a tag is currently visible (jea:8/23/00)
// return tag style property ---------------------------------------------
function isVisible(sTag) {
	return (eval(sElement + sTag + sMethod) == sShow);
}