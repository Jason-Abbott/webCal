// Copyright 2000 Jason Abbott (webcal@webott.com)
// Last updated 3/30/2000

// toggle visibility of help division

var ie4 = document.all;
var ns4 = document.layers;
var divHelp;

if (ie4) {
	divHelp = Help.style;
	divHelp.display = "none";
} else if (ns4) {
	divHelp = document.layers["Help"];
	divHelp.visibility = "hide";
}
function viewHelp() {
	if (ie4) {
		divHelp.display = (divHelp.display == "block") ? "none" : "block"; }
	else if (ns4) {
		divHelp.visibility = (divHelp.visibility == "show") ? "hide" : "show"; }
}