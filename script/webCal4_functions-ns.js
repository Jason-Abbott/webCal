function buttonOn(v_sDivID) {
	document.layers['div' + v_sDivID].document.bgColor = 'ff0000';
}
function buttonOff(v_sDivID) {
	document.layers['div' + v_sDivID].document.bgColor = 'ffffff';
}
function showSettings() {
	oDivSettings = document.layers['divSettings'];
	oDivSettings.visibility = 'show';
}

function hideSettings() {
	oDivSettings = document.layers['divSettings'];
	oDivSettings.visibility = 'hide';
}