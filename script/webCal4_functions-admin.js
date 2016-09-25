// Copyright 2001 Jason Abbott (webcal@webott.com)
// Last updated 2/24/2001

var m_oFields = { 
	fldNameFirst:{desc:"First Name",type:"String",req:1},
	fldNameLast:{desc:"Last Name",type:"String",req:1},
	fldEmail:{desc:"E-mail Address",type:"Email",req:0},
	fldPassword:{desc:"Password (they don't match)",type:"Password",req:1}
	};

var m_oForm;

function initPage() {
	m_oForm = document.frmEdit;
	showMessage();
	
	if (document.images) {
		// back to calendar icon
		var iconMonth = new Image();
		iconMonth.src = "images/icon_calprev_grey.gif";
		var iconMonthOn = new Image();
		iconMonthOn.src = "images/icon_calprev.gif"
		statusMonth = "Return to calendar";
	}	
}

function iconOver(name){
	if (document.images) {
  		document.images[name].src=eval("icon"+name+"On.src");
		status=eval("status"+name);
	}
}

function iconOut(name){
	if (document.images) {
  		document.images[name].src=eval("icon"+name+".src");
		status="";
	}
}