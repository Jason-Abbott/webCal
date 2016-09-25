// Copyright 2000 Jason Abbott (webcal@webott.com)
// Last updated 4/20/2000

// show requested help text (updated 1/4/2001)
// pops alert ------------------------------------------------------------
function showHelp(v_strHelp) {
	alert(oHelp[v_strHelp]);
}

var m_oHelp = new Object;

m_oHelp["month"] = 
"To Login\n  Click the key icon\n\n" +
"To Add an Event\n  Click the day\n\n" +
"To View a week at at time\n  Click on the week tabs on the left";