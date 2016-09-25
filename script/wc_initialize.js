var m_isIE = document.all;
var m_isNS = document.layers;
var m_isMac = (navigator.appVersion.indexOf("Mac") != -1);
var m_isWin = (navigator.appVersion.indexOf("Win") != -1);

function initCalendar(v_sView, v_sDate) {
	var sPage;
	var sFonts;

	sFonts = (isFontInstalled("Webdings")) ? "1" : "0";
	sFonts += "," + ((isFontInstalled("Wingdings")) ? "1" : "0") + ",0"
	sPage = "wc_" + v_sView + ".asp?date=" + v_sDate + "&fs=" + sFonts;

	location.replace(sPage);
}

// http://www.webreference.com/dhtml/column30/5.html
function isFontInstalled(v_sFont) {
	var isVer4 = (m_isIE || m_isNS );
	if (!isVer4) { return false; }
	var arNotNavWin = ["Webdings","Marlett"];

	if (m_isNS && m_isWin) {
		for (var i = 0; i < arNotNavWin.length; i++) {
			if (v_sFont == arNotNavWin[i]) { return false; }
		}
	}
	var sTest = "font&nbsp;existence&nbsp;test";
	var sLayer0, sLayer1, oLayer0, oLayer1, llWidth0, llWidth1;
	if (m_isIE) {
		if (!window.oLayer0) {
			sLayer0 = "<SPAN ID=oLayer0 STYLE='position:absolute;visibility:hidden;width:30;font:12pt Courier'>"+ sTest +"</SPAN>";
			sLayer1 = "<SPAN ID=oLayer1 STYLE='position:absolute;visibility:hidden;width:30;font-size:12pt'>"+ sTest +"</SPAN>";
			document.body.insertAdjacentHTML("BeforeEnd", sLayer0);		
			document.body.insertAdjacentHTML("BeforeEnd", sLayer1);
		}
		window.oLayer1.style.fontFamily = v_sFont + ",Courier";
		lWidth0 = (m_isMac) ? oLayer0.offsetWidth : window.oLayer0.scrollWidth;
		lWidth1 = (m_isMac) ? oLayer1.offsetWidth : window.oLayer1.scrollWidth;
	}
	if (m_isNS) {
		sLayer1 = "<FONT FACE='"+ v_sFont +",Courier' POINT-SIZE=12>"+ sTest +"</FONT>";
		if(!window.oLayer0) {
			// create new layer
			sLayer0 = "<FONT FACE='Courier' POINT-SIZE=12>"+ sTest +"</FONT>";
			oLayer0 = new Layer(400);
			oLayer0.document.write(sLayer0);
			oLayer0.document.close();
			oLayer1 = new Layer(400);
			oLayer1.document.write(sLayer1);
			oLayer1.document.close();
		} else {
			// write to existing layer
			oLayer1.document.write(sLayer1);
			oLayer1.document.close();
		}
		lWidth0 = oLayer0.clip.width;
		lWidth1 = oLayer1.clip.width;
	}
	return (lWidth0 != lWidth1);
}