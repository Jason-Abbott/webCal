<SCRIPT LANGUAGE=JScript RUNAT=Server>
function y2k(number) {
	return (number < 1000) ? number + 1900 : number;
}

function milliDif() {
	var d = new Date();
	return d.getTime()
}
                  
function elapsedpretty(parm1) {
	var elapsedsecs = 0
	var elapsedmins = 0
  
	elapsedsecs=Math.floor(parm1/1000)
	parm1=parm1%1000
  
	elapsedmins=Math.floor(elapsedsecs/60)
	elapsedsecs=elapsedsecs%60
	elapsedpretty=elapsedmins + " minute"

	if(elapsedmins!=1) {
		elapsedpretty=elapsedpretty+"s"
	}
	elapsedpretty = elapsedpretty+" " + elapsedsecs+" second"

	if(elapsedsecs!=1) {
		elapsedpretty=elapsedpretty+"s"
	}
	elapsedpretty = elapsedpretty+ " "+parm1+" millisecond"

	if(parm1!=1) {
		elapsedpretty=elapsedpretty+"s"
	}
  
	return elapsedpretty;
}
</script>
