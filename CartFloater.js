self.onError=null;
cartcurrentX = 0;
cartcurrentY =400;  
cartwhichIt = null;           
cartlastScrollX = 0; cartlastScrollY =-50;
cartNS = (document.layers) ? 1 : 0;
cartIE = (document.all) ? 1: 0;
function cartheartBeat() {
	if(cartIE) { diffY = document.body.scrollTop; diffX = document.body.scrollLeft; }
	if(cartNS) { diffY = self.pageYOffset; diffX = self.pageXOffset; }
	if(diffY != cartlastScrollY) {
	percent = .1 * (diffY - cartlastScrollY);
	if(percent > 0) percent = Math.ceil(percent);
	                else percent = Math.floor(percent);
			if(cartIE) document.all.CartFloater.style.pixelTop += percent;
			if(cartNS) document.CartFloater.top += percent; 
	                cartlastScrollY = cartlastScrollY + percent;
                  	}
	
	}
	function cartcheckFocus(x,y) { 
	        stalkerx = document.CartFloater.pageX;
	        stalkery = document.CartFloater.pageY;
	        stalkerwidth = document.CartFloater.clip.width;
	        stalkerheight = document.CartFloater.clip.height;
	        if( (x > stalkerx && x < (stalkerx+stalkerwidth)) && (y > stalkery && y < (stalkery+stalkerheight))) return true;
	        else return false;
	}
	function cartgrabIt(e) {
		if(cartIE) {
		cartwhichIt = event.srcElement;
		while (cartwhichIt.id.indexOf("CartFloater") == -1) {
			cartwhichIt = cartwhichIt.parentElement;
			if (cartwhichIt == null) { return true; }
		}

		cartwhichIt.style.pixelTop = cartwhichIt.offsetTop;
	   		cartcurrentY = (event.clientY + document.body.scrollTop); 	
		}
	        if(cartcheckFocus (e.pageX,e.pageY)) { 
	                cartwhichIt = document.CartFloater;
	                StalkerTouchedX = e.pageX-document.CartFloater.pageX;
	                StalkerTouchedY = e.pageY-document.CartFloater.pageY;
	        } else { 
	        window.captureEvents(Event.MOUSEMOVE);
		}
	    return true;
                }
document.onmousedown = "cartgrabIt";
action = window.setInterval("cartheartBeat()",1);    