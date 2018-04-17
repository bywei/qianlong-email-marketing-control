self.onError=null;
currentX = 0;
currentY =400;  
whichIt = null;           
lastScrollX = 0; lastScrollY =-150;
NS = (document.layers) ? 1 : 0;
IE = (document.all) ? 1: 0;
function heartBeat() {
	if(IE) { diffY = document.body.scrollTop; diffX = document.body.scrollLeft; }
	if(NS) { diffY = self.pageYOffset; diffX = self.pageXOffset; }
	if(diffY != lastScrollY) {
	percent = .1 * (diffY - lastScrollY);
	if(percent > 0) percent = Math.ceil(percent);
	                else percent = Math.floor(percent);
			if(IE) document.all.floater.style.pixelTop += percent;
			if(NS) document.floater.top += percent; 
	                lastScrollY = lastScrollY + percent;
                  	}
	
	}
	function checkFocus(x,y) { 
	        stalkerx = document.floater.pageX;
	        stalkery = document.floater.pageY;
	        stalkerwidth = document.floater.clip.width;
	        stalkerheight = document.floater.clip.height;
	        if( (x > stalkerx && x < (stalkerx+stalkerwidth)) && (y > stalkery && y < (stalkery+stalkerheight))) return true;
	        else return false;
	}
	function grabIt(e) {
		if(IE) {
		whichIt = event.srcElement;
		while (whichIt.id.indexOf("floater") == -1) {
			whichIt = whichIt.parentElement;
			if (whichIt == null) { return true; }
		}

		whichIt.style.pixelTop = whichIt.offsetTop;
	   		currentY = (event.clientY + document.body.scrollTop); 	
		}
	        if(checkFocus (e.pageX,e.pageY)) { 
	                whichIt = document.floater;
	                StalkerTouchedX = e.pageX-document.floater.pageX;
	                StalkerTouchedY = e.pageY-document.floater.pageY;
	        } else { 
	        window.captureEvents(Event.MOUSEMOVE);
		}
	    return true;
                }
document.onmousedown = "grabIt";
action = window.setInterval("heartBeat()",1);    