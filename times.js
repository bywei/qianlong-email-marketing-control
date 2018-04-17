//	Written by cloudchen, 2004/03/15
function minute(name,fName) {
	this.name = name;
	this.fName = fName || "m_input";
	this.timer = null;
	this.fObj = null;
	
	this.toString = function() {
		var objDate = new Date();
		var sMinute_Common = "class=\"m_input\" maxlength=\"2\" name=\""+this.fName+"\" onfocus=\""+this.name+".setFocusObj(this)\" onblur=\""+this.name+".setTime(this)\" onkeyup=\""+this.name+".prevent(this)\" onkeypress=\"if (!/[0-9]/.test(String.fromCharCode(event.keyCode)))event.keyCode=0\" onpaste=\"return false\" ondragenter=\"return false\" style=\"ime-mode:disabled\"";
		var sButton_Common = "class=\"m_arrow\" onfocus=\"this.blur()\" onmouseup=\""+this.name+".controlTime()\" disabled"
		var str = "";
		//str += "<div class=\"m_frameborder\">"
		str += "<input  value=\""+this.formatTime(objDate.getHours())+"\" "+sMinute_Common+">:"
		str += "<input  value=\""+this.formatTime(objDate.getMinutes())+"\" "+sMinute_Common+">:"
		str += "<input  value=\""+this.formatTime(objDate.getSeconds())+"\" "+sMinute_Common+""
		//str += "</div>"
		return str;
	}
	this.play = function() {
		this.timer = setInterval(this.name+".playback()",1000);
	}
	this.formatTime = function(sTime) {
		sTime = ("0"+sTime);
		return sTime.substr(sTime.length-2);
	}
	this.playback = function() {
		var objDate = new Date();
		var arrDate = [objDate.getHours(),objDate.getMinutes(),objDate.getSeconds()];
		var objMinute = document.getElementsByName(this.fName);
		for (var i=0;i<objMinute.length;i++) {
			objMinute[i].value = this.formatTime(arrDate[i])
		}
	}
	this.prevent = function(obj) {
		clearInterval(this.timer);
		this.setFocusObj(obj);
		var value = parseInt(obj.value,10);
		var radix = parseInt(obj.radix,10)-1;
		if (obj.value>radix||obj.value<0) {
			obj.value = obj.value.substr(0,1);
		}
	}
	this.controlTime = function(cmd) {
		event.cancelBubble = true;
		if (!this.fObj) return;
		clearInterval(this.timer);
		var cmd = event.srcElement.innerText=="5"?true:false;
		var i = parseInt(this.fObj.value,10);
		var radix = parseInt(this.fObj.radix,10)-1;
		if (i==radix&&cmd) {
			i = 0;
		} else if (i==0&&!cmd) {
			i = radix;
		} else {
			cmd?i++:i--;
		}
		this.fObj.value = this.formatTime(i);
		this.fObj.select();
	}
	this.setTime = function(obj) {
		obj.value = this.formatTime(obj.value);
	}
	this.setFocusObj = function(obj) {
		eval(this.fName+"_up").disabled = eval(this.fName+"_down").disabled = false;
		this.fObj = obj;
	}
	this.getTime = function() {
		var arrTime = new Array(2);
		for (var i=0;i<document.getElementsByName(this.fName).length;i++) {
			arrTime[i] = document.getElementsByName(this.fName)[i].value;
		}
		return arrTime.join(":")
	}
}
var m = new minute("m");
document.write(m);
m.play();