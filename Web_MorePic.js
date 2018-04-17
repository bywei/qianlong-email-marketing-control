 function getAbsolutePos(el){
	var SL = 0, ST = 0;
	var is_div = /^div$/i.test(el.tagName);
	if (is_div && el.scrollLeft)
		SL = el.scrollLeft;
	if (is_div && el.scrollTop)
		ST = el.scrollTop;
	var r = { x: el.offsetLeft-SL, y: el.offsetTop - ST };
	if (el.offsetParent)
	{
		var tmp = this.getAbsolutePos(el.offsetParent);
		r.x += tmp.x;
		r.y += tmp.y;
	}
	return r;
};
function img$mousedown(e){
	if(!e)e=window.event;
	var div, img;
	var d = document;
	img = e.srcElement?e.srcElement:e.target;
	var len;
	div = document.getElementById("processbarDiv");
	len = div.offsetWidth-img.offsetWidth;
	var o = getAbsolutePos(div);
	var minX = o.x, maxX = o.x + div.offsetWidth - img.offsetWidth;
	var minY = o.y, maxY = o.y + div.offsetHeight - img.offsetHeight;
	
	var x=e.layerX?e.layerX:e.offsetX, y=e.layerY?e.layerY:e.offsetY;
	function img$mousemove(e){
		if(!e)e=window.event;
		if(!e.pageX)e.pageX=e.clientX+d.body.scrollLeft;
		if(!e.pageY)e.pageY=e.clientY+d.body.scrollTop;
		var tx=e.pageX-minX-x,ty=e.pageY-minY-y;
		img.style.left = (tx>len?len:(tx<0?0:tx))+"px";
		img.pNum = tx>len?len:(tx<0?0:tx);
		
		var oImg = document.getElementById('Image1');
        oImg.style.zoom = (img.pNum*100/len).toString()  + '%';
	};
	function img$mouseup(){
		if(img.releaseCapture)
		{
			img.releaseCapture();
			d.detachEvent("onmousemove", img$mousemove);
			d.detachEvent("onmouseup", img$mouseup);
		}
		else if(img.removeEventListener)
		{
			d.removeEventListener("mousemove", img$mousemove, true);
			d.removeEventListener("mouseup", img$mouseup, true);
		}
	};
	if(img.attachEvent)
	{
		img.setCapture();
		d.attachEvent("onmousemove", img$mousemove);
		d.attachEvent("onmouseup", img$mouseup);
	}
	else if(img.addEventListener)
	{
		d.addEventListener("mousemove", img$mousemove, true);
		d.addEventListener("mouseup", img$mouseup, true);
	}
};
if(window.attachEvent){
    document.getElementById("processbarDiv").getElementsByTagName("img")[0].attachEvent("onmousedown", img$mousedown);
}
else if(window.addEventListener){
	document.getElementById("processbarDiv").getElementsByTagName("img")[0].addEventListener("mousedown", img$mousedown, false);

}        
var _shoeImages = new Array();
  function colorchanged(ulist)
{
	var   sarray=new   Array();   
	var sarray=ulist.split("|");
	var jj=sarray.length ;
	 for(i=0;i<=jj;i++) 
	{
             _shoeImages[i] = sarray[i];
	}
            indexChanged(0);
		
 }
   function indexChanged(index)
  {
         document.getElementById("Image1").src = ""+_shoeImages[index];
 }
  function loadEvent()
  {
    document.getElementById("DataList1").getElementsByTagName("img")[0].click();
   }
  if(window.attachEvent)
{
   window.attachEvent("onload", loadEvent);
   }
 else if(window.addEventListener)
 window.addEventListener("load", loadEvent, false);  