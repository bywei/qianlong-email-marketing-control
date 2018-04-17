document.write("<SCRIPT FOR='DoubleFloat' EVENT='fscommand()' LANGUAGE='JavaScript'>");
document.write("LayerLeft.style.visibility='hidden';");
document.write("LayerRight.style.visibility='hidden';");
document.write("</script>");
function StartDoubleFloat() {
document.all.LayerLeft.style.posTop = -200;
document.all.LayerLeft.style.visibility = 'visible'
document.all.LayerRight.style.posTop = -200;
document.all.LayerRight.style.visibility = 'visible'
MoveLeftLayer('LayerLeft');
MoveRightLayer('LayerRight');
}
function MoveLeftLayer(layerName) {
var x = 5;
var y = 50;
var diff = (document.body.scrollTop + y - document.all.LayerLeft.style.posTop)*.40;
var y = document.body.scrollTop + y - diff;
eval("document.all." + layerName + ".style.posTop = y");
eval("document.all." + layerName + ".style.posLeft = x");
setTimeout("MoveLeftLayer('LayerLeft');", 20);
}
function MoveRightLayer(layerName) {
var x = 5;
var y = 50;
var diff = (document.body.scrollTop + y - document.all.LayerRight.style.posTop)*.40;
var y = document.body.scrollTop + y - diff;
eval("document.all." + layerName + ".style.posTop = y");
eval("document.all." + layerName + ".style.posRight = x");
setTimeout("MoveRightLayer('LayerRight');", 20);
}
function close()
{
	document.getElementById("floatDiv1").style.display="none";
	document.getElementById("floatDiv2").style.display="none";
}
document.write("<div id=LayerLeft style='position: absolute;visibility:hidden;z-index:1'><div id='floatDiv1'><div style='position:absolute;left:0;padding:2px;background-color:#FFFFFF;'><a href='javascript:close();'><img src='images/close1.gif' border='0' alt='关闭'></a></div><a href='"+lurl+"' target=_blank><img src='images/"+lpic+"' width=90 height=360 id=DoubleFloat border=0></a></div></div>"+"<div id=LayerRight style='position: absolute;visibility:hidden;z-index:1'><div id='floatDiv2'><div style='position:absolute;right:0;padding:2px;background-color:#FFFFFF;'><a href='javascript:close();'><img src='images/close1.gif' border='0' alt='关闭'></a></div><a href='"+rurl+"' target=_blank><img src='images/"+rpic+"' width=90 height=360 id=DoubleFloat border=0></a></div></div>");
StartDoubleFloat();
