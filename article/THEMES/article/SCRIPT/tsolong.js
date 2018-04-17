//为了实 "用户无需修改代码即可正常应用该主题" 所以加了一个js,否则用户需要自行修改blog的导航链接
document.getElementById("divNavBar").getElementsByTagName("li")[0].className="current";
var nav=document.getElementById("divNavBar").getElementsByTagName("a");
for(var i=0;i<nav.length;i++){
	nav[i].innerHTML="<b>"+nav[i].innerHTML+"</b>";
}