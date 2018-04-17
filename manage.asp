<!--#Include File="Conn.asp"-->
<!--#Include File="CF_MyFunCtion.asp"-->
<!--#Include File="CF_Md5.asp"-->
<!--#Include File="CF_Getcookie.asp"-->
<!--#Include File="CF_Manage_Pre.asp"-->
<!--#Include File="CF_Manage_Do.asp"-->
<!--#Include File="inc.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link rel="stylesheet" href="images/alimm000.css" type="text/css" media="screen" title="no title" charset="utf-8">
<link rel="stylesheet" href="images/ystats00.css" type="text/css" media="screen" title="no title" charset="utf-8">
<title>综合报告 -<%=RsSet("SysTitle")%></title>
<!--链接效果控制-->
<style type="text/css">
<!--
#control a {
	text-decoration:none;
	color:#0162F6;
}
#info a {
	text-decoration: underline;
	color:#ffffff;
}
#aboutus a {
	text-decoration: underline;
	color:#ffffff;
}
#copyright {
	color:#ffffff;
}
-->
</style>
<script language="javascript"> 
function submit_frm(selItem){var p = selItem;p.ei.value = "UTF-8";p.method = "get";return true;}
var flag=0;
function selid(){
    if(!flag){
        document.getElementById("slist").style.display="block";
        flag = 1;
    }else{
        document.getElementById("slist").style.display="none";
        flag = 0;
    }
};

function hlit(obj){obj.style.borderColor = "#829EDD";}
function delit(obj){obj.style.borderColor = "#ccc";}

//document.onclick = function(e){var e = e || window.event;var el = (e.target)?((e.target.nodeValue == 3)?e.target.parentNode : e.target) : e.srcElement;if(el.id != 'selid'){document.getElementById("slist").style.display="none";flag = 0;}};

function AddFavorite(sURL, sTitle) { 
	try { 
		window.external.addFavorite(sURL, sTitle); 
	} catch (e) { 
		try { 
			window.sidebar.addPanel(sTitle, sURL, ""); 
		} catch (e) { 
			alert("请点击浏览器收藏夹进行添加"); 
		} 
	} 
} 
</script>
<script src="images/changeSt.htm"></script>
</head>
<body style="margin:0; padding:0; background-color:#4a4b4f; font-family:Arial, Helvetica, sans-serif;" <%If Action="jsqset" Then Response.Write " onload="&show%>>
<!--最外部div-->
<div id="container" style="width:100%; height:100%; background:url(images/header_bg.jpg); background-repeat:repeat-x;">
  <!--页头div-->
  <div id="header" style="width:960px; height:75px; margin:0 auto; font-size:12px;">
    <!--logo-->
    <div id="logo" style="width:221px; height:49px; margin-top:20px; float:left;"> <a href='#'><img border=0 src='images/logo0000.png' /></a> </a> </div>
    <!--登录信息-->
    <div id="info" style="color:#ffffff; width:500px; height:20px; margin-top:20px; float:right;text-align:right"> 您好，<a href='Manage.asp'><%=UserName%></a>&nbsp;&nbsp;|&nbsp;<a href='#' target='_blank'>论坛</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="#" target='_blank'>帮助</a>&nbsp;&nbsp;|&nbsp;&nbsp;<img src="images/fav00000.gif" align="absmiddle" border="0"><a href='' onClick="AddFavorite(window.location,document.title)">收藏本站</a></img>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="?Action=userlogout">退出</a> </div>
  </div>
  <!--导入部分div-->
  <div id="import" style="width:960px; margin:0 auto; background-color:#FFFFFF">
    <div id="main">
      <div id="where" class="wherecss">
        <div class="current"> <a name="1">
          <label>当前用户：</label>
          </a>
          <div class="sites"><strong><%=UserName%></strong></div>
          <div id="selid" onClick="selid()">
            <div id="slist" style="z-index:9999;">
              <div id="slist">
                <ul>
                  <li><a target="_parent" href="#"><%=UserName%></a></li>
                </ul>
              </div>
            </div>
            <img style="height:16px; width:16px;" align="absmiddle" src="images/tj_00100.gif"> 选择 </div>
        </div>
        <div style="float:right"> <span style="display:block;float:left;font-size:13px;padding-top:11px;padding-left:10px">换个颜色? </span>
          <ul id="changeStyleList" style="display:block;float:left;height:20px;line-height:15px;margin-top:7px;padding-left:10px;width:90px;">
            <li color="grey" onClick="styleBtnClick('grey')" class="changeStyleLi" title="默认"><img src="images/grey00.gif"></li>
            <li color="orange" onClick="styleBtnClick('orange')" class="changeStyleLi" title="橙色"><img src="images/orange00.gif"></li>
            <li color="blue" onClick="styleBtnClick('blue')" class="changeStyleLi" title="蓝色"><img src="images/blue00.gif"></li>
            <li color="purple" onClick="styleBtnClick('purple')" class="changeStyleLi" title="紫色"><img src="images/purple00.gif"></li>
          </ul>
        </div>
        <script>
	(function() {
		var cur_color = getCookie("_ystat_style");
		if(cur_color!="" && cur_color!="grey") {
			styleBtnClick(cur_color);
		}
	}());
</script>
        <div class="back"> <a href="Manage.asp">返回统计首页</a></div>
        <div class="clearing"></div>
      </div>
      <div class="content">
        <!--left-->
        <!--#Include File="CF_Manage_Menu.asp"-->
        <!--left end-->
        <!--right-->
        <div class="col">
          <div class="current">
            <ul class="path">
              <li class="title">当前位置：</li>
              <li><a href="manage.asp"><%=RsSet("SysTitle")%></a></li>
              <li class="arrow">></li>
              <li>综合报告</li>
            </ul>
          </div>
          <div style="float:right">
            <p > <a target="_blank" style="color:red;text-decoration:underline" href="#">【<%=RsSet("SysTitle")%>】官方权威数据分析工具，此刻激活还可免费试用！</a><img src="images/new00000.gif" border="0"> </p>
          </div>
          <!-- ********************************* current link end ********************************* -->
          <div class="clearing"></div>
          <!--#Include File="CF_Manage_Select.asp"-->
          <p class="back-top"><a href="">返回顶部</a></p>
        </div>
        <!--reight end-->
      </div>
    </div>
    <br>
    <br>
  </div>
  <!--页脚div-->
  <div id="footer" style="width:960px; height:80px; margin:0 auto; font-size:12px;">
    <!--about us div-->
    <div id="aboutus" style="width:250px; height:20px; float:left;"> <a target="_blank" href="#">关于我们</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a target="_blank" href="#">人才招聘</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a target="_blank"  href="#">服务条约</a> </div>
    <!--(R) div-->
    <div id="copyright" style="width:180px; height:20px; float:right;"> CopyRight &copy; Qianlongsoft 2009 </div>
  </div>
</div>
</body>
<span id="slist"></span>
</html>
