<!--#include file="Conn.asp"-->
<!--#include file="CF_MyFunction.asp"-->
<!--#include file="CF_Md5.asp"-->
<%
Action=Request("Action")
%>
<!--#Include File="CF_Do.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3c.org/TR/1999/REC-html401-19991224/loose.dtd">
<!-- saved from url=(0036)http://www.linezing.com/register.php -->
<HTML xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<META content="IE=7.0000" 
http-equiv="X-UA-Compatible">
<TITLE><%=RsSet("SysTitle")%>-找回密码</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<LINK 
rel=stylesheet type=text/css href="images/style.css">
<LINK 
rel="shortcut icon" href="image_new/favicon.ico">
<SCRIPT language=javascript> 
function correctPNG()  
{ 
	for(var i=0; i<document.images.length; i++) 
	{ 
		var img = document.images[i] 
			var imgName = img.src.toUpperCase() 
			if (imgName.substring(imgName.length-3, imgName.length) == "PNG") 
			{ 
				var imgID = (img.id) ? "id='" + img.id + "' " : "" 
					var imgClass = (img.className) ? "class='" + img.className + "' " : "" 
					var imgTitle = (img.title) ? "title='" + img.title + "' " : "title='" + img.alt + "' " 
					var imgStyle = "display:inline-block;" + img.style.cssText  
					if (img.align == "left") imgStyle = "float:left;" + imgStyle 
						if (img.align == "right") imgStyle = "float:right;" + imgStyle 
							if (img.parentElement.href) imgStyle = "cursor:hand;" + imgStyle   
								var strNewHTML = "<span " + imgID + imgClass + imgTitle 
									+ " style=\"" + "width:" + img.width + "px; height:" + img.height + "px;" + imgStyle + ";" 
									+ "filter:progid:DXImageTransform.Microsoft.AlphaImageLoader" 
									+ "(src=\'" + img.src + "\', sizingMethod='scale');\"></span>"  
									img.outerHTML = strNewHTML 
									i = i-1 
			} 
	} 
} 
function mbStringLength(s) {
	var totalLength = 0;
	var i;
	var charCode;
	for (i = 0; i < s.length; i++) {
		charCode = s.charCodeAt(i);
		if (charCode < 0x007f) {
			totalLength = totalLength + 1;
		} else if ((0x0080 <= charCode) && (charCode <= 0x07ff)) {
			totalLength += 2;
		} else if ((0x0800 <= charCode) && (charCode <= 0xffff)) {
			totalLength += 3;
		}
	}
	//alert(totalLength);
	return totalLength;
}

function checkPwdCode(s)
{
	var i;
	var charCode;
	for (i = 0; i < s.length; i++)
	{
		charCode = s.charCodeAt(i);
		if (charCode < 33 || charCode > 126)
		{
			return false;
		}
	}
	return true;
}

function chgImg(){
	document.getElementById("vcode_img").src="vcode.php?t="+Math.random();
	return false;
}

if(navigator.userAgent.indexOf("MSIE")>-1)
window.attachEvent("onload", correctPNG); 

</SCRIPT>
<META name=GENERATOR content="MSHTML 8.00.6001.19154">
</HEAD>
<BODY>
<DIV id=container>
  <DIV id=main>
    <DIV id=header>
      <DIV id=logo><A href="index.asp"><IMG title=到邮件营销网站首页 
border=0 src="images/logo.png"></A></DIV>
      <SCRIPT language=javascript> 
			function AddFavorite(sURL, sTitle) 
			{ 
				try 
				{ 
					window.external.addFavorite(sURL, sTitle); 
				} 
				catch (e) 
				{ 
					try 
					{ 
						window.sidebar.addPanel(sTitle, sURL, ""); 
					} 
					catch (e) 
					{ 
						alert("请点击浏览器收藏夹进行添加"); 
					} 
				} 
			} 
			if(navigator.userAgent.indexOf("MSIE")>-1)
			window.attachEvent("onload", correctPNG); 
		</SCRIPT>
      <DIV id=user_info><A 
href="index.asp">登录</A>&nbsp;&nbsp;|&nbsp;&nbsp; <A 
href="#">论坛</A>&nbsp;&nbsp;|&nbsp;&nbsp; <A 
href="#">帮助</A>&nbsp;&nbsp;|&nbsp;&nbsp; <IMG border=0 align=absMiddle src="images/fav.gif"><A 
onclick=AddFavorite(window.location,document.title) 
href="http://www.qianlongsoft.com/">收藏本站</A></IMG> </DIV>
    </DIV>
    <DIV id=block></DIV>
    <DIV id=page>
      <DIV id=bug_tips></DIV>
      <DIV id=txzcxx style="height:300px;">
         <!--#include file="CF_PwdRecover_Pre_2.asp"-->
      </DIV>
     
    </DIV>
    <DIV id=page_end></DIV>
    <DIV id=footer>
      <DIV id=info><A href="#" target=_blank>关于我们</A>&nbsp;&nbsp;|&nbsp;&nbsp; <A 
href="#" 
target=_blank>人才招聘<A>&nbsp;&nbsp;|&nbsp;&nbsp; <A 
href="#" target=_blank>服务条款</A> </DIV>
      <DIV id=copyright>CopyRight &copy; QianLongSoft 2009</DIV>
    </DIV>
  </DIV>
</DIV>

</BODY>
</HTML>
<%Call ConnClose()%>