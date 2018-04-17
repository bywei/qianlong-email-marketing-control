<!--#Include File="Conn.asp"-->
<!--#Include File="CF_MyFunCtion.asp"-->
<!--#Include File="CF_Md5.asp"-->
<!--#Include File="CF_Getcookie.asp"-->
<%
Action=Request("Action")
%>
<!--#Include File="CF_Do.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<link href="css/index_style.css" rel="stylesheet" type="text/css" />
<link rel="shortcut icon" href="images/favicon.ico" />
<title><%=RsSet("SysTitle")%> - 专业电子邮件营销服务 稳定 精准 高效！</title>
<meta name="keywords" content="电子邮件,邮件跟踪查询系统,邮件营销,邮件分析系统,邮件格式,邮件群发,邮件群发软件,电子邮件大全,qq邮件,什么是电子邮件,怎么发邮件"/>
<meta http-equiv="description" content="电子邮件营销数据跟踪分析系统，专业的邮件跟踪查询系统。结合邮件群发软件实现实时统计。怎么发邮件?如何获取电子邮件大全?免费提供十万邮件列表地址！400-6677-270"/>
<!--png去背景-->
<script language="javascript"> 			
function correctPNG(){for(var i=0; i<document.images.length; i++) { 				
var img = document.images[i] 		
var imgName = img.src.toUpperCase() 		
if (imgName.substring(imgName.length-3, imgName.length) == "PNG") 			
{ 						var imgID = (img.id) ? "id='" + img.id + "' " : "" 	
var imgClass = (img.className) ? "class='" + img.className + "' " : "" 			
var imgTitle = (img.title) ? "title='" + img.title + "' " : "title='" + img.alt + "' " 	
var imgStyle = "display:inline-block;" + img.style.cssText  			
if (img.align == "left") imgStyle = "float:left;" + imgStyle 			
if (img.align == "right") imgStyle = "float:right;" + imgStyle 			
if (img.parentElement.href) imgStyle = "cursor:hand;" + imgStyle   	
var strNewHTML = "<span " + imgID + imgClass + imgTitle + " style=\"" + "width:" + img.width + "px; height:" + img.height + "px;" + imgStyle + ";" 							+ "filter:progid:DXImageTransform.Microsoft.AlphaImageLoader" 							+ "(src=\'" + img.src + "\', sizingMethod='scale');\"></span>" 
img.outerHTML = strNewHTML 	
i = i-1 				
} 				} 			} 	
function mbStringLength(s) {var totalLength = 0;
var i;
var charCode;for (i = 0; i < s.length; i++) {charCode = s.charCodeAt(i);if (charCode < 0x007f) {totalLength = totalLength + 1;						} else if ((0x0080 <= charCode) && (charCode <= 0x07ff)) {totalLength += 2;} else if ((0x0800 <= charCode) && (charCode <= 0xffff)) {						totalLength += 3;}				}				return totalLength;			}			function mbStringLength(s) {				var totalLength = 0;				var i;				var charCode;				for (i = 0; i < s.length; i++) {					charCode = s.charCodeAt(i);					if (charCode < 0x007f) {						totalLength = totalLength + 1;						} else if ((0x0080 <= charCode) && (charCode <= 0x07ff)) {						totalLength += 2;						} else if ((0x0800 <= charCode) && (charCode <= 0xffff)) {						totalLength += 3;					}				}						
return totalLength;			}			function AddFavorite(sURL, sTitle) 			{ 				try 				{ 					window.external.addFavorite(sURL, sTitle); 				} 				catch (e) 				{ 					try 					{ 						window.sidebar.addPanel(sTitle, sURL, ""); 					} 					catch (e) 					{ 						alert("请点击浏览器收藏夹进行添加"); 					} 				} 			} 			if(navigator.userAgent.indexOf("MSIE")>-1)			window.attachEvent("onload", correctPNG); 		
<!--png去背景-->
function ccimgchange()
{
 document.getElementById("ccimg").src = 'CF_CheckCode.asp?a=' + Math.random();
}
</script>
</head>
<body>
<div id="container_index">
  <div id="body_index">
    <div id="header_index">
      <div id="logo_index"> <a href="/article"><img src="images/logo.png" border=0 title="到钱龙统计首页"/></a> </div>
      <div id="user_info_index"> <a href='#' style="color:#ffffff; text-decoration:none;">论坛</a>&nbsp;&nbsp;|&nbsp;&nbsp; <a href="/article" style="color:#ffffff; text-decoration:none;">帮助</a>&nbsp;&nbsp;|&nbsp;&nbsp; <img src="images/fav.gif" align="absmiddle" border="0"><a href='#' onclick="AddFavorite(window.location,document.title)" style="color:#ffffff; text-decoration:none;">收藏本站</a></img> </div>
    </div>
    <div id="flash_login_index">
      <div id="flash_index"> <a target='_blank' href="/article"><img border=0 src="images/banner520.jpg"/></a> </div>
      <form action="Index.asp?Action=userloginsave" method="post" name="logindata" onsubmit="return (regcheck(this));">
        <input type="hidden" name="referer" value="" />
        <input type="hidden" name="webname" value="index" />
        <div id="login_index">
          <div id="one">
            <%
If Session("CFCountUser")="" Then
%>
            <div id="id"> 用户名：
              <input name="username" id="username" type="text" value="" style="width:160px;height:16px">
            </div>
            <div id="error" style="display:none" > <font color="#FF0000"><span id="username_tag"></span></font> </div>
            <div id="password"> 密&nbsp;&nbsp;&nbsp;&nbsp;码：
              <input name="pwd" id="pwd" type="password" value="" style="width:160px;height:16px" onpaste="return false;">
            </div>
            <div id="error" style="display:none"> <font color="#FF0000"><span id="password_tag"></span></font> </div>
            <div id="validecode"> 验证 码：
              <INPUT id=checkcode type=text size=6 name=checkcode>
              <a href="javascript:ccimgchange()" title="看不清楚？换一个！"><IMG id="ccimg" height=14 src="CF_CheckCode.asp" border="0"></a> </div>
            <div id="remember">
              <label for="cookies_time">Cookie：</label>
              <select name="cookies_time" id="cookies_time" style="width:115px;">
                <option value="60" selected>不保留</option>
                <option value="1440">保留一天</option>
                <option value="43200">保留一个月</option>
                <option value="5256000">永久保留</option>
              </select>
              <a href="PwdRecover.asp">取回密码</a> </div>
            <div id="login_button">
              <input name="submit" type="submit" value="登录" id="login_button_index">
            </div>
            <%Else%>
            <div style="text-align:center; margin-top:50px;">
              <INPUT onclick="location.href='Manage.asp';" type='button' value='进入管理' name='submit' class='button' style="margin-top:10px">
              <INPUT onclick="location.href='Manage.asp?action=userlogout';" type='button' value='退出管理' name='submit' class='button'  style="margin-top:10px">
            </div>
            <%End if%>
          </div>
          <div id="two">
            <!-- disnable register-->
            <div id="new_user"> 快来拥有您自己的钱龙统计账号~ </div>
            <div id="free_register" style="float: left; width:260px; margin-left:0px;">
              <div style="float: left; width:130px"> <a style="margin-left:20px;" href="reg.asp"><img border="0" src="images/free_register.gif" name="image"></a></div>
              <div style="float: left; width:130px;"> <a style="margin-left:10px;" href="/article"><img border="0" src="images/View_demo.gif" name="image"></a></div>
            </div>
            <!-- disnable register -->
          </div>
          <script language="javascript">								regminpwd = 6;								regmaxpwd = 16;								document.logindata.pwd.value = '';								function regcheck(formct){									if (!checkpwd())									{										formct.password.focus();										return false;									}									return true;								}								function CheckUserNameFunc()								{									max_username_len = 20;									min_username_len = 5;									var input = document.logindata.username;									var input_len = mbStringLength(input.value);									var tip = document.getElementById('username_tag');									if ('' == input.value)									{										tip.innerHTML = "用户名不能为空";										tip.parentNode.parentNode.style.display = "block";										input.focus();										return false;									}									if (max_username_len < input_len || min_username_len > input_len)									{										tip.innerHTML = "用户名长度应在5-20间";										tip.parentNode.parentNode.style.display = "block";										input.focus();										return false;									}									if (false)									{										tip.innerHTML = "用户名只能使用提示的字符";										tip.parentNode.parentNode.style.display = "block";										input.focus();										return false;									}									tip.innerHTML = '';									tip.parentNode.parentNode.style.display = "none";									return true;								}								function checkPwdCode(s) {									var i;									var charCode;									for (i = 0; i < s.length; i++)									{										charCode = s.charCodeAt(i);										if (charCode < 33 || charCode > 126)										{											return false;										}									}									return true;								}									function checkpwd(type){									var tip = document.getElementById('password_tag');									var pwd = document.logindata.pwd.value;									var pwd_len = mbStringLength(pwd);									var msg = '';									var ret = false;									if (pwd == '')									{										msg = "密码不能为空";									}									else if (pwd_len<regminpwd)									{										msg = "密码最小值不能小于"+regminpwd;										} else if (regmaxpwd>0 && pwd_len>regmaxpwd) {										msg = "密码最大值不能大于"+regmaxpwd;										} else if (!checkPwdCode(pwd)) {										msg = "密码只能使用英文字母、数字和标点";									}									if(msg == '')	{										tip.parentNode.parentNode.style.display = "none";										}else{										tip.innerHTML = msg;										tip.parentNode.parentNode.style.display = "block";									}									return ('' === msg );								}								document.logindata.username.focus();							</script>
        </div>
      </form>
    </div>
    <div id="flash_login_shadow"> </div>
    <div id="content_index">
      <div id="our_services">
        <div class="title_index"> <img src="images/linezing_xx.jpg"> </div>
        <div class="text_area_index"> <strong>钱龙统计是专业的邮件营销流量统计工具。<br />
          目前提供以下服务：</strong>
          <p />
          <img src="images/xjj.jpg"><strong>邮件营销统计</strong><br />
          .&nbsp;&nbsp;&nbsp;钱龙统计邮件营销统计秉承百味统计的服务理念，为用户提供精确、稳定的邮件营销流量统计服务。【<a target="_blank" href="http://www.bywei.cn/">点击查看更多</a>】
          <p />
          <img src="images/xjj.jpg"><strong><a target='_blank' href="http://www.bywei.cn">钱龙邮件营销</a></strong><br />
          .&nbsp;&nbsp;&nbsp;钱龙邮件营销统计是为钱龙客户量身打造的专业邮件营销数据统计系统，每分钟更新访问数据，实时跟踪在线访客，提供最实时、最精准的数据服务，为邮件营销推广和商品展示提供充分依据。【<a target="_blank" href="http://www.bywei.cn">点击查看更多</a>】
          <p />
          <br />
        </div>
      </div>
      <div id="hot_faq">
        <div class="title_index_x"> <a href="/article" title="去论坛"><img src="images/linezing_xxx.jpg" border="0"></a> </div>
        <div class="text_area_index"> <img src="images/xjj.jpg"><a target="_blank" href="/article">钱龙统计新增"页面点击热图"统计功能--招募测试用户！</a><br />
          <img src="images/xjj.jpg"><a target="_blank" href="/article">冲PV、比成交，现金红包天天送！</a><br />
          <img src="images/xjj.jpg"><a target="_blank" href="/article">雅虎统计迁移公告</a><br />
          <img src="images/xjj.jpg"><a target="_blank" href="/article">钱龙统计论坛上线公告</a><br />
          <img src="images/xjj.jpg"><a target="_blank" href="/article">关于钱龙统计的相关解释和说明</a> <br />
          <br />
        </div>
        <div class="title_index_xxxx"> <a href="/article" title="去帮助中心"><img src="images/linezing_x.jpg" border="0"></a> </div>
        <div class="text_area_index"> <img src="images/xjj.jpg"><a target="_blank" href="/article">雅虎统计向钱龙统计迁移的FAQ</a><br />
          <img src="images/xjj.jpg"><a target="_blank" href="/article">从雅虎统计迁移到钱龙统计的步骤</a><br />
          <img src="images/xjj.jpg"><a target="_blank" href="/article">什么是迁移过渡期</a><br />
          【<a target="_blank" href="/article">查看更多问题</a>】<br />
        </div>
      </div>
      <div id="try">
        <div class="title_index_xxx"></div>
        <div class="text_area_short_index"> <a target='_blank' href="/article"><img src="images/T1b5SuXmVVXXXXXXXX-223-129.jpg" width="223" height="129" border="0" /></a> <br>
          <br>
          <a target='_blank' href="/article"><img src="images/paidai.jpg" border="0" /></a> <br>
          <br>
          2010-12-14<br>
          <a target='_blank' href="/article">2011钱龙邮件营销效果（钱龙官方版）正式推出！ </a> </div>
      </div>
    </div>
    <div id="content_footer_index"> </div>
    <div id="footer_index">
      <div id="help_index"><a target="_blank" href="/article">关于我们</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a target="_blank" href="/article">人才招聘</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="/article" target="_blank">服务条款</a></div>
      <div id="copyright_index">CopyRight &copy; QianLongSoft 2009</div>
    </div>
  </div>
</div>
</body>
</html>
<!-- web1 -->