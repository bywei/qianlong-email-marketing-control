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
<title><%=RsSet("SysTitle")%> - רҵ�����ʼ�Ӫ������ �ȶ� ��׼ ��Ч��</title>
<meta name="keywords" content="�����ʼ�,�ʼ����ٲ�ѯϵͳ,�ʼ�Ӫ��,�ʼ�����ϵͳ,�ʼ���ʽ,�ʼ�Ⱥ��,�ʼ�Ⱥ�����,�����ʼ���ȫ,qq�ʼ�,ʲô�ǵ����ʼ�,��ô���ʼ�"/>
<meta http-equiv="description" content="�����ʼ�Ӫ�����ݸ��ٷ���ϵͳ��רҵ���ʼ����ٲ�ѯϵͳ������ʼ�Ⱥ�����ʵ��ʵʱͳ�ơ���ô���ʼ�?��λ�ȡ�����ʼ���ȫ?����ṩʮ���ʼ��б��ַ��400-6677-270"/>
<!--pngȥ����-->
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
return totalLength;			}			function AddFavorite(sURL, sTitle) 			{ 				try 				{ 					window.external.addFavorite(sURL, sTitle); 				} 				catch (e) 				{ 					try 					{ 						window.sidebar.addPanel(sTitle, sURL, ""); 					} 					catch (e) 					{ 						alert("����������ղؼн������"); 					} 				} 			} 			if(navigator.userAgent.indexOf("MSIE")>-1)			window.attachEvent("onload", correctPNG); 		
<!--pngȥ����-->
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
      <div id="logo_index"> <a href="/article"><img src="images/logo.png" border=0 title="��Ǯ��ͳ����ҳ"/></a> </div>
      <div id="user_info_index"> <a href='#' style="color:#ffffff; text-decoration:none;">��̳</a>&nbsp;&nbsp;|&nbsp;&nbsp; <a href="/article" style="color:#ffffff; text-decoration:none;">����</a>&nbsp;&nbsp;|&nbsp;&nbsp; <img src="images/fav.gif" align="absmiddle" border="0"><a href='#' onclick="AddFavorite(window.location,document.title)" style="color:#ffffff; text-decoration:none;">�ղر�վ</a></img> </div>
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
            <div id="id"> �û�����
              <input name="username" id="username" type="text" value="" style="width:160px;height:16px">
            </div>
            <div id="error" style="display:none" > <font color="#FF0000"><span id="username_tag"></span></font> </div>
            <div id="password"> ��&nbsp;&nbsp;&nbsp;&nbsp;�룺
              <input name="pwd" id="pwd" type="password" value="" style="width:160px;height:16px" onpaste="return false;">
            </div>
            <div id="error" style="display:none"> <font color="#FF0000"><span id="password_tag"></span></font> </div>
            <div id="validecode"> ��֤ �룺
              <INPUT id=checkcode type=text size=6 name=checkcode>
              <a href="javascript:ccimgchange()" title="�����������һ����"><IMG id="ccimg" height=14 src="CF_CheckCode.asp" border="0"></a> </div>
            <div id="remember">
              <label for="cookies_time">Cookie��</label>
              <select name="cookies_time" id="cookies_time" style="width:115px;">
                <option value="60" selected>������</option>
                <option value="1440">����һ��</option>
                <option value="43200">����һ����</option>
                <option value="5256000">���ñ���</option>
              </select>
              <a href="PwdRecover.asp">ȡ������</a> </div>
            <div id="login_button">
              <input name="submit" type="submit" value="��¼" id="login_button_index">
            </div>
            <%Else%>
            <div style="text-align:center; margin-top:50px;">
              <INPUT onclick="location.href='Manage.asp';" type='button' value='�������' name='submit' class='button' style="margin-top:10px">
              <INPUT onclick="location.href='Manage.asp?action=userlogout';" type='button' value='�˳�����' name='submit' class='button'  style="margin-top:10px">
            </div>
            <%End if%>
          </div>
          <div id="two">
            <!-- disnable register-->
            <div id="new_user"> ����ӵ�����Լ���Ǯ��ͳ���˺�~ </div>
            <div id="free_register" style="float: left; width:260px; margin-left:0px;">
              <div style="float: left; width:130px"> <a style="margin-left:20px;" href="reg.asp"><img border="0" src="images/free_register.gif" name="image"></a></div>
              <div style="float: left; width:130px;"> <a style="margin-left:10px;" href="/article"><img border="0" src="images/View_demo.gif" name="image"></a></div>
            </div>
            <!-- disnable register -->
          </div>
          <script language="javascript">								regminpwd = 6;								regmaxpwd = 16;								document.logindata.pwd.value = '';								function regcheck(formct){									if (!checkpwd())									{										formct.password.focus();										return false;									}									return true;								}								function CheckUserNameFunc()								{									max_username_len = 20;									min_username_len = 5;									var input = document.logindata.username;									var input_len = mbStringLength(input.value);									var tip = document.getElementById('username_tag');									if ('' == input.value)									{										tip.innerHTML = "�û�������Ϊ��";										tip.parentNode.parentNode.style.display = "block";										input.focus();										return false;									}									if (max_username_len < input_len || min_username_len > input_len)									{										tip.innerHTML = "�û�������Ӧ��5-20��";										tip.parentNode.parentNode.style.display = "block";										input.focus();										return false;									}									if (false)									{										tip.innerHTML = "�û���ֻ��ʹ����ʾ���ַ�";										tip.parentNode.parentNode.style.display = "block";										input.focus();										return false;									}									tip.innerHTML = '';									tip.parentNode.parentNode.style.display = "none";									return true;								}								function checkPwdCode(s) {									var i;									var charCode;									for (i = 0; i < s.length; i++)									{										charCode = s.charCodeAt(i);										if (charCode < 33 || charCode > 126)										{											return false;										}									}									return true;								}									function checkpwd(type){									var tip = document.getElementById('password_tag');									var pwd = document.logindata.pwd.value;									var pwd_len = mbStringLength(pwd);									var msg = '';									var ret = false;									if (pwd == '')									{										msg = "���벻��Ϊ��";									}									else if (pwd_len<regminpwd)									{										msg = "������Сֵ����С��"+regminpwd;										} else if (regmaxpwd>0 && pwd_len>regmaxpwd) {										msg = "�������ֵ���ܴ���"+regmaxpwd;										} else if (!checkPwdCode(pwd)) {										msg = "����ֻ��ʹ��Ӣ����ĸ�����ֺͱ��";									}									if(msg == '')	{										tip.parentNode.parentNode.style.display = "none";										}else{										tip.innerHTML = msg;										tip.parentNode.parentNode.style.display = "block";									}									return ('' === msg );								}								document.logindata.username.focus();							</script>
        </div>
      </form>
    </div>
    <div id="flash_login_shadow"> </div>
    <div id="content_index">
      <div id="our_services">
        <div class="title_index"> <img src="images/linezing_xx.jpg"> </div>
        <div class="text_area_index"> <strong>Ǯ��ͳ����רҵ���ʼ�Ӫ������ͳ�ƹ��ߡ�<br />
          Ŀǰ�ṩ���·���</strong>
          <p />
          <img src="images/xjj.jpg"><strong>�ʼ�Ӫ��ͳ��</strong><br />
          .&nbsp;&nbsp;&nbsp;Ǯ��ͳ���ʼ�Ӫ��ͳ�Ʊ��а�ζͳ�Ƶķ������Ϊ�û��ṩ��ȷ���ȶ����ʼ�Ӫ������ͳ�Ʒ��񡣡�<a target="_blank" href="http://www.bywei.cn/">����鿴����</a>��
          <p />
          <img src="images/xjj.jpg"><strong><a target='_blank' href="http://www.bywei.cn">Ǯ���ʼ�Ӫ��</a></strong><br />
          .&nbsp;&nbsp;&nbsp;Ǯ���ʼ�Ӫ��ͳ����ΪǮ���ͻ���������רҵ�ʼ�Ӫ������ͳ��ϵͳ��ÿ���Ӹ��·������ݣ�ʵʱ�������߷ÿͣ��ṩ��ʵʱ���׼�����ݷ���Ϊ�ʼ�Ӫ���ƹ����Ʒչʾ�ṩ������ݡ���<a target="_blank" href="http://www.bywei.cn">����鿴����</a>��
          <p />
          <br />
        </div>
      </div>
      <div id="hot_faq">
        <div class="title_index_x"> <a href="/article" title="ȥ��̳"><img src="images/linezing_xxx.jpg" border="0"></a> </div>
        <div class="text_area_index"> <img src="images/xjj.jpg"><a target="_blank" href="/article">Ǯ��ͳ������"ҳ������ͼ"ͳ�ƹ���--��ļ�����û���</a><br />
          <img src="images/xjj.jpg"><a target="_blank" href="/article">��PV���ȳɽ����ֽ��������ͣ�</a><br />
          <img src="images/xjj.jpg"><a target="_blank" href="/article">�Ż�ͳ��Ǩ�ƹ���</a><br />
          <img src="images/xjj.jpg"><a target="_blank" href="/article">Ǯ��ͳ����̳���߹���</a><br />
          <img src="images/xjj.jpg"><a target="_blank" href="/article">����Ǯ��ͳ�Ƶ���ؽ��ͺ�˵��</a> <br />
          <br />
        </div>
        <div class="title_index_xxxx"> <a href="/article" title="ȥ��������"><img src="images/linezing_x.jpg" border="0"></a> </div>
        <div class="text_area_index"> <img src="images/xjj.jpg"><a target="_blank" href="/article">�Ż�ͳ����Ǯ��ͳ��Ǩ�Ƶ�FAQ</a><br />
          <img src="images/xjj.jpg"><a target="_blank" href="/article">���Ż�ͳ��Ǩ�Ƶ�Ǯ��ͳ�ƵĲ���</a><br />
          <img src="images/xjj.jpg"><a target="_blank" href="/article">ʲô��Ǩ�ƹ�����</a><br />
          ��<a target="_blank" href="/article">�鿴��������</a>��<br />
        </div>
      </div>
      <div id="try">
        <div class="title_index_xxx"></div>
        <div class="text_area_short_index"> <a target='_blank' href="/article"><img src="images/T1b5SuXmVVXXXXXXXX-223-129.jpg" width="223" height="129" border="0" /></a> <br>
          <br>
          <a target='_blank' href="/article"><img src="images/paidai.jpg" border="0" /></a> <br>
          <br>
          2010-12-14<br>
          <a target='_blank' href="/article">2011Ǯ���ʼ�Ӫ��Ч����Ǯ���ٷ��棩��ʽ�Ƴ��� </a> </div>
      </div>
    </div>
    <div id="content_footer_index"> </div>
    <div id="footer_index">
      <div id="help_index"><a target="_blank" href="/article">��������</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a target="_blank" href="/article">�˲���Ƹ</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="/article" target="_blank">��������</a></div>
      <div id="copyright_index">CopyRight &copy; QianLongSoft 2009</div>
    </div>
  </div>
</div>
</body>
</html>
<!-- web1 -->