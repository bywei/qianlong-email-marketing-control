
<%If Action="" Then%>
<%If RsSet("RegState")=-1 Then%>
<script>
function usercheck() {
var username = document.getElementById("username").value;
if(username == "")
{
 usernametext.innerHTML="<br>用户名不能为空";
 return false;
}

var xmlDom = jb(); 
var strData = "code=123";
var url = "?action=userchecksave&username="+username;
xmlDom.open("POST", url, false); 
xmlDom.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
xmlDom.send(strData);
var webcode="<br>"+xmlDom.responseText

usernametext.innerHTML=webcode;
}

function jb() { 
 var A=null; 
 try { 
 A=new ActiveXObject("Msxml2.XMLHTTP");
 } 
 catch(e) { 
 try { 
 A=new ActiveXObject("Microsoft.XMLHTTP");
 } 
 catch(oc) { 
 A=null;
 } 
 } 

 if ( !A && typeof XMLHttpRequest != "undefined" ) { 
 A=new XMLHttpRequest();
 } 
 return A;
}
</script>


<script language="JavaScript">

function ccimgchange()
{
 document.getElementById("ccimg").src = 'CF_CheckCode.asp?a=' + Math.random();
}

function chkname(String)
{ 
var Letters = "abcdefghijklmnopqrstuvwxyz1234567890_"; //可以自己增加可输入值
var i;
var c;
if(String.charAt( 0 )=='-')
return false;
if( String.charAt( String.length - 1 ) == '-' )
return false;
for( i = 0; i < String.length; i ++ )
{
c = String.charAt( i );
if (Letters.indexOf( c ) < 0)
return false;
}
return true;
}

function userregcheck(){
if ((document.form2.username.value)==""){
 window.alert ('用户名必须填写!');
 document.form2.username.focus();
 return false;
}
if(!chkname(document.form2.username.value)){
 window.alert ('用户名不规则!');
 document.form2.username.focus();
 return false;
}
if ((document.form2.pwd.value)==""){
 window.alert ('密码必须填写!');
 document.form2.pwd.focus();
 return false;
}
if ((document.form2.pwd2.value)==""){
 window.alert ('重复密码必须填写!');
 document.form2.pwd2.focus();
 return false;
}
if ((document.form2.pwd.value)!==(document.form2.pwd2.value)){
 window.alert ('两次填写的密码不一样，请重写填写!');
 document.form2.pwd2.focus();
 return false;
}
if (form2.password.value!==form2.password2.value){
 window.alert ('两次输入的密码不一致!');
 document.form2.password2.focus();
 return false;
}
if ((document.form2.email.value)==""){
 window.alert ('email必须填写!');
 document.form2.email.focus();
 return false;
}
if (!isValidEmail(document.form2.email.value)){
 window.alert ('email的格式不正确!');
 document.form2.email.focus();
 return false;
}
if ((document.form2.pagename.value)==""){
 window.alert ('项目名称必须填写!');
 document.form2.pagename.focus();
 return false;
}
if ((document.form2.pageurl.value)==""){
 window.alert ('项目网址必须填写!');
 document.form2.pageurl.focus();
 return false;
}
var check_length = document.form2.quickstyle.length;
var i_count=0
for(var i=0;i<check_length;i++){
 if (document.form2.quickstyle(i).checked){
  i_count=i_count+1;
 }
}	  
if(i_count==0){
 window.alert("请选择一种基本样式!");
 return false;
}
if (((document.form2.quickstyle(0).checked)||(document.form2.quickstyle(1).checked)||(document.form2.quickstyle(2).checked)||(document.form2.quickstyle(6).checked))&&(form2.tjopen.options[form2.tjopen.selectedIndex].value)==""){
 window.alert ('请选择是否公开统计!');
 document.form2.tjopen.focus();
 return false;
}

if (((document.form2.quickstyle(0).checked)||(document.form2.quickstyle(3).checked))&&(document.form2.showtotal.value=="")){
 window.alert ('请输入初始显示数!');
 document.form2.showtotal.focus();
 return false;
}

if ((document.form2.checkcode.value)==""){
 window.alert ('请输入四位数的验证码!');
 document.form2.checkcode.focus();
 return false;
}

if (document.form2.agreeprotocol.checked==false){
 window.alert("同意协议才能注册用户!");
 return false;
}
return true;
}

function isValidEmail(s)
{
 var reg1 = new RegExp('^[a-zA-Z0-9][a-zA-Z0-9@._-]{3,}[a-zA-Z]$');
 var reg2 = new RegExp('[@.]{2}');
 if (s.search(reg1) == -1
  || s.indexOf('@') == -1
  || s.lastIndexOf('.') < s.lastIndexOf('@')
  || s.lastIndexOf('@') != s.indexOf('@')
  || s.search(reg2) != -1)
  return false;		
 return true;
}
</script>


<SCRIPT>
function WriteCookie (cookieName, cookieValue, expiry)   
{   
var expDate = new Date();   
if(expiry)   
{   
expDate.setTime (expDate.getTime() + expiry);   
document.cookie = cookieName + "=" + escape (cookieValue) + "; path=/ expires=" + expDate.toGMTString();   
}   
else   
{   
document.cookie = cookieName + "=" + escape (cookieValue);   
}   
}   
sparetime=1000*60*60*24*3650;
WriteCookie('CFCountGGCookie',"1",sparetime);

function ashow_0(){
a_1.style.display = "none";
a_2.style.display = "none";
a_3.style.display = "none";
a_4.style.display = "none";
a_5.style.display = "none";
a_6.style.display = "none";
b_1.style.display = "none";
b_2.style.display = "none";
}
function ashow_1(){
a_1.style.display = "";
a_2.style.display = "none";
a_3.style.display = "none";
a_4.style.display = "none";
a_5.style.display = "none";
a_6.style.display = "none";
b_1.style.display = "";
b_2.style.display = "";
}
function ashow_2(){
a_1.style.display = "none";
a_2.style.display = "";
a_3.style.display = "none";
a_4.style.display = "none";
a_5.style.display = "none";
a_6.style.display = "none";
b_1.style.display = "";
b_2.style.display = "none";
}
function ashow_3(){
a_1.style.display = "none";
a_2.style.display = "none";
a_3.style.display = "";
a_4.style.display = "none";
a_5.style.display = "none";
a_6.style.display = "none";
b_1.style.display = "";
b_2.style.display = "none";
}
function ashow_4(){
a_1.style.display = "none";
a_2.style.display = "none";
a_3.style.display = "none";
a_4.style.display = "";
a_5.style.display = "none";
a_6.style.display = "none";
b_1.style.display = "none";
b_2.style.display = "";
}
function ashow_5(){
a_1.style.display = "none";
a_2.style.display = "none";
a_3.style.display = "none";
a_4.style.display = "none";
a_5.style.display = "";
a_6.style.display = "none";
b_1.style.display = "none";
b_2.style.display = "none";
}
function ashow_6(){
a_1.style.display = "none";
a_2.style.display = "none";
a_3.style.display = "none";
a_4.style.display = "none";
a_5.style.display = "none";
a_6.style.display = "";
b_1.style.display = "none";
b_2.style.display = "none";
}
function ashow_7(){
a_1.style.display = "none";
a_2.style.display = "none";
a_3.style.display = "none";
a_4.style.display = "none";
a_5.style.display = "none";
a_6.style.display = "none";
b_1.style.display = "";
b_2.style.display = "none";
}
</script>
        <table class="tb_2">

<form method="POST" action="?Action=userregsave" name="form2" onsubmit="return userregcheck()">
          <tr class="tr_1">
              <td height="20" colspan="2" >请 填 写 申 
                  请 的 帐 号 资 料（带*为必填项）</td>
            </tr>
            <tr  > 
              <td width="30%" class="td_3">用户名：</td>
              <td  ><input type="text" name="username" id="username" size="30"> 
               <font color="#0000FF">*</font><input type="button" name="Submit3" value="检查用户名是否可用" style="width:128px" onClick="usercheck();">(只能为小写英文字母,数字和下划线)<span id='usernametext'></span></td>
            </tr>
            <tr  > 
              <td height="24"   class="td_3">密　码：</td>
              <td  >
                <input name="pwd" type="password" id="password"  size="33">
                <font color="#0000FF">*</font></td>
            </tr>
            <tr  > 
              <td   class="td_3">重复密码：</td>
              <td  >
                <input name="pwd2" type="password" id="password2"  size="33">
                <font color="#0000FF">*</font></td>
            </tr>
            <tr  > 
              <td   class="td_3">密码提示问题：</td>
              <td  >
                <input name="passwordask" type="text" id="passwordask"  size="30">
                </td>
            </tr>
            <tr  > 
              <td   class="td_3">密码回答答案：</td>
              <td  >
                <input name="passwordanswer" type="text" id="passwordanswer"  size="30">
                </td>
            </tr>
            <tr  > 
              <td   class="td_3">E-Mail：</td>
              <td  >
                <input type="text" name="email" size="30" >
                <font color="#0000FF">*</font></td>
            </tr>
            <tr  > 
              <td   class="td_3">QQ号码：</td>
              <td  >
                <input name="qq" type="text" id="qq" size="30" >
                </td>
            </tr>
            <tr  > 
              <td   class="td_3">项目名称：</td>
              <td  >
                <input type="text" name="pagename" size="30" value="邮件营销效果跟踪" >
                <font color="#0000FF">*</font></td>
            </tr>
            <tr  > 
              <td  class="td_3">项目网址： 
                </td>
              <td  >
                <input name="pageurl" type="text" value="http://www.qianlongsoft.com/" size="30" >
                <font color="#0000FF">*</font>可以填写自己的网站地址</td>
            </tr>
            <tr  > 
              <td   class="td_3">请选择一种基本样式：</td>
              <td  > <input name="quickstyle" type="radio" value="1" onclick="ashow_1()">
                数字型 
                <input type="radio" name="quickstyle" value="2" onclick="ashow_2()">
                图标型 
                <input type="radio" name="quickstyle" value="3" onclick="ashow_3()">
                文字型 
                <input name="quickstyle" type="radio" value="4" onclick="ashow_4()">
                数字型+描述 
                <input type="radio" name="quickstyle" value="5" onclick="ashow_5()">
                图标型+描述 
                <input type="radio" name="quickstyle" value="6" onclick="ashow_6()">
                文字型+描述
                <input type="radio" name="quickstyle" value="7" onclick="ashow_7()">
                留言样式
				</td>
            </tr>
            <tr   id="a_1"> 
              <td   class="td_3">样式效果示范：</td>
              <td  ><a href=<%=HttpPath(2)%> target=_blank title=统计服务由[<%=RsSet("SysTitle")%>]免费提供><img src=CounterPic/1/0.gif border='0'><img src=CounterPic/1/0.gif border='0'><img src=CounterPic/1/0.gif border='0'><img src=CounterPic/1/1.gif border='0'><img src=CounterPic/1/2.gif border='0'><img src=CounterPic/1/3.gif border='0'><img src=CounterPic/1/4.gif border='0'><img src=CounterPic/1/5.gif border='0'></a>&nbsp;注：申请帐号后可以在几十种里选择一种你选择的图片</td>
            </tr>
            <tr   id="a_2"> 
              <td   class="td_3">样式效果示范：</td>
              <td  ><a href=<%=HttpPath(2)%> target=_blank title=统计服务由[<%=RsSet("SysTitle")%>]免费提供><img src=images/counter_2.gif border='0'></a></td>
            </tr>
            <tr   id="a_3"> 
              <td  class="td_3" >样式效果示范：</td>
              <td  ><a href=<%=HttpPath(2)%> target=_blank title=统计服务由[<%=RsSet("SysTitle")%>]免费提供><%=RsSet("TjTextName")%></a></td>
            </tr>
            <tr   id="a_4"> 
              <td   class="td_3">样式效果示范：</td>
              <td  ><table border='0' cellpadding='0' cellspacing='0'>
                  <tr> 
                    <td><div align='center'><a href=<%=HttpPath(2)%> target=_blank title=网站名称：康佑原创程序&#13;今天浏览：73&#13;今天IP：16&#13;在线人数：2&#13;&#13;统计服务由[<%=RsSet("SysTitle")%>]免费提供><img src=CounterPic/1/0.gif border='0'><img src=CounterPic/1/0.gif border='0'><img src=CounterPic/1/0.gif border='0'><img src=CounterPic/1/1.gif border='0'><img src=CounterPic/1/2.gif border='0'><img src=CounterPic/1/3.gif border='0'><img src=CounterPic/1/4.gif border='0'><img src=CounterPic/1/5.gif border='0'></a></td>
                    <td><div align='center'><a href=<%=HttpPath(2)%> target=_blank title='网站名称：康佑原创程序&#13;今天浏览：73&#13;今天IP：16&#13;在线人数：2&#13;&#13;统计服务由[<%=RsSet("SysTitle")%>]免费提供'><span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>同时在线[<font color='#FF0000'>2</font>]人</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>今天浏览[<font color='#FF0000'>73</font>]</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>今天IP[<font color='#FF0000'>16</font>]</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>你的IP[<font color='#FF0000'>127.0.0.1</font>]</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>欢迎你第[<font color='#FF0000'>5</font>]次访问本站</span></a></td>
                  </tr>
                </table></td>
            </tr>
            <tr   id="a_5"> 
              <td class="td_3">样式效果示范：</td>
              <td  ><table border='0' cellpadding='0' cellspacing='0'>
                  <tr> 
                    <td><div align='center'><a href=<%=HttpPath(2)%> target=_blank title=统计服务由[<%=RsSet("SysTitle")%>]免费提供><img src=images/counter_2.gif border='0'></a></td>
                    <td><div align='center'><a href=<%=HttpPath(2)%> target=_blank title='统计服务由[<%=RsSet("SysTitle")%>]免费提供'><span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>同时在线[<font color='#FF0000'>1</font>]人</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>今天浏览[<font color='#FF0000'>47</font>]</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>今天IP[<font color='#FF0000'>15</font>]</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>你的IP[<font color='#FF0000'>127.0.0.1</font>]</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>欢迎你第[<font color='#FF0000'>4</font>]次访问本站</span></a></td>
                  </tr>
                </table></td>
            </tr>
            <tr   id="a_6"> 
              <td   class="td_3">样式效果示范：</td>
              <td  ><table border='0' cellpadding='0' cellspacing='0'>
                  <tr> 
                    <td><div align='center'><a href=<%=HttpPath(2)%> target=_blank title=统计服务由[<%=RsSet("SysTitle")%>]免费提供><%=RsSet("TjTextName")%></a></td>
                    <td><div align='center'><a href=<%=HttpPath(2)%> target=_blank title='统计服务由[<%=RsSet("SysTitle")%>]免费提供'><span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>同时在线[<font color='#FF0000'>1</font>]人</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>今天浏览[<font color='#FF0000'>47</font>]</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>今天IP[<font color='#FF0000'>15</font>]</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>你的IP[<font color='#FF0000'>127.0.0.1</font>]</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>欢迎你第[<font color='#FF0000'>4</font>]次访问本站</span></a></td>
                  </tr>
                </table></td>
            </tr>

            <tr   id="b_1"> 
              <td  class="td_3" >是否公开统计：</td>
              <td  ><select name="tjopen" style="width:80px;">
                  <option  value="" selected>请选择</option>
                  <option value="-1">公开</option>
                  <option value="0">不公开</option>
                </select></td>
            </tr>
			
            <tr   id="b_2">
              <td   class="td_3">初始显示数：</td>
              <td  ><input name="showtotal" type="text" id="showtotal" value="0" size="10"></td>
            </tr>
			
            <tr  > 
              <td  class="td_3" >验证码：</td>
              <td  ><INPUT id=checkcode type=text size=10 name=checkcode> 
                <IMG height=14 id=ccimg src="CF_CheckCode.asp">&nbsp;<a href="javascript:ccimgchange()">看不清楚？换一个</a></td>
            </tr>
            <tr  > 
              <td  >&nbsp;</td>
              <td  ><input name="agreeprotocol" type="checkbox" id="agreeprotocol" value="-1" checked>
                注册此系统用户即表示你同意并承诺遵守《<a href="?Action=userrules" target="_blank">“<%=RsSet("SysTitle")%>” 用户守则</a>》的全部条款。 </td>
            </tr>
            <tr  > 
              <td  >&nbsp;</td>
              <td  ><input type="submit" name="Submit" value="申请">
                　 
                <input type="reset" name="Submit2" value="重填"></td>
            </tr>
          </form>
        </table>
		<%Else%>
<table class="tb_2">
<tr class="tr_1">
<td>对不起,本统计系统已经暂停新用户注册!</td>
</tr>
</table>

 <%End If%>
<%End if%>

<%If Request("Action")="userrules" Then%>
        <table class="tb_2">
          <tr class="tr_1">
              
    <TD height=19>&nbsp;&nbsp;<%=RsSet("SysTitle")%>用户守则</TD>
            </TR>
            <TR> 
              <TD></TD>
            </TR>
            <TR> 
              <TD class="td_1"> 第一条　本守则所称“<%=RsSet("SysTitle")%>用户”(以下简称“用户”)是指所有使用“<%=RsSet("SysTitle")%>及相关服务”(以下简称“<%=RsSet("SysTitle")%> 的服务”) 的单位或个人。 
                <p>第二条　用户享有 <%=RsSet("SysTitle")%> 的服务。除非用户自行公开，<%=RsSet("SysTitle")%> 将保护用户个人信息，并且不会主动将这些隐私提供给第三方使用。用户会对由非主观因素导致的资料泄漏事件给予充分的谅解且不予追究 
                  <%=RsSet("SysTitle")%> 的法律责任。 </p>
                <p>第三条　用户不得在以下网站上享用 <%=RsSet("SysTitle")%> 的服务。 <br>
                  1、存在反动、淫秽色情、政治敏感问题等内容的网站； <br>
                  2、存在宣扬暴力、邪教理论、种族歧视思想的内容的网站； <br>
                  3、存在贩卖、传播非法制品及管制产品的行为或内容的网站； <br>
                  4、存在其它与国家法律法规、地方法规、国际公约、公共道德相违背的内容的网站。 </p>
                <p>第四条　用户不得通过任何手段隐藏统计代码输出的图标或文字，这些手段包括但不限于将统计代码放在隐藏的框架或层中。 </p>
                <p>第五条　用户可随时停止使用 <%=RsSet("SysTitle")%> 的服务，并可随时删除其在 <%=RsSet("SysTitle")%> 的所有信息。用户允许 <%=RsSet("SysTitle")%> 随时中断或终止为其提供的任何服务。 </p>
                <p>第六条　用户不得通过已知或者未知的任何作弊手段影响统计数据和排名数据的准确性。作弊手段包括但不限于使用漫游器点击网站、恶意使用代理服务器点击网站、在同一个页面上放置多个统计代码、将统计代码放在自动刷新的页面或框架中等。 
                </p>
                <p>第七条　用户不得买卖 <%=RsSet("SysTitle")%> 的服务和用户名，不得将 <%=RsSet("SysTitle")%> 的服务或用户名用于交换其他可被交换的产品或信息。 </p>
                <p>第八条　用户承认并承诺其所进行的一切活动与 <%=RsSet("SysTitle")%> 及 <%=RsSet("SysTitle")%> 的服务无关，用户承认并承诺其进行的活动所造成的损失及赔偿责任由用户承担。 </p>
                <p>第九条　用户不得将 <%=RsSet("SysTitle")%> 的统计数据及相关内容作为诈骗等非法活动的工具。 </p>
                <p>第十条　用户不得以 <%=RsSet("SysTitle")%> 的名义从事任何活动。 </p>
                <p>第十一条　<%=RsSet("SysTitle")%> 不会强迫任何单位或个人成为用户，但必须接受并遵守本守则才能成为用户。<br>
                  <br>
                </p></TD>
            </TR>
          </TBODY>
        </TABLE>
<%End If%>	

<%
If Request("Action")="" And RsSet("RegState")=-1 Then
Show="ashow_0();"

Response.write "<script>"               &Chr(13)&Chr(10)
Response.write "function myshow()"      &Chr(13)&Chr(10)
Response.write "{"					    &Chr(13)&Chr(10)
Response.write  Show                    &Chr(13)&Chr(10)
Response.write "clearInterval(myst);"   &Chr(13)&Chr(10)
Response.write "}"                      &Chr(13)&Chr(10)
Response.write "myst=setInterval('myshow()',1);"  &Chr(13)&Chr(10)
Response.write "</script>"              &Chr(13)&Chr(10)
End If
%>