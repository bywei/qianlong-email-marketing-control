
<%If Action="" Then%>
<%If RsSet("RegState")=-1 Then%>
<script>
function usercheck() {
var username = document.getElementById("username").value;
if(username == "")
{
 usernametext.innerHTML="<br>�û�������Ϊ��";
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
var Letters = "abcdefghijklmnopqrstuvwxyz1234567890_"; //�����Լ����ӿ�����ֵ
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
 window.alert ('�û���������д!');
 document.form2.username.focus();
 return false;
}
if(!chkname(document.form2.username.value)){
 window.alert ('�û���������!');
 document.form2.username.focus();
 return false;
}
if ((document.form2.pwd.value)==""){
 window.alert ('���������д!');
 document.form2.pwd.focus();
 return false;
}
if ((document.form2.pwd2.value)==""){
 window.alert ('�ظ����������д!');
 document.form2.pwd2.focus();
 return false;
}
if ((document.form2.pwd.value)!==(document.form2.pwd2.value)){
 window.alert ('������д�����벻һ��������д��д!');
 document.form2.pwd2.focus();
 return false;
}
if (form2.password.value!==form2.password2.value){
 window.alert ('������������벻һ��!');
 document.form2.password2.focus();
 return false;
}
if ((document.form2.email.value)==""){
 window.alert ('email������д!');
 document.form2.email.focus();
 return false;
}
if (!isValidEmail(document.form2.email.value)){
 window.alert ('email�ĸ�ʽ����ȷ!');
 document.form2.email.focus();
 return false;
}
if ((document.form2.pagename.value)==""){
 window.alert ('��Ŀ���Ʊ�����д!');
 document.form2.pagename.focus();
 return false;
}
if ((document.form2.pageurl.value)==""){
 window.alert ('��Ŀ��ַ������д!');
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
 window.alert("��ѡ��һ�ֻ�����ʽ!");
 return false;
}
if (((document.form2.quickstyle(0).checked)||(document.form2.quickstyle(1).checked)||(document.form2.quickstyle(2).checked)||(document.form2.quickstyle(6).checked))&&(form2.tjopen.options[form2.tjopen.selectedIndex].value)==""){
 window.alert ('��ѡ���Ƿ񹫿�ͳ��!');
 document.form2.tjopen.focus();
 return false;
}

if (((document.form2.quickstyle(0).checked)||(document.form2.quickstyle(3).checked))&&(document.form2.showtotal.value=="")){
 window.alert ('�������ʼ��ʾ��!');
 document.form2.showtotal.focus();
 return false;
}

if ((document.form2.checkcode.value)==""){
 window.alert ('��������λ������֤��!');
 document.form2.checkcode.focus();
 return false;
}

if (document.form2.agreeprotocol.checked==false){
 window.alert("ͬ��Э�����ע���û�!");
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
              <td height="20" colspan="2" >�� �� д �� 
                  �� �� �� �� �� �ϣ���*Ϊ�����</td>
            </tr>
            <tr  > 
              <td width="30%" class="td_3">�û�����</td>
              <td  ><input type="text" name="username" id="username" size="30"> 
               <font color="#0000FF">*</font><input type="button" name="Submit3" value="����û����Ƿ����" style="width:128px" onClick="usercheck();">(ֻ��ΪСдӢ����ĸ,���ֺ��»���)<span id='usernametext'></span></td>
            </tr>
            <tr  > 
              <td height="24"   class="td_3">�ܡ��룺</td>
              <td  >
                <input name="pwd" type="password" id="password"  size="33">
                <font color="#0000FF">*</font></td>
            </tr>
            <tr  > 
              <td   class="td_3">�ظ����룺</td>
              <td  >
                <input name="pwd2" type="password" id="password2"  size="33">
                <font color="#0000FF">*</font></td>
            </tr>
            <tr  > 
              <td   class="td_3">������ʾ���⣺</td>
              <td  >
                <input name="passwordask" type="text" id="passwordask"  size="30">
                </td>
            </tr>
            <tr  > 
              <td   class="td_3">����ش�𰸣�</td>
              <td  >
                <input name="passwordanswer" type="text" id="passwordanswer"  size="30">
                </td>
            </tr>
            <tr  > 
              <td   class="td_3">E-Mail��</td>
              <td  >
                <input type="text" name="email" size="30" >
                <font color="#0000FF">*</font></td>
            </tr>
            <tr  > 
              <td   class="td_3">QQ���룺</td>
              <td  >
                <input name="qq" type="text" id="qq" size="30" >
                </td>
            </tr>
            <tr  > 
              <td   class="td_3">��Ŀ���ƣ�</td>
              <td  >
                <input type="text" name="pagename" size="30" value="�ʼ�Ӫ��Ч������" >
                <font color="#0000FF">*</font></td>
            </tr>
            <tr  > 
              <td  class="td_3">��Ŀ��ַ�� 
                </td>
              <td  >
                <input name="pageurl" type="text" value="http://www.qianlongsoft.com/" size="30" >
                <font color="#0000FF">*</font>������д�Լ�����վ��ַ</td>
            </tr>
            <tr  > 
              <td   class="td_3">��ѡ��һ�ֻ�����ʽ��</td>
              <td  > <input name="quickstyle" type="radio" value="1" onclick="ashow_1()">
                ������ 
                <input type="radio" name="quickstyle" value="2" onclick="ashow_2()">
                ͼ���� 
                <input type="radio" name="quickstyle" value="3" onclick="ashow_3()">
                ������ 
                <input name="quickstyle" type="radio" value="4" onclick="ashow_4()">
                ������+���� 
                <input type="radio" name="quickstyle" value="5" onclick="ashow_5()">
                ͼ����+���� 
                <input type="radio" name="quickstyle" value="6" onclick="ashow_6()">
                ������+����
                <input type="radio" name="quickstyle" value="7" onclick="ashow_7()">
                ������ʽ
				</td>
            </tr>
            <tr   id="a_1"> 
              <td   class="td_3">��ʽЧ��ʾ����</td>
              <td  ><a href=<%=HttpPath(2)%> target=_blank title=ͳ�Ʒ�����[<%=RsSet("SysTitle")%>]����ṩ><img src=CounterPic/1/0.gif border='0'><img src=CounterPic/1/0.gif border='0'><img src=CounterPic/1/0.gif border='0'><img src=CounterPic/1/1.gif border='0'><img src=CounterPic/1/2.gif border='0'><img src=CounterPic/1/3.gif border='0'><img src=CounterPic/1/4.gif border='0'><img src=CounterPic/1/5.gif border='0'></a>&nbsp;ע�������ʺź�����ڼ�ʮ����ѡ��һ����ѡ���ͼƬ</td>
            </tr>
            <tr   id="a_2"> 
              <td   class="td_3">��ʽЧ��ʾ����</td>
              <td  ><a href=<%=HttpPath(2)%> target=_blank title=ͳ�Ʒ�����[<%=RsSet("SysTitle")%>]����ṩ><img src=images/counter_2.gif border='0'></a></td>
            </tr>
            <tr   id="a_3"> 
              <td  class="td_3" >��ʽЧ��ʾ����</td>
              <td  ><a href=<%=HttpPath(2)%> target=_blank title=ͳ�Ʒ�����[<%=RsSet("SysTitle")%>]����ṩ><%=RsSet("TjTextName")%></a></td>
            </tr>
            <tr   id="a_4"> 
              <td   class="td_3">��ʽЧ��ʾ����</td>
              <td  ><table border='0' cellpadding='0' cellspacing='0'>
                  <tr> 
                    <td><div align='center'><a href=<%=HttpPath(2)%> target=_blank title=��վ���ƣ�����ԭ������&#13;���������73&#13;����IP��16&#13;����������2&#13;&#13;ͳ�Ʒ�����[<%=RsSet("SysTitle")%>]����ṩ><img src=CounterPic/1/0.gif border='0'><img src=CounterPic/1/0.gif border='0'><img src=CounterPic/1/0.gif border='0'><img src=CounterPic/1/1.gif border='0'><img src=CounterPic/1/2.gif border='0'><img src=CounterPic/1/3.gif border='0'><img src=CounterPic/1/4.gif border='0'><img src=CounterPic/1/5.gif border='0'></a></td>
                    <td><div align='center'><a href=<%=HttpPath(2)%> target=_blank title='��վ���ƣ�����ԭ������&#13;���������73&#13;����IP��16&#13;����������2&#13;&#13;ͳ�Ʒ�����[<%=RsSet("SysTitle")%>]����ṩ'><span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>ͬʱ����[<font color='#FF0000'>2</font>]��</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>�������[<font color='#FF0000'>73</font>]</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>����IP[<font color='#FF0000'>16</font>]</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>���IP[<font color='#FF0000'>127.0.0.1</font>]</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>��ӭ���[<font color='#FF0000'>5</font>]�η��ʱ�վ</span></a></td>
                  </tr>
                </table></td>
            </tr>
            <tr   id="a_5"> 
              <td class="td_3">��ʽЧ��ʾ����</td>
              <td  ><table border='0' cellpadding='0' cellspacing='0'>
                  <tr> 
                    <td><div align='center'><a href=<%=HttpPath(2)%> target=_blank title=ͳ�Ʒ�����[<%=RsSet("SysTitle")%>]����ṩ><img src=images/counter_2.gif border='0'></a></td>
                    <td><div align='center'><a href=<%=HttpPath(2)%> target=_blank title='ͳ�Ʒ�����[<%=RsSet("SysTitle")%>]����ṩ'><span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>ͬʱ����[<font color='#FF0000'>1</font>]��</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>�������[<font color='#FF0000'>47</font>]</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>����IP[<font color='#FF0000'>15</font>]</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>���IP[<font color='#FF0000'>127.0.0.1</font>]</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>��ӭ���[<font color='#FF0000'>4</font>]�η��ʱ�վ</span></a></td>
                  </tr>
                </table></td>
            </tr>
            <tr   id="a_6"> 
              <td   class="td_3">��ʽЧ��ʾ����</td>
              <td  ><table border='0' cellpadding='0' cellspacing='0'>
                  <tr> 
                    <td><div align='center'><a href=<%=HttpPath(2)%> target=_blank title=ͳ�Ʒ�����[<%=RsSet("SysTitle")%>]����ṩ><%=RsSet("TjTextName")%></a></td>
                    <td><div align='center'><a href=<%=HttpPath(2)%> target=_blank title='ͳ�Ʒ�����[<%=RsSet("SysTitle")%>]����ṩ'><span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>ͬʱ����[<font color='#FF0000'>1</font>]��</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>�������[<font color='#FF0000'>47</font>]</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>����IP[<font color='#FF0000'>15</font>]</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>���IP[<font color='#FF0000'>127.0.0.1</font>]</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>��ӭ���[<font color='#FF0000'>4</font>]�η��ʱ�վ</span></a></td>
                  </tr>
                </table></td>
            </tr>

            <tr   id="b_1"> 
              <td  class="td_3" >�Ƿ񹫿�ͳ�ƣ�</td>
              <td  ><select name="tjopen" style="width:80px;">
                  <option  value="" selected>��ѡ��</option>
                  <option value="-1">����</option>
                  <option value="0">������</option>
                </select></td>
            </tr>
			
            <tr   id="b_2">
              <td   class="td_3">��ʼ��ʾ����</td>
              <td  ><input name="showtotal" type="text" id="showtotal" value="0" size="10"></td>
            </tr>
			
            <tr  > 
              <td  class="td_3" >��֤�룺</td>
              <td  ><INPUT id=checkcode type=text size=10 name=checkcode> 
                <IMG height=14 id=ccimg src="CF_CheckCode.asp">&nbsp;<a href="javascript:ccimgchange()">�����������һ��</a></td>
            </tr>
            <tr  > 
              <td  >&nbsp;</td>
              <td  ><input name="agreeprotocol" type="checkbox" id="agreeprotocol" value="-1" checked>
                ע���ϵͳ�û�����ʾ��ͬ�Ⲣ��ŵ���ء�<a href="?Action=userrules" target="_blank">��<%=RsSet("SysTitle")%>�� �û�����</a>����ȫ����� </td>
            </tr>
            <tr  > 
              <td  >&nbsp;</td>
              <td  ><input type="submit" name="Submit" value="����">
                �� 
                <input type="reset" name="Submit2" value="����"></td>
            </tr>
          </form>
        </table>
		<%Else%>
<table class="tb_2">
<tr class="tr_1">
<td>�Բ���,��ͳ��ϵͳ�Ѿ���ͣ���û�ע��!</td>
</tr>
</table>

 <%End If%>
<%End if%>

<%If Request("Action")="userrules" Then%>
        <table class="tb_2">
          <tr class="tr_1">
              
    <TD height=19>&nbsp;&nbsp;<%=RsSet("SysTitle")%>�û�����</TD>
            </TR>
            <TR> 
              <TD></TD>
            </TR>
            <TR> 
              <TD class="td_1"> ��һ�������������ơ�<%=RsSet("SysTitle")%>�û���(���¼�ơ��û���)��ָ����ʹ�á�<%=RsSet("SysTitle")%>����ط���(���¼�ơ�<%=RsSet("SysTitle")%> �ķ���) �ĵ�λ����ˡ� 
                <p>�ڶ������û����� <%=RsSet("SysTitle")%> �ķ��񡣳����û����й�����<%=RsSet("SysTitle")%> �������û�������Ϣ�����Ҳ�����������Щ��˽�ṩ��������ʹ�á��û�����ɷ��������ص��µ�����й©�¼������ֵ��½��Ҳ���׷�� 
                  <%=RsSet("SysTitle")%> �ķ������Ρ� </p>
                <p>���������û�������������վ������ <%=RsSet("SysTitle")%> �ķ��� <br>
                  1�����ڷ���������ɫ�顢����������������ݵ���վ�� <br>
                  2���������ﱩ����а�����ۡ���������˼������ݵ���վ�� <br>
                  3�����ڷ����������Ƿ���Ʒ�����Ʋ�Ʒ����Ϊ�����ݵ���վ�� <br>
                  4��������������ҷ��ɷ��桢�ط����桢���ʹ�Լ������������Υ�������ݵ���վ�� </p>
                <p>���������û�����ͨ���κ��ֶ�����ͳ�ƴ��������ͼ������֣���Щ�ֶΰ����������ڽ�ͳ�ƴ���������صĿ�ܻ���С� </p>
                <p>���������û�����ʱֹͣʹ�� <%=RsSet("SysTitle")%> �ķ��񣬲�����ʱɾ������ <%=RsSet("SysTitle")%> ��������Ϣ���û����� <%=RsSet("SysTitle")%> ��ʱ�жϻ���ֹΪ���ṩ���κη��� </p>
                <p>���������û�����ͨ����֪����δ֪���κ������ֶ�Ӱ��ͳ�����ݺ��������ݵ�׼ȷ�ԡ������ֶΰ�����������ʹ�������������վ������ʹ�ô�������������վ����ͬһ��ҳ���Ϸ��ö��ͳ�ƴ��롢��ͳ�ƴ�������Զ�ˢ�µ�ҳ������еȡ� 
                </p>
                <p>���������û��������� <%=RsSet("SysTitle")%> �ķ�����û��������ý� <%=RsSet("SysTitle")%> �ķ�����û������ڽ��������ɱ������Ĳ�Ʒ����Ϣ�� </p>
                <p>�ڰ������û����ϲ���ŵ�������е�һ�л�� <%=RsSet("SysTitle")%> �� <%=RsSet("SysTitle")%> �ķ����޹أ��û����ϲ���ŵ����еĻ����ɵ���ʧ���⳥�������û��е��� </p>
                <p>�ھ������û����ý� <%=RsSet("SysTitle")%> ��ͳ�����ݼ����������Ϊթƭ�ȷǷ���Ĺ��ߡ� </p>
                <p>��ʮ�����û������� <%=RsSet("SysTitle")%> ����������κλ�� </p>
                <p>��ʮһ����<%=RsSet("SysTitle")%> ����ǿ���κε�λ����˳�Ϊ�û�����������ܲ����ر�������ܳ�Ϊ�û���<br>
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