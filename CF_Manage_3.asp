
<%If Action="jsqset" Then%>
<%
 If Session("CFCountUser_View")<>"" Then'Ϊ�����鿴��ʱ
  Call AlertBack("�����鿴���޷��Դ˲���",1)
 End If
%>
<SCRIPT>
function ashow_1(){
a_1.style.display = "";
a_2.style.display = "";
a_3.style.display = "none";
}
function ashow_2(){
a_1.style.display = "none";
a_2.style.display = "none";
a_3.style.display = "";
}
function bshow_1(){
b_1.style.display = "";
}
function bshow_2(){
b_1.style.display = "none";
}
function cshow_1(){
c_1.style.display = "";
}
function cshow_2(){
c_1.style.display = "none";
}
function dshow_1(){
d_1.style.display = "";
}
function dshow_2(){
d_1.style.display = "none";
}
</SCRIPT>

  <div class="filters">
			 <h3>��������</h3>
       </div>
        <table width="100%" >
         
  <form name="form3" method="post" action="?Action=jsqsetsave">
 <tr >
      <td colspan="2"> �������²���(��*Ϊ�����</td>
    </tr>
    <tr > 
      <td width="24%">�Զ�����ʾ����</td>
      <td > <input name="ShowTotal" type="text" id="ShowTotal" value="<%=rs("ShowTotal")%>" size="30" > 
        <font color="#0000FF">*</font>(�Զ���)</td>
    </tr>
    <tr > 
      <td >ʵ����ʾ����</td>
      <td ><%=rs("RealShowTotal")%>&nbsp;&nbsp;(ҳ����������)</td>
    </tr>
    <tr > 
      <td >ʵ����ֵ��</td>
      <td ><%=rs("RealIpTotal")%>&nbsp;&nbsp;(IP��)</td>
    </tr>
    <tr > 
      <td >ͳ����ʾλ����</td>
      <td ><input name="PicNum" type="text" id="total2" value="<%=rs("PicNum")%>" size="30" > 
        <font color="#0000FF">*</font></td>
    </tr>
    <tr > 
      <td >ͳ�ƿ�ʼ���ڣ�</td>
      <td > <input name="adddate" type="text" id="adddate" value="<%=rs("adddate")%>" size="30" > 
        <font color="#0000FF">*</font></td>
    </tr>
    <tr > 
      <td >Scriptͳ����ʾ���֣�</td>
      <td > <input type="radio" name="CounterShow" value="-1"<%if rs("CounterShow")=-1 then response.write " checked"%> onclick="ashow_1()">
        �� 
        <input type="radio" name="CounterShow" value="0"<%if rs("CounterShow")=0 then response.write " checked"%> onclick="ashow_2()">
        ��</td>
    </tr>
    <tr  id="a_1"> 
      <td >Scriptͳ����ʾ���֣�</td>
      <td > <input name="ShowType" type="radio" id="ShowType" value="1" <%if rs("ShowType")=1 then response.write "checked"%>>
        �Զ�����ʾ�� 
        <input name="ShowType" type="radio" id="ShowType" value="2" <%if rs("ShowType")=2 then response.write "checked"%>>
        ʵ����ʾ��ֵ 
        <input name="ShowType" type="radio" id="ShowType" value="3" <%if rs("ShowType")=3 then response.write "checked"%>>
        ʵ��Ip��ֵ </td>
    </tr>
    <tr  id="a_2"> 
      <td >�����õļ�����ʽ��</td>
      <td ><%=GetPicStyle(UserName)%>&nbsp;[<a href="?Action=stylelist">����ͼƬ��ʽ</a>] </td>
    </tr>
    <tr  id="a_3"> 
      <td >Scriptͳ������ʱĬ����ʾ��</td>
      <td ><input name="CounterHiddenPic" type="radio" id="radio" value="0" <%if rs("CounterHiddenPic")=0 then response.write "checked"%>> 
        ���� 
		<input name="CounterHiddenPic" type="radio" id="radio" value="1" <%if rs("CounterHiddenPic")=1 then response.write "checked"%>> 
        <%=RsSet("TjTextName")%> <input name="CounterHiddenPic" type="radio" id="radio" value="2" <%if rs("CounterHiddenPic")=2 then response.write "checked"%>>
        ͼƬ1 <img src="images/counter_2.gif" border="0"> <input name="CounterHiddenPic" type="radio" id="radio" value="3" <%if rs("CounterHiddenPic")=3 then response.write "checked"%>>
        ͼƬ2 <img src="images/counter_3.gif" border="0"> </td>
    </tr>
    <tr > 
      <td >Scriptͳ����ʾλ�ã�</td>
      <td ><input name="CounterSite" type="radio" id="radio" value="1" <%if rs("CounterSite")=1 then response.write "checked"%>>
        �������Ķ����� 
        <input name="CounterSite" type="radio" id="radio" value="2" <%if rs("CounterSite")=2 then response.write "checked"%>>
        �������Ķ����� 
        <input name="CounterSite" type="radio" id="radio" value="3" <%if rs("CounterSite")=3 then response.write "checked"%>>
        �������Ķ����� 
        <input name="CounterSite" type="radio" id="radio2" value="4" <%if rs("CounterSite")=4 then response.write "checked"%>>
        �������Ķ����� </td>
    </tr>
    <tr > 
      <td >Imgͳ����ʾ���֣�</td>
      <td > <input type="radio" name="ImgCounterShow" value="-1"<%if rs("ImgCounterShow")=-1 then response.write " checked"%> onclick="bshow_1()">
        �� 
        <input type="radio" name="ImgCounterShow" value="0"<%if rs("ImgCounterShow")=0 then response.write " checked"%> onclick="bshow_2()">
        ��</td>
    </tr>
    <tr  id="b_1"> 
      <td >Imgͳ����ʾ���֣�</td>
      <td ><input name="ImgShowType" type="radio" id="radio4" value="1" <%if rs("ImgShowType")=1 then response.write "checked"%>>
        �Զ�����ʾ�� 
        <input name="ImgShowType" type="radio" id="radio4" value="2" <%if rs("ImgShowType")=2 then response.write "checked"%>>
        ʵ����ʾ��ֵ 
        <input name="ImgShowType" type="radio" id="radio4" value="3" <%if rs("ImgShowType")=3 then response.write "checked"%>>
        ʵ��Ip��ֵ</td>
    </tr>
    <tr > 
      <td >��¼�����Ķ���</td>
      <td > <input type="radio" name="online" value="-1"<%if rs("online")=-1 then response.write " checked"%> onclick="cshow_1()">
        �� 
        <input type="radio" name="online" value="0"<%if rs("online")=0 then response.write " checked"%> onclick="cshow_2()">
        ��</td>
    </tr>
    <tr  id="c_1"> 
      <td >��¼�����Ķ��ļ����</td>
      <td ><input name="onlinetime" type="text" id="onlinetime" value="<%=rs("onlinetime")%>" size="30">
        ����(������180��������,һ����30����)</td>
    </tr>
    <tr > 
      <td >ͳ������ʾ�����Ķ���</td>
      <td > <input name="onlineshow" type="checkbox" id="onlineshow" value="-1" <%if rs("onlineshow")=-1 then response.write "checked"%>>
        ��</td>
    </tr>
    <tr > 
      <td >ͳ������ʾ�������:</td>
      <td ><input name="TodayShow" type="checkbox" id="TodayShow" value="-1" <%if rs("TodayShow")=-1 then response.write "checked"%>>
        ��</td>
    </tr>
    <tr > 
      <td >ͳ������ʾ����IP:</td>
      <td ><input name="TodayIpShow" type="checkbox" id="TodayIpShow" value="-1" <%if rs("TodayIpShow")=-1 then response.write "checked"%>>
        ��</td>
    </tr>
    <tr > 
      <td >ͳ������ʾIP��</td>
      <td ><input name="IpShow" type="checkbox" id="IpShow" value="-1" <%if rs("IpShow")=-1 then response.write "checked"%>>
        ��</td>
    </tr>
    <tr > 
      <td >ͳ������ʾ���ô�����</td>
      <td ><input name="VisitShow" type="checkbox" id="VisitShow" value="-1" <%if rs("VisitShow")=-1 then response.write "checked"%>>
        ��</td>
    </tr>
 <tr >
      <td colspan="2" >��������Ϣ�Ƿ񹫿�����(�����������˿ɲ鿴)��</td>
    </tr>
    <tr > 
      <td >ͳ����Ϣ������</td>
      <td ><input name="TjOpen" type="radio" id="TjOpen" value="-1" <%if rs("TjOpen")=-1 then response.write "checked"%> onclick="dshow_1()">
        �ǡ� 
        <input name="TjOpen" type="radio" id="radio3" value="0" <%if rs("TjOpen")=0 then response.write "checked"%> onclick="dshow_2()">
        ��</td>
    </tr>
    <tr  id="d_1"> 
      <td >����ѡ�е�ͳ����Ϣ��</td>
      <td ><input name="SurveyOpen" type="checkbox" id="SurveyOpen" value="-1" <%if rs("SurveyOpen")=-1 then response.write "checked"%>>
        ͳ�Ƹſ� 
        <input name="TodayLyOpen" type="checkbox" id="TodayLyOpen" value="-1" <%if rs("TodayLyOpen")=-1 then response.write "checked"%>>
        �������Դ��ϸ 
        <input name="OnlineOpen" type="checkbox" id="OnlineOpen" value="-1" <%if rs("OnlineOpen")=-1 then response.write "checked"%>>
        �����Ķ� 
        <input name="TodayHourOpen" type="checkbox" id="TodayHourOpen" value="-1" <%if rs("TodayHourOpen")=-1 then response.write "checked"%>>
        �����Сʱͳ�� 
        <input name="EveryDayOpen" type="checkbox" id="EveryDayOpen" value="-1" <%if rs("EveryDayOpen")=-1 then response.write "checked"%>>
        ÿ���ͳ��</td>
    </tr>
 <tr >
      <td colspan="2" >���԰�</td>
    </tr>
	<tr > 
      <td >��תҳ����ʾ���԰壺</td>
      <td ><input name="GbookState" type="radio" id="GbookState" value="-1" <%if rs("GbookState")=-1 then response.write "checked"%>>
        �ǡ� 
        <input name="GbookState" type="radio" id="GbookState" value="0" <%if rs("GbookState")=0 then response.write "checked"%>>
        ��</td>
    </tr>

	<tr > 
      <td >&nbsp;</td>
      <td > <input type="submit" name="Submit2" value="�ޡ���" >
        �� 
        <input type="reset" name="Submit22" value="ȡ����" > </td>
    </tr>
  </form>
 <tr >
    <td colspan="2">���ͳ��Ч�� 
    </td>
  </tr>
  <tr > 
    <td colspan="2"  >Script���ã� 
    </td>
  </tr>
  <tr > 
    <td colspan="2"  ><script src="<%=Tmp%>cf.asp?UserName=<%=Rs("UserName")%>"></script></td>
  </tr>
</table>
<%End If%>
  <%If Action="jsqquickset" Then%>
<SCRIPT>
function ashow_1(){
a_1.style.display = "";
}
function ashow_2(){
a_1.style.display = "";
}
function ashow_3(){
a_1.style.display = "";
}
function ashow_4(){
a_1.style.display = "none";
}
function ashow_5(){
a_1.style.display = "none";
}
function ashow_6(){
a_1.style.display = "none";
}
function ashow_7(){
a_1.style.display = "";
}



function jsqquickset(){
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
return true;
}
</script>
  <div class="filters">
			 <h3>ͳ�ƻ�����ʽ��������</h3>
       </div>
        <table width="100%" >


          <form name="form2" method="post" action="?Action=jsqquicksetsave"  onsubmit="return jsqquickset()">
            <tr > 
              <td width="120"  >�����ͣ�</td>
              <td  ><input type="radio" name="quickstyle" value="1" onclick="ashow_1()"></td>
              <td  ><%=GetPicStyle(UserName)%></td>
            </tr>
            <tr > 
              <td  >ͼ���ͣ�</td>
              <td  ><input type="radio" name="quickstyle" value="2" onclick="ashow_2()"></td>
              <td  > <a href=<%=Tmp%> target=_blank title=����������[<%=RsSet("SysTitle")%>]����ṩ><img src=<%=Tmp%>images/counter_2.gif border='0'></a></td>
            </tr>
            <tr > 
              <td  >�����ͣ�</td>
              <td  ><input type="radio" name="quickstyle" value="3" onclick="ashow_3()"></td>
              <td  ><a href=<%=HttpPath(2)%> target=_blank title=ͳ�Ʒ�����[<%=RsSet("SysTitle")%>]����ṩ><%=RsSet("TjTextName")%></a></td>
            </tr>
            <tr > 
              <td  >������+��ϸ������</td>
              <td  ><input type="radio" name="quickstyle" value="4"  onclick="ashow_4()"></td>
              <td  > <table border='0' cellpadding='2' cellspacing='1'>
                  <tr> 
                    <td><div align='center'><%=GetPicStyle(UserName)%></td>
                    <td><div align='center'><a href=<%=Tmp%>View.asp?UserName=<%=UserName%> target=_blank title='��վ���ƣ�����ԭ������&#13;���������73&#13;����IP��16&#13;����������2&#13;&#13;����������[<%=RsSet("SysTitle")%>]����ṩ'><span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>ͬʱ�Ķ�[<font color='#FF0000'>2</font>]��</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>�������[<font color='#FF0000'>73</font>]</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>���������[<font color='#FF0000'>16</font>]</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>���IP[<font color='#FF0000'>127.0.0.1</font>]</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>��ӭ���[<font color='#FF0000'>5</font>]���Ķ��ʼ�</span></a></td>
                  </tr>
                </table></td>
            </tr>
            <tr > 
              <td  >ͼ����+��ϸ������</td>
              <td  ><input type="radio" name="quickstyle" value="5" onclick="ashow_5()"></td>
              <td  > <table border='0' cellpadding='2' cellspacing='1'>
                  <tr> 
                    <td><div align='center'><a href=<%=Tmp%>View.asp?UserName=<%=UserName%> target=_blank title=����������[<%=RsSet("SysTitle")%>]����ṩ><img src=<%=Tmp%>images/counter_2.gif border='0'></a></td>
                    <td><div align='center'><a href=<%=Tmp%>View.asp?UserName=<%=UserName%> target=_blank title='����������[<%=RsSet("SysTitle")%>]����ṩ'><span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>ͬʱ�Ķ�[<font color='#FF0000'>1</font>]��</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>�������[<font color='#FF0000'>47</font>]</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>���������[<font color='#FF0000'>15</font>]</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>���IP[<font color='#FF0000'>127.0.0.1</font>]</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>��ӭ���[<font color='#FF0000'>4</font>]���Ķ��ʼ�</span></a></td>
                  </tr>
                </table></td>
            </tr>
            <tr > 
              <td  >������+��ϸ������</td>
              <td  ><input type="radio" name="quickstyle" value="6" onclick="ashow_6()"></td>
              <td   ><table border='0' cellpadding='0' cellspacing='0'>
                  <tr> 
                    <td><div align='center'><a href=<%=HttpPath(2)%> target=_blank title=ͳ�Ʒ�����[<%=RsSet("SysTitle")%>]����ṩ><%=RsSet("TjTextName")%></a></td>
                    <td><div align='center'><a href=<%=HttpPath(2)%> target=_blank title='ͳ�Ʒ�����[<%=RsSet("SysTitle")%>]����ṩ'><span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>ͬʱ�Ķ�[<font color='#FF0000'>1</font>]��</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>�������[<font color='#FF0000'>47</font>]</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>����IP[<font color='#FF0000'>15</font>]</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>��ķ�����[<font color='#FF0000'>127.0.0.1</font>]</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>��ӭ���[<font color='#FF0000'>4</font>]���Ķ��ʼ�</span></a></td>
                  </tr>
                </table></td>
            </tr>
            <tr > 
              <td  >�����ͣ�</td>
              <td  ><input type="radio" name="quickstyle" value="7" onclick="ashow_7()"></td>
              <td  >���԰�</td>
            </tr>
			<tr  id="a_1">
              <td colspan="2"  >�Ƿ񹫿�ͳ�ƣ�</td>
              <td   ><select name="tjopen">
                  <option value="" selected>��ѡ��</option>
                  <option value="-1">����</option>
                  <option value="0">������</option>
                </select></td>
            </tr>
            <tr > 
              <td colspan="2"  >&nbsp;</td>
              <td   > 
                <input type="submit" name="Submit2" value="�衡��">
                <br>
                (ע������ֻ�ṩ���ֻ�����ʽѡ��,�������������룢<a href="?Action=jsqset">��������</a>&quot;��&quot;<a href="?Action=stylelist">ͼƬ��ʽѡ��</a>&quot;���������ͳ��ʵ��Ч����&quot;<a href="?Action=getcode">Ԥ������ô���</a>&quot;����Կ���)</td>
            </tr>
          </form>
        </table>
<%End If%>      

<%if Action="imglist" then%>
<%
If Request("Assort")="" Then
 Assort=2
Else
 Assort=CInt(Request("Assort"))
End if



Set Rs= Server.CreateObject("Adodb.RecordSet")
Sql="Select * From CFCount_User Where UserName='"&UserName&"'"
Rs.Open Sql,Conn,1,1
ImgFileName=Rs("ImgFileName")
%>
<div class="filters">
			 <h3>�������Ч������</h3>
       </div>
<table width="100%" >

  <form name="form5" method="post" action="?Action=imgmodifysave&Assort=<%=Assort%>">

    <tr> 
      <td colspan="2"><img src="<%=Tmp%>cf.asp?c=<%=UserName%>" border="0">
      </td>
    </tr>
<tr >
      <td colspan="2">&nbsp;</td>
    </tr>
    <tr> 
      <td colspan="2">
<a href="?Action=<%=Action%>&Assort=1"><b>ʹ���Զ���ͼƬ</b></a>&nbsp;&nbsp;&nbsp;&nbsp;
<a href="?Action=<%=Action%>&Assort=2"><b>ʹ��ϵͳ�Դ�ͼƬ</b></a>&nbsp;&nbsp;&nbsp;&nbsp;
<a href="?Action=<%=Action%>&Assort=3"><b>ʹ��ϵͳĬ��ͼƬ</b></a>&nbsp;&nbsp;&nbsp;&nbsp;
</td>
    </tr>

<%
If Assort=1 Then
%>
<tr id="a_1"> 
<td colspan="2">  
<iframe style="top:2px" ID="UploadFiles" src="CF_Upfile.asp?UserType=1&CheckCode=<%=Rs("UserCode")%>" frameborder=0 scrolling=no width="450" height="25"></iframe>
</td>
</tr>
<%
ElseIf Assort=2 Then
%>
<tr id="a_2"> 
      <td colspan="2" >
	  <table border="0" cellpadding="0" cellspacing="0" width=100%>
<tr><td height="25" colspan='6' bgcolor='#FFFFFF'>��ѡ������ͼƬ��</td></tr>
<%
'�г���Ч��Ϊ����״̬��ͼƬ
Sql="Select * From CFCount_Upfile Order By AddTime Desc"

Set Rs=Server.CreateObject("Adodb.RecordSet")
Rs.open Sql,Conn,1,1

linenum=3
tdwidth=int(100/linenum)&"%"
jishu=1

If Not rs.eof then
 If ChkStr(Request("Page"),2) = "" Then
  Page = 1
 Else 
  Page = CInt(ChkStr(Request("Page"),2))
 End If
 
 rs.pagesize=6
 totalrs=rs.RecordCount
 totalpage=rs.pageCount
 mypagesize=rs.pagesize
 rs.absolutepage=page
End If
While Not Rs.Eof And MyPageSize>0

if jishu mod linenum=1 or linenum=1 then
Response.write "<tr>"
end if%>
<td width="<%=tdwidth%>" valign="top">
<%
Response.write "<img src='Upload/"&Rs("FileName")&"'><br>"
 
Response.write "<input type='radio' name='filename_2' value='"&Rs("FileName")&"'"
If ImgFileName=Rs("FileName") Then Response.write " checked"
Response.write ">ѡ�ô�ͼ<br>"

Response.write "�ļ���С��" & CInt(Rs("FileSizeNum")/1024) &"k<br>"

Response.write "ʹ��������"
Sql="Select Count(*) From CFCount_User Where ImgFileName='"&Rs("FileName")&"'"
Set Rs2=Conn.Execute(Sql)
Response.write Rs2(0)
Rs2.Close()
Response.write "<br>"

Response.write "�ϴ�ʱ�䣺"&Rs("AddTime")&"<br>"
%>
</td>
<%
if jishu mod linenum=0 then
Response.write "</tr>"
Response.write "<tr><td colspan='6' height='1' bgcolor='#FF0000'></td></tr>"
end if
jishu=jishu+1
mypagesize=mypagesize-1
Rs.movenext
wend 'дtr��td,jishu mod ����Ϊ1ʱ��ʼtr,Ϊ0ʱ����tr

jishu=jishu-1
if jishu mod linenum <> 0 then
for i= 1 to linenum-(jishu mod linenum)
	response.write "<td width='"&tdwidth&"'>&nbsp;</td>"
	if  i = linenum-(jishu mod linenum) then response.write "</tr>"
next
end if '�ж����һ��tr�Ƿ����ñպ�,��������td,��������ո�
%>
      </table></td>
    </tr>
<%
ElseIf Assort=3 Then
%>
<tr> 
<td colspan="2">
ʹ�����Ĭ��ͼƬ<br />
<img src="upload/default.jpg" border="0"></td>
</tr>

<%
End If
%>

<%
If Assort=2 or Assort=3 Then
%>
    <tr id="a_3"> 
      <td colspan="2"> <input type="submit" name="Submit" value="ȷ��"> 
      </td>
    </tr>
<%End If%>

</form>

<%
If Assort=2 Then
%>

<tr id="a_4"> 
      <td colspan="2">
	  <table class="tb_3">
  <tr>
    <td><a href="?Action=<%=Action%>&Assort=<%=Assort%>&Page=1">��һҳ</a>
            <%
if page>1 then%>
            <a href='?Action=<%=Action%>&Assort=<%=Assort%>&Page=<%=Page-1%>'>��һҳ</a>
            <%
end if
%>
            <%
if page<rs.pagecount   then%>
            <a href='?Action=<%=Action%>&Assort=<%=Assort%>&Page=<%=Page+1%>'>��һҳ</a>
            <%

end if
%>
            <a href="?Action=<%=Action%>&Assort=<%=Assort%>&Page=<%=totalpage%>">���һҳ</a> ҳ�Σ�<font color="#ff0000"><%=page%></font>/<%=totalpage%>ҳ&nbsp;&nbsp;��<%=totalrs%>����¼&nbsp;&nbsp;ÿҳ��ʾ<%=rs.pagesize%>��
<%
Response.write "&nbsp;&nbsp;ת����<select id='page' onChange=""window.location=document.getElementById('page').options[document.getElementById('page').selectedIndex].value"">"
For I=1 To TotalPage
    Response.Write "<option value=?Action="&Action&"&Assort="&Assort&"&Page="&i
 If Page=I Then Response.Write " selected"
 Response.Write ">"& I &"</option>"
Next
Response.write "</select>ҳ"
%>
</td>
  </tr>
</table></td>
  </tr>
<%End If%>

</table>


<%end if%>
		
		<%If Action="jsqreset" Then%>
        <div class="filters">
			 <h3>ͳ������</h3>
       </div>
        <table width="100%" >


          <form name="form3" method="post" action="?Action=jsqresetsave">
          <tr >
              <td colspan="2"> ������������������������ͳ�ƣ������ز�����</td>
            </tr>
            <tr > 
              <td width="18%">������Դͳ�ƣ�</td>
              <td ><input name="ly" type="checkbox" id="ly" value="-1">
                ȫ��ɾ��</td>
            </tr>
            <tr > 
              <td >�����շ���ͳ�ƣ�</td>
              <td ><input name="daytj" type="checkbox" id="daytj" value="-1">
                ȫ��ɾ��</td>
            </tr>
            <tr > 
              <td >����Сʱͳ�ƣ�</td>
              <td ><input name="hourtj" type="checkbox" id="hourtj" value="-1">
                ȫ��ɾ��</td>
            </tr>
            <tr > 
              <td >��ͷ��ͳ�ƣ�</td>
              <td ><input name="backtj" type="checkbox" id="backtj" value="-1">
                ȫ��ɾ��</td>
            </tr>
            <tr > 
              <td >Alexaͳ�ƣ�</td>
              <td ><input name="alexatj" type="checkbox" id="alexatj" value="-1">
                ȫ��ɾ��</td>
            </tr>
            <tr > 
              <td >������������ͳ�ƣ�</td>
              <td > 
                <input name="searchtj" type="checkbox" id="searchtj" value="-1">
                ȫ��ɾ��</td>
            </tr>
            <tr > 
              <td >��������ؼ���ͳ�ƣ�</td>
              <td >  
                <input name="keywordtj" type="checkbox" id="keywordtj" value="-1">
                ȫ��ɾ��</td>
            </tr>
			<tr > 
              <td >ҳ��ͳ�ƣ�</td>
              <td ><input name="webtj" type="checkbox" id="webtj" value="-1">
                ȫ��ɾ��</td>
            </tr>
            <tr > 
              <td >ʵ����ʾ����</td>
              <td > <input name="RealShowTotal" type="checkbox" id="hidden" value="-1">
                ����Ϊ0</td>
            </tr>
            <tr > 
              <td >ʵ��IP��ֵ��</td>
              <td ><input name="RealIpTotal" type="checkbox" id="RealIpTotal" value="-1">
                ����Ϊ0</td>
            </tr>
            <tr >
              <td >�����������ȷ�ϣ�</td>
              <td ><input name="Pwd" type="password" id="Pwd"></td>
            </tr>
            <tr > 
              <td >&nbsp;</td>
              <td > 
                <input type="submit" name="Submit2" value="�ء���" onClick="{if(confirm('ȷ��Ҫ���ü���ô��������ɾ�����һЩѡ�������!')){return true;}return false;};">
                �� 
                <input type="reset" name="Submit22" value="ȡ����" >
              </td>
            </tr>
          </form>
        </table>
<%End If%>

<%If Action="userdel" Then%>
<div class="filters">
			 <h3>�˺�ɾ��</h3>
       </div>
        <table width="100%" >

<form name="form3" method="post" action="?Action=userdelsave">
          <tr >
              <td colspan="2">��ȷ���Ƿ����Ҫɾ����</td>
    </tr>
            <tr > 
              <td width="18%">�����������ȷ�ϣ�</td>
              <td ><input name="Pwd" type="password" id="Pwd"></td>
            </tr>
            <tr > 
              <td >&nbsp;</td>
              <td > 
                <input type="submit" name="Submit2" value="ɾ����" onClick="{if(confirm('ȷ��Ҫɾ���Լ����˺�ô������˺��ڼ������в�������!')){return true;}return false;};">
                �� 
                <input type="reset" name="Submit22" value="ȡ����" >
              </td>
            </tr>
          </form>
        </table>
<%End If%>


        <%If Action="pwdmodify" Then%>
<%Set Rs= Server.CreateObject("Adodb.RecordSet")
Sql="Select * From CFCount_User Where UserName='"&UserName&"'"
Rs.Open Sql,Conn,1,1%>
<div class="filters">
			 <h3>�޸�����(��*Ϊ�����</h3>
       </div>
        <table width="100%" >

  <form name="form3" method="post" action="?Action=pwdmodifysave">
    <tr > 
      <td width="20%">�������¹������룺</td>
      <td > <input name="Pwd" type="password" id="Pwd_Old" size="15">
        (��������޸�������)</td>
    </tr>
    <tr > 
      <td >ȷ���¹������룺</td>
      <td > <input name="Pwd2" type="password" id="Pwd" size="15">
      </td>
    </tr>
    <tr > 
      <td >�������²鿴���룺</td>
      <td > <input name="Pwd_View" type="password" id="Pwd_Old" size="15">
        (����鿴�޸�������) </td>
    </tr>
    <tr > 
      <td >ȷ���²鿴���룺</td>
      <td > <input name="Pwd_View2" type="password" id="Pwd_Old" size="15">
      </td>
    </tr>
    <tr > 
      <td >������ɹ�������ȷ�ϣ�</td>
      <td >
        <input name="Pwd_Old" type="password" id="Pwd2" size="15">
        <font color="#0000FF">*</font></td>
    </tr>
    <tr > 
      <td >&nbsp;</td>
      <td >
        <input type="submit" name="Submit2" value="�ޡ���" >
        �� 
        <input type="reset" name="Submit22" value="ȡ����" >
      </td>
    </tr>
    <tr > 
      <td colspan="2" >ע�⣺��ʼ�鿴����͸�ע��ʱ�Ĺ�������һ���������óɺ͹������벻һ��</td>
    </tr>
  </form>
</table>
<%End if%>

<%If Action="datamodify" Then%>
          <script language="JavaScript">
function modifyuser()
{
if ((document.form2.email.value)=="")
{
window.alert ('email������д');
document.form2.email.focus();
return false;
}
else if (!isValidEmail(document.form2.email.value)){
 window.alert ('email�ĸ�ʽ����ȷ!');
 document.form2.email.focus();
 return false;
}
else if ((document.form2.pagename.value)=="")
{
window.alert ('��ҳ���Ʊ�����д');
document.form2.pagename.focus();
return false;
}
else if ((document.form2.pageurl.value)=="")
{
window.alert ('��ҳ��ַ������д');
document.form2.pageurl.focus();
return false;
}
else
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

<%Set Rs= Server.CreateObject("Adodb.RecordSet")
Sql="Select * From CFCount_User Where UserName='"&UserName&"'"
Rs.Open Sql,Conn,1,1%>
<div class="filters">
			 <h3>�޸�ע������(��*Ϊ�����</h3>
       </div>
        <table width="100%" >

  <form name="form2" method="post" action="?Action=datamodifysave" onsubmit="return modifyuser()">
    <tr > 
      <td width="18%">E- mail��</td>
      <td > <input name="email" type="text"  value="<%=rs("email")%>" size="30"> 
        <font color="#0000FF">*</font></td>
    </tr>
    <tr > 
      <td >QQ ���룺</td>
      <td > <input name="qq" type="text" id="qq"  value="<%=rs("qq")%>" size="30"> 
      </td>
    </tr>
    <tr > 
      <td >��Ŀ���ƣ�</td>
      <td > <input name="pagename" type="text"  value="<%=rs("pagename")%>" size="30"> 
        <font color="#0000FF">*</font></td>
    </tr>
    <tr > 
      <td >��Ŀ��ַ��</td>
      <td > <input name="pageurl" type="text"  value="<%=rs("pageurl")%>" size="30"> 
        <font color="#0000FF">*</font></td>
    </tr>
    <tr >
      <td >����������ȷ���޸ģ�</td>
      <td ><input name="Pwd" type="password" size="30"><font color="#0000FF">*</font></td>
    </tr>
    <tr > 
      <td >&nbsp;</td>
      <td > <input type="submit" name="Submit2" value="�ޡ���" >
        �� 
        <input type="reset" name="Submit22" value="ȡ����" > </td>
    </tr>
  </form>
</table>
<%End If%>
        <%If Action="passwordanswermodify" Then%>
        <%Set Rs= Server.CreateObject("Adodb.RecordSet")
Sql="Select * From CFCount_User Where UserName='"&UserName&"'"
Rs.Open Sql,Conn,1,1%>
<div class="filters">
			 <h3>�޸�������ʾ�ʹ�(��*Ϊ�����</h3>
       </div>
        <table width="100%" >


          <form name="form3" method="post" action="?Action=passwordanswermodifysave">
            <%If Rs("passwordask")<>"" Then%>
            <tr > 
              <td width="16%">ԭ������ʾ���⣺</td>
              <td ><%=Rs("passwordask")%></td>
            </tr>
            <tr > 
              <td >ԭ����ش�𰸣�</td>
              <td ><input name="passwordanswer_old" type="text" id="passwordanswer" size="20">
                <font color="#0000FF">*</font>(������ԭ��ȷ��Ҫ�޸�Ϊ��������ʾ�ʹ�)</td>
            </tr>
            <%End If%>
            <tr > 
              <td width="16%">��������ʾ���⣺</td>
              <td > 
                <input name="passwordask_new" type="text" id="passwordask" size="20">
                <font color="#0000FF">*</font> </td>
            </tr>
            <tr > 
              <td >������ش�𰸣�</td>
              <td > 
                <input name="passwordanswer_new" type="text" id="passwordanswer" size="20">
                <font color="#0000FF">*</font></td>
            </tr>
            <tr > 
              <td >&nbsp;</td>
              <td > 
                <input type="submit" name="Submit23" value="�ޡ���" >
                �� 
                <input type="reset" name="Submit222" value="ȡ����" >
              </td>
            </tr>
          </form>
        </table>
        <%End if%>
        <%If Action="stylelist" Then%>
<%
 If Session("CFCountUser_View")<>"" Then'Ϊ�����鿴��ʱ
  Call AlertBack("�����鿴���޷��Դ˲���",1)
 End If
%>
<%
Call GetCurrWeb()
IF not IsNumeric(Request("Page")) Or IsEmpty(Request("Page")) Then
 Page=1
Else
 Page=Int(Abs(Request("Page")))
End If

MyPageSize=10
MyPageSize2=10

If RsSet("StyleTotal") Mod MyPageSize=0 Then
 TotalPage=RsSet("StyleTotal")/MyPageSize
Else
 TotalPage=Int(RsSet("StyleTotal")/MyPageSize)+1
 If Page=TotalPage Then
 MyPageSize2=RsSet("StyleTotal") Mod MyPageSize
 End If
End If
I=1
%>
<%
Set Rs=Server.CreateObject("Adodb.RecordSet")
Sql="Select * From CFCount_User Where UserName='"&UserName&"'"
Rs.Open Sql,Conn,1,1%>
  <div class="filters">
			 <h3>��ʽ����-ѡ����ϲ����ͳ����ʽ</h3>
       </div>
        <table width="100%" >
          <form name="form4" method="post" action="?Action=stylemodifysave">
            <%While I<=MyPageSize2%>
            <tr> 
              <td class="td_2"> 
                  <%K=(Page-1)*MyPageSize+I%>
                  <input name="style" type="radio" value="<%=k%>"<%If Rs("Style")=K Then Response.Write " Checked"%>>
              </td>
              <td >
<div ><%For J=0 to 9%><img src="CounterPic/<%=K%>/<%=J%>.gif"><%Next%></td>
            </tr>
            <%
I=I+1
Wend%>
            <tr>
              <td width="10%">ȷ��ʹ�ã�</td>
              <td ><div >
                  <input type="checkbox" name="CounterShow" value="-1"<%If Rs("CounterShow")=-1 Then Response.Write " checked"%>>��
              </td>
            </tr>
            <tr> 
              <td >&nbsp; </td>
              <td > <div > 
                  <input type="submit" name="Submit3" value="ȷ��">
              </td>
            </tr>
          </form>
          <tr> 
            <td colspan="2" ><table class="tb_3">
                <tr> 
                  <td><a href="?action=<%=action%>&page=1">��һҳ</a> 
                      <%
if page>1 then%>
                      <a href='?action=<%=action%>&page=<%=page-1%>'>��һҳ</a> 
                      <%
end if
%>
                      <%
if page<totalpage   then%>
                      <a href='?action=<%=action%>&page=<%=page+1%>'>��һҳ</a> 
      <%
end if
%>
                      <a href="?action=<%=action%>&page=<%=totalpage%>">���һҳ</a> 
                      ҳ�Σ�<font color="#ff0000"><%=page%></font>/<%=totalpage%>ҳ
                            
<%
Response.write "&nbsp;&nbsp;ת����<select id='page' onChange=""window.location=document.getElementById('page').options[document.getElementById('page').selectedIndex].value"">"
For I=1 To TotalPage
    Response.Write "<option value=?action="&action&"&page="&i     
 If Page=I Then Response.Write " selected"
 Response.Write ">"& I &"</option>"
Next
Response.write "</select>ҳ"
%></td>
              </tr>
                    </table></td>
          </tr>
</table>
          <table >
           <tr >
            <td colspan="2">���ͳ��Ч�� 
            </td>
          </tr>
          <tr > 
            <td colspan="2"  >Script���ã� 
            </td>
          </tr>
          <tr > 
            <td colspan="2"  ><script src="<%=Tmp%>cf.asp?UserName=<%=Rs("UserName")%>"></script></td>
          </tr>
        </table>
<%end if%>



<%If Action="gbooklist" Then%>
<%
 If Session("CFCountUser_View")<>"" Then'Ϊ�����鿴��ʱ
  Call AlertBack("�����鿴���޷��Դ˲���",1)
 End If
%>
<div class="filters">
			 <h3>�����б�</h3>
       </div>
        <table width="100%" >
          <tr > 
            <td class="td_nowrap">ID</td>
            <td class="td_nowrap">��������</td>
            <td class="td_nowrap">��ϵ��ʽ</td>
            <td class="td_nowrap">����ʱ��</td>
			<td class="td_nowrap">������Դ</td>
			<td class="td_nowrap">���Ե�ǰҳ</td>
            <td class="td_nowrap">����</td>
          </tr>
<%

Sql="Select * From CFCount_Gbook where UserName='"&UserName&"' Order By AddTime Desc"

Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.open Sql,Conn,1,1

If Not rs.eof then
 IF not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Then
  page=1
 Else
  Page=Int(Abs(Request("page")))
 End if
 rs.pagesize=15
 totalrs=rs.RecordCount
 totalpage=rs.pageCount
 mypagesize=rs.pagesize
 rs.absolutepage=page
End If
%>
          <%
While Not Rs.Eof And MyPageSize>0
%>
          <tr class="td_2"> 
            <td class="td_nowrap"><%=Rs("ID")%></td>
            <td><%=Rs("Content")%></td>
            <td class="td_nowrap"><%=Rs("Contact")%></td>
            <td class="td_nowrap"><%=Rs("AddTime")%></td>
			<td><a href="<%=Rs("Ly")%>" target="_blank"><%=Rs("Ly")%></a></td>
			<td><a href="<%=Rs("CurrWeb")%>" target="_blank"><%=Rs("CurrWeb")%></a></td>
            <td class="td_nowrap"><a href="?Action=gbookmodify&ID=<%=Rs("ID")%>">�޸�</a> <a href="?Action=gbookdel&ID=<%=Rs("ID")%>" onClick="{if(confirm('ȷ��Ҫɾ��ô?')){return true;}return false;}">ɾ��</a></td>
          </tr>
          <%rs.movenext
mypagesize=mypagesize-1
wend%>
        </table>
		
<table class="tb_3">
  <tr>
    <td><a href="?Action=<%=Action%>&Page=1">��һҳ</a>
            <%
if page>1 then%>
            <a href='?Action=<%=Action%>&Page=<%=Page-1%>'>��һҳ</a>
            <%
end if
%>
            <%
if page<rs.pagecount   then%>
            <a href='?Action=<%=Action%>&Page=<%=Page+1%>'>��һҳ</a>
            <%

end if
%>
            <a href="?Action=<%=Action%>&Page=<%=totalpage%>">���һҳ</a> ҳ�Σ�<font color="#ff0000"><%=page%></font>/<%=totalpage%>ҳ&nbsp;&nbsp;��<%=totalrs%>����¼&nbsp;&nbsp;ÿҳ��ʾ<%=rs.pagesize%>��
<%
Response.write "&nbsp;&nbsp;ת����<select id='page' onChange=""window.location=document.getElementById('page').options[document.getElementById('page').selectedIndex].value"">"
For I=1 To TotalPage
    Response.Write "<option value=?Action="&Action&"&Page="&i
 If Page=I Then Response.Write " selected"
 Response.Write ">"& I &"</option>"
Next
Response.write "</select>ҳ"
%>
      </td>

  </tr>
</table>
		  
<%End If%>
		  

<%If Action="gbookmodify" Then%>
<%
 If Session("CFCountUser_View")<>"" Then'Ϊ�����鿴��ʱ
  Call AlertBack("�����鿴���޷��Դ˲���",1)
 End If
%>
<%
ID=ChkStr(Request("ID"),1)

Sql="Select * From CFCount_Gbook Where UserName='"&UserName&"' And ID="&ID
Set Rs=Conn.Execute(Sql)
%>
        
<table width="100%" >
  <form name="form1" method="post" action="?Action=gbookmodifysave&id=<%=ID%>">
    <tr > 
      <td colspan="2">�����޸�</td>
    </tr>
    <tr> 
      <td class="td_wauto">�������ݣ�</td>
      <td><textarea name='content' id='content' cols='40' rows='8'><%=server.HTMLEncode(Rs("content"))%></textarea></td>
    </tr>
    <tr> 
      <td>��ϵ��ʽ��</td>
      <td> <textarea name='contact' id='contact' cols='40' rows='8'><%=server.HTMLEncode(Rs("contact"))%></textarea> </td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><input type="submit" name="Submit" value="�޸�">      </td>
    </tr>
  </form>
</table>
        
<%End If%>
			  

<%If Action="getcode" Then%>

<script language=javascript>
function copy(ob){
var obj=findObj(ob); if (obj) { 
obj.select();js=obj.createTextRange();js.execCommand("Copy");}
alert('�ɹ������븴�Ƶ��������У�');
}
function cut(ob){
var obj=findObj(ob); if (obj) { 
obj.select();js=obj.createTextRange();js.execCommand("Cut");}
}
function findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}
//-->
</script>
 <div class="filters">
			 <h3>���ͳ�ƴ��뼰ͳ��Ԥ��</h3>
       </div>
        <table >
  <tr > 
    <td  >���ִ��뷽ʽ�ɹ�ѡ��<input type="submit" name="Submit5" value="�����෽ʽ����" onclick="location.href='?Action=getcode'"<%If Action="getcode" Then%> style="BACKGROUND-COLOR: #e8e8e8;"<%End If%>> 
      <input type="submit" name="Submit6" value="��վ��ʽ����" onclick="location.href='?Action=getcode_2'"<%If Action="getcode_2" Then%> style="BACKGROUND-COLOR: #e8e8e8;"<%End If%>></td>
  </tr>
  <tr > 
    <td  >Img��ʽ���ã��˷�ʽ�ʺϷ����������ʼ������ݵ��У������ʹ�õ����ǵ��ʼ�Ⱥ��ϵͳ������Ҫ�����ã����ǽ��Զ�����ͳ�ơ�</td>
  </tr>
  <tr > 
    <td  >���·����������ͳ�ƴ��룬Img���÷�ʽ���븴�ƺ���뵽�ʼ�����HTMLԴ������,���ַ�ʽ��������԰����ͳ�ƿ�����һ��ͼƬ��ͼƬ��ַ��<a href="<%=Tmp%>cf.asp?c=<%=UserName%>" target="_blank"><%=Tmp%>cf.asp?c=<%=Rs("UserName")%></a>����������ʱ��ʾ��Ϊ0��<br> <textarea name="countercode2" cols="50" rows="6" id="textarea"><a href="<%=Tmp%>Index.asp" target="_blank"><img src="<%=Tmp%>cf.asp?c=<%=UserName%>" border="0" title="<%=server.HTMLEncode(GetUnicode(RsSet("SysTitle")))%>"></a></textarea> 
      <input type="submit" name="Submit4" value="����" onClick=copy('countercode2')> 
      <br> <a href="<%=Tmp%>Index.asp" target="_blank"><img src="<%=Tmp%>cf.asp?c=<%=UserName%>" border="0" title="<%=RsSet("SysTitle")%>"></a> 
  </tr>
</table>
<%end if%>

<%If Action="getcode_2" Then%>

<script language=javascript>
function copy(ob){
var obj=findObj(ob); if (obj) { 
obj.select();js=obj.createTextRange();js.execCommand("Copy");}
alert('�ɹ������븴�Ƶ��������У�');
}
function cut(ob){
var obj=findObj(ob); if (obj) { 
obj.select();js=obj.createTextRange();js.execCommand("Cut");}
}
function findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}
//-->
</script>
  
   <div class="filters">
			 <h3>���ͳ�ƴ��뼰ͳ��Ԥ��</h3>
       </div>
        <table >
  <tr > 
    <td  >���ִ��뷽ʽ�ɹ�ѡ��<input type="submit" name="Submit5" value="�����෽ʽ����" onclick="location.href='?Action=getcode'"<%If Action="getcode" Then%> style="BACKGROUND-COLOR: #e8e8e8;"<%End If%>> 
      <input type="submit" name="Submit6" value="��վ��ʽ����" onclick="location.href='?Action=getcode_2'"<%If Action="getcode_2" Then%> style="BACKGROUND-COLOR: #e8e8e8;"<%End If%>></td>
  </tr>
  <tr > 
    <td  >Script��վ��ʽ���ã��˷�ʽ�ʺϷ����Լ�����վҳ���ϣ��м�ʮ��ͼƬ��ʽ����ѡ�����ڲ˵�"<a href="?Action=stylelist" target="_blank">ͳ������>>ͼƬ��ʽѡ��</a>"��ѡ��!����ͳ�Ƶ�һЩ�����������ڲ˵�"<a href="?Action=jsqset" target="_blank">ͳ������>>��������</a>"�����ã�����������ͳ��!</td>
  </tr>
  <tr > 
    <td  >���·����������ͳ�ƴ���,script���÷�ʽ���븴�ƺ���뵽��ҳԴ������:<br> 
      <textarea name="countercode1" cols="50" rows="4"><script src="<%=Tmp%>cf.asp?username=<%=UserName%>"></script></textarea> 
      <input type="submit" name="Submit4" value="����" onClick=copy('countercode1')> 
      <br>
      ������õ�ͳ��script��ʽ����Ԥ���� <br> <script src="<%=Tmp%>cf.asp?username=<%=UserName%>"></script> 
    </td>
  </tr>
</table>  
<%end if%>