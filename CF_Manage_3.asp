
<%If Action="jsqset" Then%>
<%
 If Session("CFCountUser_View")<>"" Then'为独立查看者时
  Call AlertBack("独立查看者无法对此操作",1)
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
			 <h3>参数设置</h3>
       </div>
        <table width="100%" >
         
  <form name="form3" method="post" action="?Action=jsqsetsave">
 <tr >
      <td colspan="2"> 配置以下参数(带*为必填项）</td>
    </tr>
    <tr > 
      <td width="24%">自定义显示数：</td>
      <td > <input name="ShowTotal" type="text" id="ShowTotal" value="<%=rs("ShowTotal")%>" size="30" > 
        <font color="#0000FF">*</font>(自定义)</td>
    </tr>
    <tr > 
      <td >实际显示数：</td>
      <td ><%=rs("RealShowTotal")%>&nbsp;&nbsp;(页面的浏览次数)</td>
    </tr>
    <tr > 
      <td >实际数值：</td>
      <td ><%=rs("RealIpTotal")%>&nbsp;&nbsp;(IP数)</td>
    </tr>
    <tr > 
      <td >统计显示位数：</td>
      <td ><input name="PicNum" type="text" id="total2" value="<%=rs("PicNum")%>" size="30" > 
        <font color="#0000FF">*</font></td>
    </tr>
    <tr > 
      <td >统计开始日期：</td>
      <td > <input name="adddate" type="text" id="adddate" value="<%=rs("adddate")%>" size="30" > 
        <font color="#0000FF">*</font></td>
    </tr>
    <tr > 
      <td >Script统计显示数字：</td>
      <td > <input type="radio" name="CounterShow" value="-1"<%if rs("CounterShow")=-1 then response.write " checked"%> onclick="ashow_1()">
        是 
        <input type="radio" name="CounterShow" value="0"<%if rs("CounterShow")=0 then response.write " checked"%> onclick="ashow_2()">
        否</td>
    </tr>
    <tr  id="a_1"> 
      <td >Script统计显示数字：</td>
      <td > <input name="ShowType" type="radio" id="ShowType" value="1" <%if rs("ShowType")=1 then response.write "checked"%>>
        自定义显示数 
        <input name="ShowType" type="radio" id="ShowType" value="2" <%if rs("ShowType")=2 then response.write "checked"%>>
        实际显示数值 
        <input name="ShowType" type="radio" id="ShowType" value="3" <%if rs("ShowType")=3 then response.write "checked"%>>
        实际Ip数值 </td>
    </tr>
    <tr  id="a_2"> 
      <td >你设置的计数样式：</td>
      <td ><%=GetPicStyle(UserName)%>&nbsp;[<a href="?Action=stylelist">更换图片样式</a>] </td>
    </tr>
    <tr  id="a_3"> 
      <td >Script统计隐藏时默认显示：</td>
      <td ><input name="CounterHiddenPic" type="radio" id="radio" value="0" <%if rs("CounterHiddenPic")=0 then response.write "checked"%>> 
        留言 
		<input name="CounterHiddenPic" type="radio" id="radio" value="1" <%if rs("CounterHiddenPic")=1 then response.write "checked"%>> 
        <%=RsSet("TjTextName")%> <input name="CounterHiddenPic" type="radio" id="radio" value="2" <%if rs("CounterHiddenPic")=2 then response.write "checked"%>>
        图片1 <img src="images/counter_2.gif" border="0"> <input name="CounterHiddenPic" type="radio" id="radio" value="3" <%if rs("CounterHiddenPic")=3 then response.write "checked"%>>
        图片2 <img src="images/counter_3.gif" border="0"> </td>
    </tr>
    <tr > 
      <td >Script统计显示位置：</td>
      <td ><input name="CounterSite" type="radio" id="radio" value="1" <%if rs("CounterSite")=1 then response.write "checked"%>>
        在正在阅读上面 
        <input name="CounterSite" type="radio" id="radio" value="2" <%if rs("CounterSite")=2 then response.write "checked"%>>
        在正在阅读下面 
        <input name="CounterSite" type="radio" id="radio" value="3" <%if rs("CounterSite")=3 then response.write "checked"%>>
        在正在阅读左面 
        <input name="CounterSite" type="radio" id="radio2" value="4" <%if rs("CounterSite")=4 then response.write "checked"%>>
        在正在阅读右面 </td>
    </tr>
    <tr > 
      <td >Img统计显示数字：</td>
      <td > <input type="radio" name="ImgCounterShow" value="-1"<%if rs("ImgCounterShow")=-1 then response.write " checked"%> onclick="bshow_1()">
        是 
        <input type="radio" name="ImgCounterShow" value="0"<%if rs("ImgCounterShow")=0 then response.write " checked"%> onclick="bshow_2()">
        否</td>
    </tr>
    <tr  id="b_1"> 
      <td >Img统计显示数字：</td>
      <td ><input name="ImgShowType" type="radio" id="radio4" value="1" <%if rs("ImgShowType")=1 then response.write "checked"%>>
        自定义显示数 
        <input name="ImgShowType" type="radio" id="radio4" value="2" <%if rs("ImgShowType")=2 then response.write "checked"%>>
        实际显示数值 
        <input name="ImgShowType" type="radio" id="radio4" value="3" <%if rs("ImgShowType")=3 then response.write "checked"%>>
        实际Ip数值</td>
    </tr>
    <tr > 
      <td >记录正在阅读：</td>
      <td > <input type="radio" name="online" value="-1"<%if rs("online")=-1 then response.write " checked"%> onclick="cshow_1()">
        是 
        <input type="radio" name="online" value="0"<%if rs("online")=0 then response.write " checked"%> onclick="cshow_2()">
        否</td>
    </tr>
    <tr  id="c_1"> 
      <td >记录正在阅读的间隔：</td>
      <td ><input name="onlinetime" type="text" id="onlinetime" value="<%=rs("onlinetime")%>" size="30">
        分钟(请设置180分钟以内,一般是30分钟)</td>
    </tr>
    <tr > 
      <td >统计上显示正在阅读：</td>
      <td > <input name="onlineshow" type="checkbox" id="onlineshow" value="-1" <%if rs("onlineshow")=-1 then response.write "checked"%>>
        是</td>
    </tr>
    <tr > 
      <td >统计上显示今日浏览:</td>
      <td ><input name="TodayShow" type="checkbox" id="TodayShow" value="-1" <%if rs("TodayShow")=-1 then response.write "checked"%>>
        是</td>
    </tr>
    <tr > 
      <td >统计上显示今日IP:</td>
      <td ><input name="TodayIpShow" type="checkbox" id="TodayIpShow" value="-1" <%if rs("TodayIpShow")=-1 then response.write "checked"%>>
        是</td>
    </tr>
    <tr > 
      <td >统计上显示IP：</td>
      <td ><input name="IpShow" type="checkbox" id="IpShow" value="-1" <%if rs("IpShow")=-1 then response.write "checked"%>>
        是</td>
    </tr>
    <tr > 
      <td >统计上显示来访次数：</td>
      <td ><input name="VisitShow" type="checkbox" id="VisitShow" value="-1" <%if rs("VisitShow")=-1 then response.write "checked"%>>
        是</td>
    </tr>
 <tr >
      <td colspan="2" >以下是信息是否公开设置(公开后所有人可查看)：</td>
    </tr>
    <tr > 
      <td >统计信息公开：</td>
      <td ><input name="TjOpen" type="radio" id="TjOpen" value="-1" <%if rs("TjOpen")=-1 then response.write "checked"%> onclick="dshow_1()">
        是　 
        <input name="TjOpen" type="radio" id="radio3" value="0" <%if rs("TjOpen")=0 then response.write "checked"%> onclick="dshow_2()">
        否</td>
    </tr>
    <tr  id="d_1"> 
      <td >公开选中的统计信息：</td>
      <td ><input name="SurveyOpen" type="checkbox" id="SurveyOpen" value="-1" <%if rs("SurveyOpen")=-1 then response.write "checked"%>>
        统计概况 
        <input name="TodayLyOpen" type="checkbox" id="TodayLyOpen" value="-1" <%if rs("TodayLyOpen")=-1 then response.write "checked"%>>
        今天的来源明细 
        <input name="OnlineOpen" type="checkbox" id="OnlineOpen" value="-1" <%if rs("OnlineOpen")=-1 then response.write "checked"%>>
        正在阅读 
        <input name="TodayHourOpen" type="checkbox" id="TodayHourOpen" value="-1" <%if rs("TodayHourOpen")=-1 then response.write "checked"%>>
        今天的小时统计 
        <input name="EveryDayOpen" type="checkbox" id="EveryDayOpen" value="-1" <%if rs("EveryDayOpen")=-1 then response.write "checked"%>>
        每天的统计</td>
    </tr>
 <tr >
      <td colspan="2" >留言板</td>
    </tr>
	<tr > 
      <td >中转页上显示留言板：</td>
      <td ><input name="GbookState" type="radio" id="GbookState" value="-1" <%if rs("GbookState")=-1 then response.write "checked"%>>
        是　 
        <input name="GbookState" type="radio" id="GbookState" value="0" <%if rs("GbookState")=0 then response.write "checked"%>>
        否</td>
    </tr>

	<tr > 
      <td >&nbsp;</td>
      <td > <input type="submit" name="Submit2" value="修　改" >
        　 
        <input type="reset" name="Submit22" value="取　消" > </td>
    </tr>
  </form>
 <tr >
    <td colspan="2">你的统计效果 
    </td>
  </tr>
  <tr > 
    <td colspan="2"  >Script调用： 
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
 window.alert("请选择一种基本样式!");
 return false;
}
if (((document.form2.quickstyle(0).checked)||(document.form2.quickstyle(1).checked)||(document.form2.quickstyle(2).checked)||(document.form2.quickstyle(6).checked))&&(form2.tjopen.options[form2.tjopen.selectedIndex].value)==""){
 window.alert ('请选择是否公开统计!');
 document.form2.tjopen.focus();
 return false;
}
return true;
}
</script>
  <div class="filters">
			 <h3>统计基本样式快速设置</h3>
       </div>
        <table width="100%" >


          <form name="form2" method="post" action="?Action=jsqquicksetsave"  onsubmit="return jsqquickset()">
            <tr > 
              <td width="120"  >数字型：</td>
              <td  ><input type="radio" name="quickstyle" value="1" onclick="ashow_1()"></td>
              <td  ><%=GetPicStyle(UserName)%></td>
            </tr>
            <tr > 
              <td  >图标型：</td>
              <td  ><input type="radio" name="quickstyle" value="2" onclick="ashow_2()"></td>
              <td  > <a href=<%=Tmp%> target=_blank title=计数服务由[<%=RsSet("SysTitle")%>]免费提供><img src=<%=Tmp%>images/counter_2.gif border='0'></a></td>
            </tr>
            <tr > 
              <td  >文字型：</td>
              <td  ><input type="radio" name="quickstyle" value="3" onclick="ashow_3()"></td>
              <td  ><a href=<%=HttpPath(2)%> target=_blank title=统计服务由[<%=RsSet("SysTitle")%>]免费提供><%=RsSet("TjTextName")%></a></td>
            </tr>
            <tr > 
              <td  >数字型+详细描述：</td>
              <td  ><input type="radio" name="quickstyle" value="4"  onclick="ashow_4()"></td>
              <td  > <table border='0' cellpadding='2' cellspacing='1'>
                  <tr> 
                    <td><div align='center'><%=GetPicStyle(UserName)%></td>
                    <td><div align='center'><a href=<%=Tmp%>View.asp?UserName=<%=UserName%> target=_blank title='网站名称：康佑原创程序&#13;今天浏览：73&#13;今天IP：16&#13;在线人数：2&#13;&#13;计数服务由[<%=RsSet("SysTitle")%>]免费提供'><span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>同时阅读[<font color='#FF0000'>2</font>]人</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>今天浏览[<font color='#FF0000'>73</font>]</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>今天访问者[<font color='#FF0000'>16</font>]</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>你的IP[<font color='#FF0000'>127.0.0.1</font>]</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>欢迎你第[<font color='#FF0000'>5</font>]次阅读邮件</span></a></td>
                  </tr>
                </table></td>
            </tr>
            <tr > 
              <td  >图标型+详细描述：</td>
              <td  ><input type="radio" name="quickstyle" value="5" onclick="ashow_5()"></td>
              <td  > <table border='0' cellpadding='2' cellspacing='1'>
                  <tr> 
                    <td><div align='center'><a href=<%=Tmp%>View.asp?UserName=<%=UserName%> target=_blank title=计数服务由[<%=RsSet("SysTitle")%>]免费提供><img src=<%=Tmp%>images/counter_2.gif border='0'></a></td>
                    <td><div align='center'><a href=<%=Tmp%>View.asp?UserName=<%=UserName%> target=_blank title='计数服务由[<%=RsSet("SysTitle")%>]免费提供'><span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>同时阅读[<font color='#FF0000'>1</font>]人</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>今天浏览[<font color='#FF0000'>47</font>]</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>今天访问者[<font color='#FF0000'>15</font>]</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>你的IP[<font color='#FF0000'>127.0.0.1</font>]</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>欢迎你第[<font color='#FF0000'>4</font>]次阅读邮件</span></a></td>
                  </tr>
                </table></td>
            </tr>
            <tr > 
              <td  >文字型+详细描述：</td>
              <td  ><input type="radio" name="quickstyle" value="6" onclick="ashow_6()"></td>
              <td   ><table border='0' cellpadding='0' cellspacing='0'>
                  <tr> 
                    <td><div align='center'><a href=<%=HttpPath(2)%> target=_blank title=统计服务由[<%=RsSet("SysTitle")%>]免费提供><%=RsSet("TjTextName")%></a></td>
                    <td><div align='center'><a href=<%=HttpPath(2)%> target=_blank title='统计服务由[<%=RsSet("SysTitle")%>]免费提供'><span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>同时阅读[<font color='#FF0000'>1</font>]人</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>今天浏览[<font color='#FF0000'>47</font>]</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>今天IP[<font color='#FF0000'>15</font>]</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>你的访问者[<font color='#FF0000'>127.0.0.1</font>]</span>&nbsp;<span style='font-size:12px;text-decoration:none;font-weight:normal;color:#000000;line-height:20px;'>欢迎你第[<font color='#FF0000'>4</font>]次阅读邮件</span></a></td>
                  </tr>
                </table></td>
            </tr>
            <tr > 
              <td  >留言型：</td>
              <td  ><input type="radio" name="quickstyle" value="7" onclick="ashow_7()"></td>
              <td  >留言板</td>
            </tr>
			<tr  id="a_1">
              <td colspan="2"  >是否公开统计：</td>
              <td   ><select name="tjopen">
                  <option value="" selected>请选择</option>
                  <option value="-1">公开</option>
                  <option value="0">不公开</option>
                </select></td>
            </tr>
            <tr > 
              <td colspan="2"  >&nbsp;</td>
              <td   > 
                <input type="submit" name="Submit2" value="设　置">
                <br>
                (注：这里只提供四种基本样式选择,更多的设置请进入＂<a href="?Action=jsqset">参数设置</a>&quot;和&quot;<a href="?Action=stylelist">图片样式选择</a>&quot;，你的设置统计实际效果在&quot;<a href="?Action=getcode">预览及获得代码</a>&quot;里可以看到)</td>
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
			 <h3>你的设置效果如下</h3>
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
<a href="?Action=<%=Action%>&Assort=1"><b>使用自定义图片</b></a>&nbsp;&nbsp;&nbsp;&nbsp;
<a href="?Action=<%=Action%>&Assort=2"><b>使用系统自带图片</b></a>&nbsp;&nbsp;&nbsp;&nbsp;
<a href="?Action=<%=Action%>&Assort=3"><b>使用系统默认图片</b></a>&nbsp;&nbsp;&nbsp;&nbsp;
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
<tr><td height="25" colspan='6' bgcolor='#FFFFFF'>请选择以下图片：</td></tr>
<%
'列出有效并为共享状态的图片
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
Response.write ">选用此图<br>"

Response.write "文件大小：" & CInt(Rs("FileSizeNum")/1024) &"k<br>"

Response.write "使用人数："
Sql="Select Count(*) From CFCount_User Where ImgFileName='"&Rs("FileName")&"'"
Set Rs2=Conn.Execute(Sql)
Response.write Rs2(0)
Rs2.Close()
Response.write "<br>"

Response.write "上传时间："&Rs("AddTime")&"<br>"
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
wend '写tr和td,jishu mod 列数为1时开始tr,为0时结束tr

jishu=jishu-1
if jishu mod linenum <> 0 then
for i= 1 to linenum-(jishu mod linenum)
	response.write "<td width='"&tdwidth&"'>&nbsp;</td>"
	if  i = linenum-(jishu mod linenum) then response.write "</tr>"
next
end if '判断最后一行tr是否正好闭合,否则增加td,里面填入空格
%>
      </table></td>
    </tr>
<%
ElseIf Assort=3 Then
%>
<tr> 
<td colspan="2">
使用这个默认图片<br />
<img src="upload/default.jpg" border="0"></td>
</tr>

<%
End If
%>

<%
If Assort=2 or Assort=3 Then
%>
    <tr id="a_3"> 
      <td colspan="2"> <input type="submit" name="Submit" value="确定"> 
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
    <td><a href="?Action=<%=Action%>&Assort=<%=Assort%>&Page=1">第一页</a>
            <%
if page>1 then%>
            <a href='?Action=<%=Action%>&Assort=<%=Assort%>&Page=<%=Page-1%>'>上一页</a>
            <%
end if
%>
            <%
if page<rs.pagecount   then%>
            <a href='?Action=<%=Action%>&Assort=<%=Assort%>&Page=<%=Page+1%>'>下一页</a>
            <%

end if
%>
            <a href="?Action=<%=Action%>&Assort=<%=Assort%>&Page=<%=totalpage%>">最后一页</a> 页次：<font color="#ff0000"><%=page%></font>/<%=totalpage%>页&nbsp;&nbsp;共<%=totalrs%>条记录&nbsp;&nbsp;每页显示<%=rs.pagesize%>条
<%
Response.write "&nbsp;&nbsp;转到第<select id='page' onChange=""window.location=document.getElementById('page').options[document.getElementById('page').selectedIndex].value"">"
For I=1 To TotalPage
    Response.Write "<option value=?Action="&Action&"&Assort="&Assort&"&Page="&i
 If Page=I Then Response.Write " selected"
 Response.Write ">"& I &"</option>"
Next
Response.write "</select>页"
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
			 <h3>统计重置</h3>
       </div>
        <table width="100%" >


          <form name="form3" method="post" action="?Action=jsqresetsave">
          <tr >
              <td colspan="2"> 将根据以下你的设置重置你的统计，请慎重操作！</td>
            </tr>
            <tr > 
              <td width="18%">所有来源统计：</td>
              <td ><input name="ly" type="checkbox" id="ly" value="-1">
                全部删除</td>
            </tr>
            <tr > 
              <td >所有日访问统计：</td>
              <td ><input name="daytj" type="checkbox" id="daytj" value="-1">
                全部删除</td>
            </tr>
            <tr > 
              <td >所有小时统计：</td>
              <td ><input name="hourtj" type="checkbox" id="hourtj" value="-1">
                全部删除</td>
            </tr>
            <tr > 
              <td >回头率统计：</td>
              <td ><input name="backtj" type="checkbox" id="backtj" value="-1">
                全部删除</td>
            </tr>
            <tr > 
              <td >Alexa统计：</td>
              <td ><input name="alexatj" type="checkbox" id="alexatj" value="-1">
                全部删除</td>
            </tr>
            <tr > 
              <td >搜索引擎数量统计：</td>
              <td > 
                <input name="searchtj" type="checkbox" id="searchtj" value="-1">
                全部删除</td>
            </tr>
            <tr > 
              <td >搜索引擎关键字统计：</td>
              <td >  
                <input name="keywordtj" type="checkbox" id="keywordtj" value="-1">
                全部删除</td>
            </tr>
			<tr > 
              <td >页面统计：</td>
              <td ><input name="webtj" type="checkbox" id="webtj" value="-1">
                全部删除</td>
            </tr>
            <tr > 
              <td >实际显示数：</td>
              <td > <input name="RealShowTotal" type="checkbox" id="hidden" value="-1">
                设置为0</td>
            </tr>
            <tr > 
              <td >实际IP数值：</td>
              <td ><input name="RealIpTotal" type="checkbox" id="RealIpTotal" value="-1">
                设置为0</td>
            </tr>
            <tr >
              <td >输入你的密码确认：</td>
              <td ><input name="Pwd" type="password" id="Pwd"></td>
            </tr>
            <tr > 
              <td >&nbsp;</td>
              <td > 
                <input type="submit" name="Submit2" value="重　置" onClick="{if(confirm('确定要重置计数么，这样会删除你的一些选择的数据!')){return true;}return false;};">
                　 
                <input type="reset" name="Submit22" value="取　消" >
              </td>
            </tr>
          </form>
        </table>
<%End If%>

<%If Action="userdel" Then%>
<div class="filters">
			 <h3>账号删除</h3>
       </div>
        <table width="100%" >

<form name="form3" method="post" action="?Action=userdelsave">
          <tr >
              <td colspan="2">你确定是否真的要删除！</td>
    </tr>
            <tr > 
              <td width="18%">输入你的密码确认：</td>
              <td ><input name="Pwd" type="password" id="Pwd"></td>
            </tr>
            <tr > 
              <td >&nbsp;</td>
              <td > 
                <input type="submit" name="Submit2" value="删　除" onClick="{if(confirm('确定要删除自己的账号么，你的账号在计数器中不复存在!')){return true;}return false;};">
                　 
                <input type="reset" name="Submit22" value="取　消" >
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
			 <h3>修改密码(带*为必填项）</h3>
       </div>
        <table width="100%" >

  <form name="form3" method="post" action="?Action=pwdmodifysave">
    <tr > 
      <td width="20%">请输入新管理密码：</td>
      <td > <input name="Pwd" type="password" id="Pwd_Old" size="15">
        (不想管理修改请留空)</td>
    </tr>
    <tr > 
      <td >确认新管理密码：</td>
      <td > <input name="Pwd2" type="password" id="Pwd" size="15">
      </td>
    </tr>
    <tr > 
      <td >请输入新查看密码：</td>
      <td > <input name="Pwd_View" type="password" id="Pwd_Old" size="15">
        (不想查看修改请留空) </td>
    </tr>
    <tr > 
      <td >确认新查看密码：</td>
      <td > <input name="Pwd_View2" type="password" id="Pwd_Old" size="15">
      </td>
    </tr>
    <tr > 
      <td >请输入旧管理密码确认：</td>
      <td >
        <input name="Pwd_Old" type="password" id="Pwd2" size="15">
        <font color="#0000FF">*</font></td>
    </tr>
    <tr > 
      <td >&nbsp;</td>
      <td >
        <input type="submit" name="Submit2" value="修　改" >
        　 
        <input type="reset" name="Submit22" value="取　消" >
      </td>
    </tr>
    <tr > 
      <td colspan="2" >注意：初始查看密码和刚注册时的管理密码一样，可设置成和管理密码不一样</td>
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
window.alert ('email必须填写');
document.form2.email.focus();
return false;
}
else if (!isValidEmail(document.form2.email.value)){
 window.alert ('email的格式不正确!');
 document.form2.email.focus();
 return false;
}
else if ((document.form2.pagename.value)=="")
{
window.alert ('主页名称必须填写');
document.form2.pagename.focus();
return false;
}
else if ((document.form2.pageurl.value)=="")
{
window.alert ('主页网址必须填写');
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
			 <h3>修改注册资料(带*为必填项）</h3>
       </div>
        <table width="100%" >

  <form name="form2" method="post" action="?Action=datamodifysave" onsubmit="return modifyuser()">
    <tr > 
      <td width="18%">E- mail：</td>
      <td > <input name="email" type="text"  value="<%=rs("email")%>" size="30"> 
        <font color="#0000FF">*</font></td>
    </tr>
    <tr > 
      <td >QQ 号码：</td>
      <td > <input name="qq" type="text" id="qq"  value="<%=rs("qq")%>" size="30"> 
      </td>
    </tr>
    <tr > 
      <td >项目名称：</td>
      <td > <input name="pagename" type="text"  value="<%=rs("pagename")%>" size="30"> 
        <font color="#0000FF">*</font></td>
    </tr>
    <tr > 
      <td >项目网址：</td>
      <td > <input name="pageurl" type="text"  value="<%=rs("pageurl")%>" size="30"> 
        <font color="#0000FF">*</font></td>
    </tr>
    <tr >
      <td >请输入密码确认修改：</td>
      <td ><input name="Pwd" type="password" size="30"><font color="#0000FF">*</font></td>
    </tr>
    <tr > 
      <td >&nbsp;</td>
      <td > <input type="submit" name="Submit2" value="修　改" >
        　 
        <input type="reset" name="Submit22" value="取　消" > </td>
    </tr>
  </form>
</table>
<%End If%>
        <%If Action="passwordanswermodify" Then%>
        <%Set Rs= Server.CreateObject("Adodb.RecordSet")
Sql="Select * From CFCount_User Where UserName='"&UserName&"'"
Rs.Open Sql,Conn,1,1%>
<div class="filters">
			 <h3>修改密码提示和答案(带*为必填项）</h3>
       </div>
        <table width="100%" >


          <form name="form3" method="post" action="?Action=passwordanswermodifysave">
            <%If Rs("passwordask")<>"" Then%>
            <tr > 
              <td width="16%">原密码提示问题：</td>
              <td ><%=Rs("passwordask")%></td>
            </tr>
            <tr > 
              <td >原密码回答答案：</td>
              <td ><input name="passwordanswer_old" type="text" id="passwordanswer" size="20">
                <font color="#0000FF">*</font>(请输入原答案确认要修改为以下新提示和答案)</td>
            </tr>
            <%End If%>
            <tr > 
              <td width="16%">新密码提示问题：</td>
              <td > 
                <input name="passwordask_new" type="text" id="passwordask" size="20">
                <font color="#0000FF">*</font> </td>
            </tr>
            <tr > 
              <td >新密码回答答案：</td>
              <td > 
                <input name="passwordanswer_new" type="text" id="passwordanswer" size="20">
                <font color="#0000FF">*</font></td>
            </tr>
            <tr > 
              <td >&nbsp;</td>
              <td > 
                <input type="submit" name="Submit23" value="修　改" >
                　 
                <input type="reset" name="Submit222" value="取　消" >
              </td>
            </tr>
          </form>
        </table>
        <%End if%>
        <%If Action="stylelist" Then%>
<%
 If Session("CFCountUser_View")<>"" Then'为独立查看者时
  Call AlertBack("独立查看者无法对此操作",1)
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
			 <h3>样式代码-选择你喜欢的统计样式</h3>
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
              <td width="10%">确认使用：</td>
              <td ><div >
                  <input type="checkbox" name="CounterShow" value="-1"<%If Rs("CounterShow")=-1 Then Response.Write " checked"%>>是
              </td>
            </tr>
            <tr> 
              <td >&nbsp; </td>
              <td > <div > 
                  <input type="submit" name="Submit3" value="确定">
              </td>
            </tr>
          </form>
          <tr> 
            <td colspan="2" ><table class="tb_3">
                <tr> 
                  <td><a href="?action=<%=action%>&page=1">第一页</a> 
                      <%
if page>1 then%>
                      <a href='?action=<%=action%>&page=<%=page-1%>'>上一页</a> 
                      <%
end if
%>
                      <%
if page<totalpage   then%>
                      <a href='?action=<%=action%>&page=<%=page+1%>'>下一页</a> 
      <%
end if
%>
                      <a href="?action=<%=action%>&page=<%=totalpage%>">最后一页</a> 
                      页次：<font color="#ff0000"><%=page%></font>/<%=totalpage%>页
                            
<%
Response.write "&nbsp;&nbsp;转到第<select id='page' onChange=""window.location=document.getElementById('page').options[document.getElementById('page').selectedIndex].value"">"
For I=1 To TotalPage
    Response.Write "<option value=?action="&action&"&page="&i     
 If Page=I Then Response.Write " selected"
 Response.Write ">"& I &"</option>"
Next
Response.write "</select>页"
%></td>
              </tr>
                    </table></td>
          </tr>
</table>
          <table >
           <tr >
            <td colspan="2">你的统计效果 
            </td>
          </tr>
          <tr > 
            <td colspan="2"  >Script调用： 
            </td>
          </tr>
          <tr > 
            <td colspan="2"  ><script src="<%=Tmp%>cf.asp?UserName=<%=Rs("UserName")%>"></script></td>
          </tr>
        </table>
<%end if%>



<%If Action="gbooklist" Then%>
<%
 If Session("CFCountUser_View")<>"" Then'为独立查看者时
  Call AlertBack("独立查看者无法对此操作",1)
 End If
%>
<div class="filters">
			 <h3>留言列表</h3>
       </div>
        <table width="100%" >
          <tr > 
            <td class="td_nowrap">ID</td>
            <td class="td_nowrap">留言内容</td>
            <td class="td_nowrap">联系方式</td>
            <td class="td_nowrap">留言时间</td>
			<td class="td_nowrap">留言来源</td>
			<td class="td_nowrap">留言当前页</td>
            <td class="td_nowrap">操作</td>
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
            <td class="td_nowrap"><a href="?Action=gbookmodify&ID=<%=Rs("ID")%>">修改</a> <a href="?Action=gbookdel&ID=<%=Rs("ID")%>" onClick="{if(confirm('确定要删除么?')){return true;}return false;}">删除</a></td>
          </tr>
          <%rs.movenext
mypagesize=mypagesize-1
wend%>
        </table>
		
<table class="tb_3">
  <tr>
    <td><a href="?Action=<%=Action%>&Page=1">第一页</a>
            <%
if page>1 then%>
            <a href='?Action=<%=Action%>&Page=<%=Page-1%>'>上一页</a>
            <%
end if
%>
            <%
if page<rs.pagecount   then%>
            <a href='?Action=<%=Action%>&Page=<%=Page+1%>'>下一页</a>
            <%

end if
%>
            <a href="?Action=<%=Action%>&Page=<%=totalpage%>">最后一页</a> 页次：<font color="#ff0000"><%=page%></font>/<%=totalpage%>页&nbsp;&nbsp;共<%=totalrs%>条记录&nbsp;&nbsp;每页显示<%=rs.pagesize%>条
<%
Response.write "&nbsp;&nbsp;转到第<select id='page' onChange=""window.location=document.getElementById('page').options[document.getElementById('page').selectedIndex].value"">"
For I=1 To TotalPage
    Response.Write "<option value=?Action="&Action&"&Page="&i
 If Page=I Then Response.Write " selected"
 Response.Write ">"& I &"</option>"
Next
Response.write "</select>页"
%>
      </td>

  </tr>
</table>
		  
<%End If%>
		  

<%If Action="gbookmodify" Then%>
<%
 If Session("CFCountUser_View")<>"" Then'为独立查看者时
  Call AlertBack("独立查看者无法对此操作",1)
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
      <td colspan="2">留言修改</td>
    </tr>
    <tr> 
      <td class="td_wauto">留言内容：</td>
      <td><textarea name='content' id='content' cols='40' rows='8'><%=server.HTMLEncode(Rs("content"))%></textarea></td>
    </tr>
    <tr> 
      <td>联系方式：</td>
      <td> <textarea name='contact' id='contact' cols='40' rows='8'><%=server.HTMLEncode(Rs("contact"))%></textarea> </td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><input type="submit" name="Submit" value="修改">      </td>
    </tr>
  </form>
</table>
        
<%End If%>
			  

<%If Action="getcode" Then%>

<script language=javascript>
function copy(ob){
var obj=findObj(ob); if (obj) { 
obj.select();js=obj.createTextRange();js.execCommand("Copy");}
alert('成功将代码复制到剪贴板中！');
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
			 <h3>获得统计代码及统计预览</h3>
       </div>
        <table >
  <tr > 
    <td  >两种代码方式可供选择：<input type="submit" name="Submit5" value="邮箱类方式调用" onclick="location.href='?Action=getcode'"<%If Action="getcode" Then%> style="BACKGROUND-COLOR: #e8e8e8;"<%End If%>> 
      <input type="submit" name="Submit6" value="网站方式调用" onclick="location.href='?Action=getcode_2'"<%If Action="getcode_2" Then%> style="BACKGROUND-COLOR: #e8e8e8;"<%End If%>></td>
  </tr>
  <tr > 
    <td  >Img方式调用，此方式适合放在所发送邮件的内容当中，如果你使用的我们的邮件群发系统，则不需要此引用，我们将自动进行统计。</td>
  </tr>
  <tr > 
    <td  >以下方框内是你的统计代码，Img调用方式，请复制后插入到邮件内容HTML源代码里,这种方式调用你可以把你的统计看做是一个图片，图片地址是<a href="<%=Tmp%>cf.asp?c=<%=UserName%>" target="_blank"><%=Tmp%>cf.asp?c=<%=Rs("UserName")%></a>你设置隐藏时显示数为0！<br> <textarea name="countercode2" cols="50" rows="6" id="textarea"><a href="<%=Tmp%>Index.asp" target="_blank"><img src="<%=Tmp%>cf.asp?c=<%=UserName%>" border="0" title="<%=server.HTMLEncode(GetUnicode(RsSet("SysTitle")))%>"></a></textarea> 
      <input type="submit" name="Submit4" value="复制" onClick=copy('countercode2')> 
      <br> <a href="<%=Tmp%>Index.asp" target="_blank"><img src="<%=Tmp%>cf.asp?c=<%=UserName%>" border="0" title="<%=RsSet("SysTitle")%>"></a> 
  </tr>
</table>
<%end if%>

<%If Action="getcode_2" Then%>

<script language=javascript>
function copy(ob){
var obj=findObj(ob); if (obj) { 
obj.select();js=obj.createTextRange();js.execCommand("Copy");}
alert('成功将代码复制到剪贴板中！');
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
			 <h3>获得统计代码及统计预览</h3>
       </div>
        <table >
  <tr > 
    <td  >两种代码方式可供选择：<input type="submit" name="Submit5" value="邮箱类方式调用" onclick="location.href='?Action=getcode'"<%If Action="getcode" Then%> style="BACKGROUND-COLOR: #e8e8e8;"<%End If%>> 
      <input type="submit" name="Submit6" value="网站方式调用" onclick="location.href='?Action=getcode_2'"<%If Action="getcode_2" Then%> style="BACKGROUND-COLOR: #e8e8e8;"<%End If%>></td>
  </tr>
  <tr > 
    <td  >Script网站方式调用，此方式适合放在自己的网站页面上，有几十种图片样式可以选择，请在菜单"<a href="?Action=stylelist" target="_blank">统计设置>>图片样式选择</a>"里选择!设置统计的一些参数请在请在菜单"<a href="?Action=jsqset" target="_blank">统计设置>>参数设置</a>"里设置，打造你满意统计!</td>
  </tr>
  <tr > 
    <td  >以下方框内是你的统计代码,script调用方式，请复制后插入到网页源代码里:<br> 
      <textarea name="countercode1" cols="50" rows="4"><script src="<%=Tmp%>cf.asp?username=<%=UserName%>"></script></textarea> 
      <input type="submit" name="Submit4" value="复制" onClick=copy('countercode1')> 
      <br>
      你的设置的统计script方式调用预览： <br> <script src="<%=Tmp%>cf.asp?username=<%=UserName%>"></script> 
    </td>
  </tr>
</table>  
<%end if%>