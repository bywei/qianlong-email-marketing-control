
<%If Action="" Or Action="tjsurvey" Then%>
<%
Set RsUser= Server.CreateObject("Adodb.RecordSet")
Sql="Select * From CFCount_User Where UserName='"&UserName&"'"
RsUser.Open Sql,Conn,1,1


Sql="Select Top 1 MyCounter,AddDate From CFCount_Count_Day Where UserName='"&UserName&"' Order By MyCounter Desc"
Set Rs=Conn.Execute(Sql)
If Not Rs.Eof Then
 TopShow=Rs("MyCounter")
 TopShowDate=Rs("AddDate")
Else
 TopShow=0
 TopShowDate=""
End If
Rs.Close

Sql="Select Top 1 IpCounter,AddDate From CFCount_Count_Day Where UserName='"&UserName&"' Order By IpCounter Desc"
Set Rs=Conn.Execute(Sql)
If Not Rs.Eof Then
 TopIp=Rs("IpCounter")
 TopIpDate=Rs("AddDate")
Else
 TopIp=0
 TopIpDate=""
End If
Rs.Close

Sql="Select Top 1 MyCounter,AddDate From CFCount_Count_Day Where UserName='"&UserName&"' And DateDiff('d','"&RsUser("AddTime")&"',AddDate)>0 And DateDiff('d',AddDate,Date())>0 Order By MyCounter"
Set Rs=Conn.Execute(Sql)
If Not Rs.Eof Then
 LowShow=Rs("MyCounter")
 LowShowDate=Rs("AddDate")
Else
 LowShow=0
 LowShowDate=""
End If
Rs.Close

Sql="Select Top 1 IpCounter,AddDate From CFCount_Count_Day Where UserName='"&UserName&"' And DateDiff('d','"&RsUser("AddTime")&"',AddDate)>0 And DateDiff('d',AddDate,Date())>0 Order By IpCounter"
Set Rs=Conn.Execute(Sql)
If Not Rs.Eof Then
 LowIp=Rs("IpCounter")
 LowIpDate=Rs("AddDate")
Else
 LowIp=0
 LowIpDate=""
End If
Rs.Close
%>
<div class="filters">
			 <h3>统计概况</h3>
<%
dim userSendNum
Set Rs= Server.CreateObject("Adodb.RecordSet")
Sql="Select * From Users Where UserName='"&UserName&"'"
Rs.Open Sql,Conn,1,1
userSendNum=Rs("userSendNum")

Set Rs= Server.CreateObject("Adodb.RecordSet")
Sql="Select * From CFCount_User Where UserName='"&UserName&"'"
Rs.Open Sql,Conn,1,1
If Rs("UserState")=0 Then Response.Write "&nbsp;&nbsp;<font color=#ff0000>你的账号已被禁用!</font>"

If Rs("ParentName")<>"" Then Response.Write "&nbsp;&nbsp;你的父账号是:"&Rs("ParentName")
%>
</div>
<div class="filters">
           常用设置链接：[<a href="?Action=jsqquickset">计数器基本样式快速设置</a>]&nbsp;[<a href="?Action=jsqset">计数器参数设置</a>]&nbsp;[<a href="?Action=stylelist">计数器图片样式选择</a>]&nbsp;[<a href="?Action=getcode">预览及获得计数器代码</a>]&nbsp;[<a href="?Action=modifypassword">修改密码</a>]&nbsp;
</div>
        <table width="100%">
          <tr > 
            <td width="12%">项目名称：</td>
            <td colspan="3" ><b><%=RsUser("PageName")%></b></td>
          </tr>
          <tr > 
            <td >用户邮箱：</td>
            <td colspan="2" ><%=RsUser("Email")%>
            <div style="display:none">
            <a href="<%=RsUser("PageUrl")%>" target="_blank"><%=RsUser("PageUrl")%></a>&nbsp;[<a href="http://www.alexa.com/data/details/traffic_details?q=&url=<%=RsUser("PageUrl")%>" target="_blank">Alexa排名查询</a>]&nbsp;[<a href="http://web.archive.org/web/*/<%=RsUser("PageUrl")%>" target="_blank">历史网页回忆</a>]
            </div>
            </td>
            <td >发件总数：<%=userSendNum%>封</td>
          </tr>
          <tr > 
            <td >开始统计：</td>
            <td colspan="2" ><%=RsUser("AddTime")%></td>
            <td title="表示发送到邮件被阅读的次数,一封邮件可以被多次查看，注意区别于‘访问总数’">阅读总数：<%=RsUser("RealShowTotal")%> 次</td>
          </tr>
          <tr > 
            <td >已计天数：</td>
            <td colspan="2" ><%=DateDiff("d",RsUser("AddTime"),Now())%>天</td>
            <td title="表示发送到邮件已经查看的人数">访问总数：<%=RsUser("RealIpTotal")%> 次</td>
          </tr>
          <tr > 
            <td >正在阅读：</td>
            <td colspan="2" >
<%
If IsEmpty(Application("CFCountOnline_"&UserName)) Then
 OnlineTotal=0
Else
 Myarray=Split(Application("CFCountOnline_"&UserName),"|")
 OnlineTotal=Ubound(Myarray)
End If
Response.Write OnlineTotal&"人"%></td>
            <td title="表示邮件还未打开,进垃圾箱，未及时统计的，和还没有查看的数量">沉睡总数：<%=cint(userSendNum)-cint(RsUser("RealIpTotal"))%> 封</td>
          </tr>
          <tr > 
            <td ></td>
            <td width="62" >&nbsp;</td>
            <td width="84" >&nbsp;</td>
            <td width="407" ></td>
          </tr>
          <tr > 
            <td >今日流量：</td>
            <%
Sql="Select MyCounter,IpCounter From CFCount_Count_Day Where UserName='"&UserName&"' And DateDiff('d',AddDate,Date())=0"
Set Rs=Conn.Execute(Sql)
TodayShow=Rs("MyCounter")
TodayIp=Rs("IpCounter")
%>
            <td ><%=Rs("MyCounter")%></td>
            <td ><%=Rs("IpCounter")%></td>
            <td >最高访问量： <%=TopIp%> 次 (发生在: <%=TopIpDate%>) </td>
          </tr>
          <%
Sql="Select MyCounter,IpCounter From CFCount_Count_Day Where UserName='"&UserName&"' And DateDiff('d',AddDate,Date())=1"
Set Rs=Conn.Execute(Sql)
If Not Rs.Eof Then
 MyCounter=Rs("MyCounter")
 IpCounter=Rs("IpCounter")
Else
 MyCounter=0
 IpCounter=0
End If
%>
          <tr > 
            <td >昨日流量：</td>
            <td ><%=MyCounter%></td>
            <td ><%=IpCounter%></td>
            <td >最高阅读量： <%=TopShow%> 次 (发生在: <%=TopShowDate%>)</td>
          </tr>
          <%
Sql="Select Sum(MyCounter),Sum(IpCounter) From CFCount_Count_Day Where UserName='"&UserName&"' And DateDiff('m',AddDate,Date())=0"
Set Rs=Conn.Execute(Sql)
%>
          <tr > 
            <td >本月累计：</td>
            <td ><%=Rs(0)%></td>
            <td ><%=Rs(1)%></td>
            <td >&nbsp;</td>
          </tr>
          <%
Sql="Select Sum(MyCounter),Sum(IpCounter) From CFCount_Count_Day Where UserName='"&UserName&"' And DateDiff('yyyy',AddDate,Date())=0"
Set Rs=Conn.Execute(Sql)
%>
          <tr > 
            <td >今年累计：</td>
            <td ><%=Rs(0)%></td>
            <td ><%=Rs(1)%></td>
            <td >最低访问量： <%=LowIp%> 次 (发生在: <%=LowIpDate%>)</td>
          </tr>
          <tr > 
            <td >平均每日：</td>
            <td ><%=Int(RsUser("RealShowTotal")/DateDiff("d",RsUser("AddTime"),Now()+1))%></td>
            <td ><%=Int(RsUser("RealIpTotal")/DateDiff("d",RsUser("AddTime"),Now()+1))%></td>
            <td >最低阅读量： <%=LowShow%> 次 (发生在: <%=LowShowDate%>)</td>
          </tr>
          <tr > 
            <td >预计今日：</td>
            <td > 
                <%
			MyTime=DateDiff("n",Date(),Now())
			If MyTime=0 Then MyTime=1
			Response.write Int(TodayShow*1440/MyTime)
			%>
            </td>
            <td > 
                <%Response.write Int(TodayIp*1440/MyTime)%>
            </td>
            <td >&nbsp;</td>
          </tr>
        </table>
		
		
<%If RsUser("ParentName")="" And Session("CFCountUser_View")="" Then'为父账号且是管理员时%>		
        <table>
          <tr >
    <td colspan="4"> 子账号</td>
  </tr>
  <tr > 
    <td colspan="4" ><input type="submit" name="Submit" value="添加子账号" onclick="location.href='?Action=subadd'">
      如果你有多个项目需要分别进行统计，可以增加子账号，无需注册多个用户名</td>
  </tr>
  <tr > 
    <td colspan="4" >已开通的子账号：</td>
  </tr>
  <%
Sql="Select * From CFCount_User Where ParentName='"&UserName&"'"
Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.open Sql,Conn,1,1
%>
  <%
If Rs.Eof Then
%>
  <tr > 
    <td colspan="4" >暂时没有子账号</td>
  </tr>
  <%
Else
%>
  <tr > 
    <td >用户名</td>
    <td >站点名称</td>
    <td >域名</td>
    <td ></td>
  </tr>
  <%
End If%>
  <%
While Not Rs.Eof
%>
  <tr > 
    <td ><%=Rs("UserName")%></td>
    <td ><%=Rs("PageName")%></td>
    <td ><a href="<%=Rs("PageUrl")%>" target="_blank"><%=Rs("PageUrl")%></a></td>
    <td ><input type="submit" name="Submit" value="进入管理" onclick="location.href='?Action=subgoto&SubUserName=<%=Rs("UserName")%>'">&nbsp;&nbsp;&nbsp;<a href="?Action=submodify&SubUserName=<%=Rs("UserName")%>">修改</a>&nbsp;&nbsp;&nbsp;<a href="?Action=subdel&SubUserName=<%=Rs("UserName")%>" onClick="{if(confirm('确定要删除这个子账号么，删除后此子账号时统计数据将不存在!')){return true;}return false;};">删除账号</a></td>
  </tr>
  <%
 Rs.MoveNext
 Wend
%>
</table>
<%End If%>

<%End If%>


<%If Action="subadd" Then%>
        
        <table>


  <form name="form3" method="post" action="?Action=subaddsave">
          <tr >
      <td colspan="2"> 添加子账号</td>
    </tr>
    <tr> 
      <td>账号名称：</td>
      <td><input name="SubUserName" type="text" id="password"></td>
    </tr>
    <tr> 
      <td>账号密码：</td>
      <td><input name="Pwd" type="password" id="password"></td>
    </tr>
    <tr>
      <td>重复账号密码：</td>
      <td><input name="Pwd2" type="password" id="password"></td>
    </tr>
    <tr> 
      <td>站点名称：</td>
      <td><input name="PageName" type="text" id="password"></td>
    </tr>
    <tr> 
      <td>域名：</td>
      <td><input name="PageUrl" type="text" id="PageName"></td>
    </tr>
    <tr > 
      <td>&nbsp;</td>
      <td> <input type="submit" name="Submit2" value="确定"> </td>
    </tr>
  </form>
</table>
<%End If%>

<%If Action="submodify" Then%>
<%
SubUserName=ChkStr(Request("SubUserName"),1)
Set Rs= Server.CreateObject("Adodb.RecordSet")
Sql="Select * From CFCount_User Where UserName='"&SubUserName&"' And ParentName='"&UserName&"'"
Rs.Open Sql,Conn,1,1%>
        
        <table>


  <form name="form3" method="post" action="?Action=submodifysave">
          <tr >
      <td colspan="2"> 修改子账号</td>
    </tr>
    <tr>
      <td>账号名称：</td>
      <td><input name="SubUserName" type="hidden" value="<%=Rs("UserName")%>">
        <%=Rs("UserName")%></td>
    </tr>
    <tr> 
      <td>站点名称：</td>
      <td><input name="PageName" type="text" id="password" value="<%=Rs("PageName")%>"></td>
    </tr>
    <tr> 
      <td>域名：</td>
      <td><input name="PageUrl" type="text" id="PageName" value="<%=Rs("PageUrl")%>"></td>
    </tr>
    <tr > 
      <td>&nbsp;</td>
      <td> <input type="submit" name="Submit2" value="确定"> 
      </td>
    </tr>
  </form>
</table>
<%End If%>

<%If Action="lastvisit" Then%>
<%
Myarray=Split(Application("CFCountLy_"&UserName),"|")
%>

         <div class="filters">
			 <h3>最新访问记录</h3>
		 </div>
        <table>
          <tr > 
            <td class="td_nowrap">序号</td>
            <td class="td_nowrap">IP</td>
            <td class="td_nowrap">访问次数</td>
            <td class="td_nowrap">访问时间</td>
            <td class="td_nowrap">来源页面</td>
            <td class="td_nowrap">最后访问时间</td>
            <td class="td_nowrap">最后浏览页面</td>
            <td class="td_nowrap">共浏览几页</td>
          </tr>
          <%For I=Ubound(Myarray)-1 To 0 Step -1%>
          <%MyArray_2=Split(MyArray(I),"\")
		  J=J+1%>
          <tr > 
            <td class="td_nowrap"><%=J%></td>
            <td class="td_nowrap"> <%If Ip=MyArray_2(0) Then Response.Write "<font color=ff0000>"%> <%=MyArray_2(0)%> <%If Ip=MyArray_2(0) Then Response.Write "</font>"%>
			<%If RsSet("IpArea")=-1 And J<=20 Then
				  Response.write "<br>"&GetIpArea(MyArray_2(0))
				 End If%></td>
            <td  class="td_nowrap"><%=MyArray_2(1)%></td>
            <td ><%=GetTurnTime(Hour(Cdate(MyArray_2(2))))&":"&GetTurnTime(Minute(Cdate(MyArray_2(2))))&":"&GetTurnTime(Second(Cdate(MyArray_2(2))))%></td>
            <td > <%If MyArray_2(3)="" Then%>
              直接从浏览器输入 
              <%Else%> <a href="?Action=urlgo&Url=<%=Server.URLEncode("http://"&BreakUrl(MyArray_2(3),1))%>" target='_blank'><%=BreakUrl(MyArray_2(3),1)%></a>[<a href="?Action=urlgo&Url=<%=Server.URLEncode(MyArray_2(3))%>" target='_blank'>详细</a>] 
              <%End If%></td>
            <td class="td_nowrap"><%=GetTurnTime(Hour(Cdate(MyArray_2(4))))&":"&GetTurnTime(Minute(Cdate(MyArray_2(4))))&":"&GetTurnTime(Second(Cdate(MyArray_2(4))))%></td>
            <td > <%If MyArray_2(5)="" Then%>
              直接从浏览器输入 
              <%Else%> <a href="?Action=urlgo&Url=<%=Server.URLEncode(MyArray_2(5))%>" target='_blank'><%=BreakUrl(MyArray_2(5),2)%></a>[<a href="?Action=urlgo&Url=<%=Server.URLEncode(MyArray_2(5))%>" target='_blank'>详细</a>] 
              <%End If%></td>
            <td class="td_nowrap"><%=MyArray_2(6)%></td>
          </tr>
          <%
Next%>
        </table>
<%End If%>
<%If Action="lylist" Then%>

<%If RsSet("LyKeep")=0 Then Response.write "系统已经禁用此功能"%>
<div class="filters">
			 <h3>来源明细</h3>
</div>
        <table width="100%">
          <tr>
            <form name="form6" method="post" action="?">
              <td colspan="7" >选择查询的日期 
                        <Select  name="adddate" size="1" onChange="window.location=form.adddate.options[form.adddate.selectedIndex].value">
                          <%Sql="Select AddDate From CFCount_Ly_Day Where UserName='"&UserName&"' Group By AddDate Order By AddDate Desc"
Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.open Sql,Conn,1,1
While Not Rs.Eof 
 Response.Write "<option value='"&filename&"?Action=lylist&AddDate="&Rs("AddDate")&"'"
 If Request("AddDate")=Cstr(Rs("AddDate")) Then Response.Write " selected"
 Response.Write ">"&Rs("AddDate")&"</option>"  
Rs.MoveNext
Wend  
%>
                        </select></td>

          </form>

          </tr>
          <%
Px=ChkStr(Request("Px"),1)
AddDate=ChkStr(Request("AddDate"),3)

If Px="" Then  Px="MyCounter"
If AddDate="" Then AddDate=Date()

Call PxFilter(Px,"id,ip,lyhead,addtime,MyCounter,IpCounter,lasttime")

Sql="Select * From CFCount_Ly_Day where UserName='"&UserName&"' And Datediff('d',AddDate,'"&AddDate&"')=0"

Sql=Sql&" Order By "&Px&" Desc"
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
          <tr > 
            <td colspan="7"><%=AddDate%> 来源统计[可点击标题排序]</td>
          </tr>
          <tr > 
            <td class="td_nowrap"><a href="?Action=<%=Action%>&AddDate=<%=AddDate%>&Px=ID">序号</a></td>
            <td class="td_nowrap"><a href="?Action=<%=Action%>&AddDate=<%=AddDate%>&Px=LyHead">来源路径</a></td>
            <td class="td_nowrap"><a href="?Action=<%=Action%>&AddDate=<%=AddDate%>&Px=MyCounter">阅读数量</a></td>
			<td class="td_nowrap"><a href="?Action=<%=Action%>&AddDate=<%=AddDate%>&Px=IpCounter">访问数量</a></td>
            <td class="td_nowrap"><a href="?Action=<%=Action%>&AddDate=<%=AddDate%>&Px=AddTime">上站时间</a></td>
            <td class="td_nowrap"><a href="?Action=<%=Action%>&AddDate=<%=AddDate%>&Px=Ip">最后IP地址</a></td>
            <td class="td_nowrap"><a href="?Action=<%=Action%>&AddDate=<%=AddDate%>&Px=LastTime">最后访问时间</a></td>
          </tr>
          <%
While Not Rs.Eof And MyPageSize>0
%>
          <tr > 
            <td class="td_nowrap"><%=Rs.RecordCount-Rs.Pagesize*(page)+mypagesize%></td>
            <td class="td_nowrap"> 
<%If Rs("LyHead")="" Then
  Response.Write "直接从浏览器输入"
 Else
  Response.Write "<a href=?Action=urlgo&Url="&Server.URLEncode("http://"&Rs("LyHead"))&" target='_blank'>"&Rs("LyHead")&"</a>&nbsp;[<a href=?Action=urlgo&Url="&Server.URLEncode(Rs("Ly"))&" target='_blank'>详细</a>]"
 End If%>
            </td>
            <td class="td_nowrap"><%=Rs("MyCounter")%></td>
			<td class="td_nowrap"><%=Rs("IpCounter")%></td>
            <td class="td_nowrap"><%=GetTurnTime(Hour(Rs("AddTime")))&":"&GetTurnTime(Minute(Rs("AddTime")))&":"&GetTurnTime(Second(Rs("AddTime")))%></td>
            <td class="td_nowrap"> 
                <%If Ip=Rs("Ip") Then Response.Write "<font color=ff0000>"%>
                <%=Rs("Ip")%> 
                <%If Ip=Rs("Ip") Then Response.Write "</font>"%>
				<%If RsSet("IpArea")=-1 Then
				  Response.write "<br>"&GetIpArea(Rs("Ip"))
				 End If%>
            </td>
            <td class="td_nowrap"><%=GetTurnTime(Hour(Rs("LastTime")))&":"&GetTurnTime(Minute(Rs("LastTime")))&":"&GetTurnTime(Second(Rs("LastTime")))%></td>
          </tr>
          <%rs.movenext
mypagesize=mypagesize-1
wend%>


        </table>
<table class="tb_3">
  <tr>
    <td><a href="?Action=<%=Action%>&AddDate=<%=AddDate%>&Px=<%=Px%>&Page=1">第一页</a>
            <%
if page>1 then%>
            <a href='?Action=<%=Action%>&AddDate=<%=AddDate%>&Px=<%=Px%>&Page=<%=Page-1%>'>上一页</a>
            <%
end if
%>
            <%
if page<rs.pagecount   then%>
            <a href='?Action=<%=Action%>&AddDate=<%=AddDate%>&Px=<%=Px%>&Page=<%=Page+1%>'>下一页</a>
            <%

end if
%>
            <a href="?Action=<%=Action%>&AddDate=<%=AddDate%>&Px=<%=Px%>&Page=<%=totalpage%>">最后一页</a> 页次：<font color="#ff0000"><%=page%></font>/<%=totalpage%>页&nbsp;&nbsp;共<%=totalrs%>条记录&nbsp;&nbsp;每页显示<%=rs.pagesize%>条
<%
Response.write "&nbsp;&nbsp;转到第<select id='page' onChange=""window.location=document.getElementById('page').options[document.getElementById('page').selectedIndex].value"">"
For I=1 To TotalPage
    Response.Write "<option value=?Action="&Action&"&AddDate="&AddDate&"&Px="&Px&"&Page="&i
 If Page=I Then Response.Write " selected"
 Response.Write ">"& I &"</option>"
Next
Response.write "</select>页"
%>
    </td>

  </tr>
</table>
<%End If%>


<%If Action="onlinetj" Then%>
<%
Set Rs= Server.CreateObject("Adodb.RecordSet")
Sql="Select * From CFCount_User Where UserName='"&UserName&"'"
Rs.Open Sql,Conn,1,1
%>
<%If Rs("Online")=0 Then Response.write "你的设置已经禁用此功能<br>"%>
<%If RsSet("OnlineKeep")=0 Then Response.write "系统已经禁用此功能"%>
<%

Myarray=Split(Application("CFCountOnline_"&UserName),"|")
TotalRs=Ubound(Myarray)

If TotalRs>1 Then
 MyArray_2=Split(Myarray(TotalRs-1),"\")
 Time_1=MyArray_2(2)
 MyArray_2=Split(Myarray(0),"\")
 Time_2=MyArray_2(2)
 TotalTime=DateDiff("n",Cdate(Time_2),Cdate(Time_1))
End If

IF not IsNumeric(Request("Page")) Or IsEmpty(Request("Page")) Then
 Page=1
Else
 Page=Int(Abs(Request("Page")))
End If

MyPageSize=10
MyPageSize2=10
MyPageSize3=MyPageSize

If TotalRs Mod MyPageSize=0 Then
 TotalPage=TotalRs/MyPageSize
Else
 TotalPage=Int(TotalRs/MyPageSize)+1
 MyPageSize2=TotalRs Mod MyPageSize
 If Page=TotalPage Then
  MyPageSize3=MyPageSize2
 End If
End If
I=0
%>
<div class="filters">
			 <h3>正在阅读统计<%If TotalRs>1 Then Response.write "&nbsp;["&TotalTime&"分钟内有"&TotalRs&"个人访问]"%></h3>
</div>
<table width="100%">
          <tr > 
            <td class="td_nowrap">序号</td>
            <td class="td_nowrap">IP</td>
            <td class="td_nowrap">访问次数</td>
            <td class="td_nowrap">访问时间</td>
            <td class="td_nowrap">来源页面</td>
            <td class="td_nowrap">最后访问时间</td>
            <td class="td_nowrap">最后浏览页面</td>
            <td class="td_nowrap">共浏览几页</td>
          </tr>
<%While I<MyPageSize3%>
<%J=(TotalPage-Page)*MyPageSize+MyPageSize2-I%>
<%
If J>0 Then
MyArray_2=Split(MyArray(J-1),"\")
If Ubound(MyArray_2)=6 Then
%>
          <tr > 
            <td class="td_nowrap"><%=J%></td>
            <td class="td_nowrap"><%If Ip=MyArray_2(0) Then Response.Write "<font color=ff0000>"%> <%=MyArray_2(0)%> <%If Ip=MyArray_2(0) Then Response.Write "</font>"%>
							<%If RsSet("IpArea")=-1 Then
				  Response.write "<br>"&GetIpArea(MyArray_2(0))
				 End If%></td>
            <td class="td_nowrap"><%=MyArray_2(1)%></td>
            <td><%=GetTurnTime(Hour(Cdate(MyArray_2(2))))&":"&GetTurnTime(Minute(Cdate(MyArray_2(2))))&":"&GetTurnTime(Second(Cdate(MyArray_2(2))))%></td>
            <td> <%If MyArray_2(3)="" Then%>
              直接从浏览器输入 
              <%Else%> <a href="?Action=urlgo&Url=<%=Server.URLEncode("http://"&BreakUrl(MyArray_2(3),1))%>" target='_blank'><%=BreakUrl(MyArray_2(3),1)%></a>[<a href="?Action=urlgo&Url=<%=Server.URLEncode(MyArray_2(3))%>" target='_blank'>详细</a>] 
              <%End If%></td>
            <td class="td_nowrap"><%=GetTurnTime(Hour(Cdate(MyArray_2(4))))&":"&GetTurnTime(Minute(Cdate(MyArray_2(4))))&":"&GetTurnTime(Second(Cdate(MyArray_2(4))))%></td>
            <td><%If MyArray_2(5)="" Then%>
              直接从浏览器输入 
              <%Else%><a href="?Action=urlgo&Url=<%=Server.URLEncode(MyArray_2(5))%>" target='_blank'><%=BreakUrl(MyArray_2(5),2)%></a>[<a href="?Action=urlgo&Url=<%=Server.URLEncode(MyArray_2(5))%>" target='_blank'>详细</a>] 
              <%End If%></td>
            <td class="td_nowrap"><%=MyArray_2(6)%></td>
          </tr>
<%
End If
End If

I=I+1
Wend%>
            <tr> 
              <td colspan="9" >(注：红色IP为你自己的IP) </td>
            </tr>
</table>
		<table class="tb_3">
    <tr> 
      <td height="21"><a href="?Action=<%=Action%>&Page=1">第一页</a> 
          <%
if page>1 then%>
          <a href='?Action=<%=Action%>&Page=<%=Page-1%>'>上一页</a> 
          <%
end if
%>
          <%
if page<TotalPage   then%>
          <a href='?Action=<%=Action%>&Page=<%=Page+1%>'>下一页</a> 
          <%
end if
%>
          <a href="?Action=<%=Action%>&Page=<%=TotalPage%>">最后一页</a> 
          页次：<font color="#ff0000"><%=Page%></font>/<%=TotalPage%>页&nbsp;&nbsp;共<%=TotalRs%>条记录&nbsp;&nbsp;每页显示<%=Rs.PageSize%>条
%>
<%
Response.write "&nbsp;&nbsp;转到第<select id='page' onChange=""window.location=document.getElementById('page').options[document.getElementById('page').selectedIndex].value"">"
For I=1 To TotalPage
    Response.Write "<option value=?Action="&Action&"&Page="&I    
 If Page=I Then Response.Write " selected"
 Response.Write ">"& I &"</option>"
Next
Response.write "</select>页"
%></td>
    </tr>
</table>

<%End If%>

<%If Action="webtj" Then%>

<%If RsSet("WebKeep")=0 Then Response.write "系统已经禁用此功能"%>
<div class="filters">
			 <h3>页面统计</h3>
</div>
        <table width="100%">

<tr>
            <form name="form6" method="post" action="?">
              <td colspan="5" >选择查询的日期 
                <Select  name="adddate" size="1" onChange="window.location=form.adddate.options[form.adddate.selectedIndex].value">
                  <%Sql="Select AddDate From CFCount_Web_Day Where UserName='"&UserName&"' Group By AddDate Order By AddDate Desc"
Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.open Sql,Conn,1,1
While Not Rs.Eof 
 Response.Write "<option value='"&filename&"?Action=webtj&AddDate="&Rs("AddDate")&"'"
 If Request("AddDate")=Cstr(Rs("AddDate")) Then Response.Write " selected"
 Response.Write ">"&Rs("AddDate")&"</option>"  
Rs.MoveNext
Wend  
%>
                </select></td>
            </form>
          </tr>
          <%
Px=ChkStr(Request("Px"),1)
AddDate=ChkStr(Request("AddDate"),3)

If Px="" Then  Px="MyCounter"
If AddDate="" Then AddDate=Cstr(Date)

Call PxFilter(Px,"id,weburl,MyCounter,adddate,lasttime")

If AddDate<>"" Then
 If IsDate(Cdate(AddDate))=False Then Call AlertBack("输入了错误的时间",1)
End if

Sql="Select Sum(MyCounter) From CFCount_Web_Day where UserName='"&UserName&"' And Datediff('d',AddDate,'"&AddDate&"')=0 And WebUrl<>'-'"
Set Rs=Conn.execute(Sql)
ThisDayTotal=Rs(0)
If ThisDayTotal=0 Then ThisDayTotal=1

Sql="Select * From CFCount_Web_Day where UserName='"&UserName&"' And Datediff('d',AddDate,'"&AddDate&"')=0 And WebUrl<>'-'"

Sql=Sql&" Order By "&Px&" Desc"
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
          <tr >
            <td colspan="5"><%=AddDate%> 页面浏览统计[可点击标题排序]</td>
          </tr>
          <tr > 
            <td class="td_nowrap"><a href="?Action=<%=Action%>&AddDate=<%=AddDate%>&Px=ID">序号</a></td>
            <td class="td_nowrap"><a href="?Action=<%=Action%>&AddDate=<%=AddDate%>&Px=WebUrl">浏览页面</a></td>
            <td class="td_nowrap"><a href="?Action=<%=Action%>&AddDate=<%=AddDate%>&Px=MyCounter">浏览数量</a></td>
            <td class="td_nowrap"><a href="?Action=<%=Action%>&AddDate=<%=AddDate%>&Px=MyCounter">所占此天的比例</a></td>
            <td class="td_nowrap"><a href="?Action=<%=Action%>&AddDate=<%=AddDate%>&Px=LastTime">最后访问时间</a></td>
          </tr>
          <%
While Not Rs.Eof And MyPageSize>0
%>
          <tr > 
            <td class="td_nowrap"><%=Rs.RecordCount-Rs.Pagesize*(page)+mypagesize%></td>
            <td class="td_nowrap"> 
                <%
  Response.Write "<a href="&Rs("WebUrl")&" target='_blank'>"&Left(BreakUrl(Rs("WebUrl"),2),30)&"</a>&nbsp;[<a href="&Rs("WebUrl")&" target='_blank'>详细</a>]"
%>
            </td>
            <td class="td_nowrap"><%=Rs("MyCounter")%></td>
            <td class="td_nowrap td_1"><img src=images/BlueBar.gif width=<%=Rs("MyCounter")/ThisDayTotal*150%> height=10><%=MyRate(Rs("MyCounter"),ThisDayTotal)%>%</td>
            <td class="td_nowrap"><%=GetTurnTime(Hour(Rs("LastTime")))&":"&GetTurnTime(Minute(Rs("LastTime")))&":"&GetTurnTime(Second(Rs("LastTime")))%></td>
          </tr>
          <%rs.movenext
mypagesize=mypagesize-1
wend%>


        </table>
<table class="tb_3">
  <tr>
    <td><a href="?Action=<%=Action%>&AddDate=<%=AddDate%>&Px=<%=Px%>&Page=1">第一页</a>
            <%
if page>1 then%>
            <a href='?Action=<%=Action%>&AddDate=<%=AddDate%>&Px=<%=Px%>&Page=<%=Page-1%>'>上一页</a>
            <%
end if
%>
            <%
if page<rs.pagecount   then%>
            <a href='?Action=<%=Action%>&AddDate=<%=AddDate%>&Px=<%=Px%>&Page=<%=Page+1%>'>下一页</a>
            <%

end if
%>
            <a href="?Action=<%=Action%>&AddDate=<%=AddDate%>&Px=<%=Px%>&Page=<%=totalpage%>">最后一页</a> 页次：<font color="#ff0000"><%=page%></font>/<%=totalpage%>页&nbsp;&nbsp;共<%=totalrs%>条记录&nbsp;&nbsp;每页显示<%=rs.pagesize%>条
<%
Response.write "&nbsp;&nbsp;转到第<select id='page' onChange=""window.location=document.getElementById('page').options[document.getElementById('page').selectedIndex].value"">"
For I=1 To TotalPage
    Response.Write "<option value=?Action="&Action&"&AddDate="&AddDate&"&Px="&Px&"&Page="&i
 If Page=I Then Response.Write " selected"
 Response.Write ">"& I &"</option>"
Next
Response.write "</select>页"
%>
    </td>

  </tr>
</table>
<%End If%>



<%If Action="urlcounttj" Then%>

<%If RsSet("WebKeep")=0 Then Response.write "系统已经禁用此功能"%>
<div class="filters">
			 <h3>链接统计</h3>
</div>
<table width="100%">

<tr>
            <form name="form6" method="post" action="?">
              <td colspan="7" >选择查询的日期 
                <Select  name="CreateDate" size="1" onChange="window.location=form.CreateDate.options[form.CreateDate.selectedIndex].value">
                  <%Sql="Select CreateDate From Url_count where Uid='"&UserName&"' Order By CreateDate Desc"
Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.open Sql,Conn,1,1
While Not Rs.Eof 
 Response.Write "<option value='"&filename&"?Action=urlcounttj&CreateDate="&Rs("CreateDate")&"'"
 If Request("CreateDate")=Cstr(Rs("CreateDate")) Then Response.Write " selected"
 Response.Write ">"&Rs("CreateDate")&"</option>"  
Rs.MoveNext
Wend  
%>
                </select></td>
            </form>
          </tr>
          <%
Px=ChkStr(Request("Px"),1)
CreateDate=ChkStr(Request("CreateDate"),3)

If Px="" Then  Px="Gid"
If CreateDate="" Then CreateDate=Cstr(Date)

Call PxFilter(Px,"id,Gid,Url,CreateDate,LastDate,UrlCount")

If CreateDate<>"" Then
 If IsDate(Cdate(CreateDate))=False Then Call AlertBack("输入了错误的时间",1)
End if

Sql="Select Sum(UrlCount) From Url_count where Uid='"&UserName&"' And Datediff('d',CreateDate,'"&CreateDate&"')=0 "
Set Rs=Conn.execute(Sql)
 ThisUrlCountTotal=Rs(0)
If ThisUrlCountTotal=0 Then ThisUrlCountTotal=1

Sql="Select * From Url_count where Uid='"&UserName&"' And Datediff('d',CreateDate,'"&CreateDate&"')=0 "

Sql=Sql&" Order By "&Px&" Desc"
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
          <tr >
            <td height="26" colspan="7"><%=CreateDate%>链接浏览统计[可点击标题排序]  邮件内部链接点击总数[<%=ThisUrlCountTotal%>]次</td>
          </tr>
          <tr > 
            <td nowrap="nowrap" class="td_nowrap"><a href="?Action=<%=Action%>&CreateDate=<%=CreateDate%>&Px=ID">序号</a></td>
            <td nowrap="nowrap" class="td_nowrap"><a href="?Action=<%=Action%>&CreateDate=<%=CreateDate%>&Px=Gid">邮件编号</a></td>
            <td nowrap="nowrap" class="td_nowrap"><a href="?Action=<%=Action%>&CreateDate=<%=CreateDate%>&Px=mail">邮件地址</a></td>
            <td nowrap="nowrap" class="td_nowrap"><a href="?Action=<%=Action%>&CreateDate=<%=CreateDate%>&Px=Url">链接地址</a></td>
            <td nowrap="nowrap" class="td_nowrap"><a href="?Action=<%=Action%>&CreateDate=<%=CreateDate%>&Px=CreateDate">最先访问</a></td>
            <td nowrap="nowrap" class="td_nowrap"><a href="?Action=<%=Action%>&CreateDate=<%=CreateDate%>&Px=UrlCount">点击数量</a></td>
            <td nowrap="nowrap" class="td_nowrap"><a href="?Action=<%=Action%>&CreateDate=<%=CreateDate%>&Px=LastDate">最后访问</a></td>
          </tr>
          <%
While Not Rs.Eof And MyPageSize>0
%>
          <tr > 
            <td class="td_nowrap"><%=Rs.RecordCount-Rs.Pagesize*(page)+mypagesize%></td>
             <td class="td_nowrap"><%=Rs("Gid")%></td>
             <td class="td_nowrap"><%=Rs("Mail")%></td>
            <td class="td_nowrap"> 
                <%
  Response.Write "<a href="&Rs("Url")&" target='_blank'>"&Left(BreakUrl(Rs("Url"),2),30)&"</a>&nbsp;[<a href="&Rs("Url")&" target='_blank'>详细</a>]"
%>
            </td>
            <td class="td_nowrap"><%=Rs("CreateDate")%></td>
            <td class="td_nowrap td_1">共<%=Rs("UrlCount")%>次  <img src=images/BlueBar.gif width=<%=Rs("UrlCount")/ThisUrlCountTotal*150%> height=10><%=MyRate(Rs("UrlCount"),ThisUrlCountTotal)%>%</td>
            <td class="td_nowrap"><%=Rs("LastDate")%></td>
          </tr>
          <%rs.movenext
mypagesize=mypagesize-1
wend%>


        </table>
<table class="tb_3">
  <tr>
    <td><a href="?Action=<%=Action%>&CreateDate=<%=CreateDate%>&Px=<%=Px%>&Page=1">第一页</a>
            <%
if page>1 then%>
            <a href='?Action=<%=Action%>&CreateDate=<%=CreateDate%>&Px=<%=Px%>&Page=<%=Page-1%>'>上一页</a>
            <%
end if
%>
            <%
if page<rs.pagecount   then%>
            <a href='?Action=<%=Action%>&CreateDate=<%=CreateDate%>&Px=<%=Px%>&Page=<%=Page+1%>'>下一页</a>
            <%

end if
%>
            <a href="?Action=<%=Action%>&CreateDate=<%=CreateDate%>&Px=<%=Px%>&Page=<%=totalpage%>">最后一页</a> 页次：<font color="#ff0000"><%=page%></font>/<%=totalpage%>页&nbsp;&nbsp;共<%=totalrs%>条记录&nbsp;&nbsp;每页显示<%=rs.pagesize%>条
<%
Response.write "&nbsp;&nbsp;转到第<select id='page' onChange=""window.location=document.getElementById('page').options[document.getElementById('page').selectedIndex].value"">"
For I=1 To TotalPage
    Response.Write "<option value=?Action="&Action&"&CreateDate="&CreateDate&"&Px="&Px&"&Page="&i
 If Page=I Then Response.Write " selected"
 Response.Write ">"& I &"</option>"
Next
Response.write "</select>页"
%>
    </td>

  </tr>

</table>
<%End If%>


<%If Action="daytj" Then%>
<div class="filters">
			 <h3>日访问统计[可点击标题排序] </h3>
</div>
<table>
<tr >
            <td>
<%
Sql="Select Sum(MyCounter),Sum(IpCounter),Count(ID) From CFCount_Count_Day Where UserName='"&UserName&"'"
Set Rs=Conn.ExeCute(Sql)

If TotalRs=0 Then
 AvgVisit=0
 AvgIp=0
Else
 AvgVisit=Int(Rs(0)/TotalRs)
 AvgIp=Int(Rs(1)/TotalRs)
End If

ShowTotal=Rs(0)
IpTotal=Rs(1)
If ShowTotal=0 Then ShowTotal=1
If IpTotal=0 Then IpTotal=1
%>


              共有<%=Rs(2)%>天记录，共<%=Rs(0)%>次浏览，<%=Rs(1)%>个访问，平均每天<%=AvgVisit%>次浏览，平均每天<%=AvgIp%>个访问</td>
  </tr>

</table> 
        <table width="100%">
<%
Sql="Select Top 1 IpCounter,AddDate From CFCount_Count_Day Where UserName='"&UserName&"' Order By IpCounter Desc"
Set Rs1=Conn.ExeCute(Sql)
If Not Rs1.Eof Then
 BestHightIp=Rs1("IpCounter")
 BestHightIpDate=Rs1("AddDate")
Else
 BestHightIp=0
 BestHightIpDate=""
End If
Rs1.Close

Sql="Select Top 1 IpCounter,AddDate From CFCount_Count_Day Where UserName='"&UserName&"' Order By IpCounter"
Set Rs1=Conn.ExeCute(Sql)
If Not Rs1.Eof Then
 BestLowIp=Rs1("IpCounter")
 BestLowIpDate=Rs1("AddDate")
Else
 BestLowIp=0
 BestLowIpDate=""
End If
Rs1.Close

Sql="Select Top 1 MyCounter,AddDate From CFCount_Count_Day Where UserName='"&UserName&"' And DateDiff('d',AddDate,Date())>0 Order By MyCounter Desc"
Set Rs1=Conn.ExeCute(Sql)
If Not Rs1.Eof Then
 BestHightVisit=Rs1("MyCounter")
 BestHightVisitDate=Rs1("AddDate")
Else
 BestHightVisit=0
 BestHightVisitDate=""
End If
Rs1.Close

Sql="Select Top 1 MyCounter,AddDate From CFCount_Count_Day Where UserName='"&UserName&"' And DateDiff('d',AddDate,Date())>0 Order By MyCounter"
Set Rs1=Conn.ExeCute(Sql)
If Not Rs1.Eof Then
 BestLowVisit=Rs1("MyCounter")
 BestLowVisitDate=Rs1("AddDate")
Else
 BestLowVisit=0
 BestLowVisitDate=""
End If
Rs1.Close
%><tr >
            <td colspan="4">
                <%
Px=ChkStr(request("Px"),1)
If Px="" Then Px="ID"
Call PxFilter(Px,"id,MyCounter,ipcounter,adddate")

Sql="Select * From CFCount_Count_Day Where UserName='"&UserName&"' Order By "&Px&" Desc"

Set Rs=Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,Conn,1,1
If Not Rs.Eof Then
If not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Then
 page=1
Else
 Page=Int(Abs(Request("page")))
End if
rs.pagesize=10
totalrs=rs.RecordCount
totalpage=rs.pageCount
mypagesize=rs.pagesize
rs.absolutepage=page
End If


%>
                <br>
    </td>
          </tr>
          <tr > 
            <td><a href="?Action=<%=Action%>&Px=AddDate">日期</a></td>
            <td><a href="?Action=<%=Action%>&Px=MyCounter">显示数量</a></td>
            <td><a href="?Action=<%=Action%>&Px=IpCounter">IP数量</a></td>
            <td><a href="?Action=<%=Action%>&Px=MyCounter">在十天内的比率</a></td>
          </tr>
          <%While Not Rs.Eof And MyPageSize>0%>

          <tr > 
            <td><%=Rs("AddDate")%></td>
            <td><%=Rs("MyCounter")%><br>
            </td>
            <td><%=Rs("IpCounter")%></td>
            <td class="td_1"><img src=images/BlueBar.gif width=<%=Rs("MyCounter")/ShowTotal*300%> height=10><%=MyRate(Rs("MyCounter"),ShowTotal)%>%<br>
              <img src=images/GreenBar.gif width=<%=Rs("IpCounter")/IpTotal*300%> height=10><%=MyRate(Rs("IpCounter"),IpTotal)%>%</td>
          </tr>
          <%Rs.MoveNext
MyPageSize=MyPageSize-1
Wend%>
          <tr> 
            <td colspan="4" >图例： <img src=images/BlueBar.gif width=30 height=10>显示数，<img src=images/GreenBar.gif width=30 height=10>IP数</td>
          </tr>
          <tr> 
            <td colspan="4"> <%
Sql="Select Sum(MyCounter),Sum(IpCounter) From CFCount_Count_Day Where UserName='"&UserName&"'"
Set Rs1=Conn.ExeCute(Sql)

If TotalRs=0 Then
 AvgVisit=0
 AvgIp=0
Else
 AvgVisit=Int(Rs1(0)/TotalRs)
 AvgIp=Int(Rs1(1)/TotalRs)
End If
%>
              共有<%=TotalRs%>天记录，共<%=Rs1(0)%>次浏览，<%=Rs1(1)%>个访问，平均每天<%=AvgVisit%>次访问，平均每天IP<%=AvgIp%>个</td>
          </tr>
          <tr>
            <td colspan="4">
</td>
          </tr>
        </table>
  <table class="tb_3">
    <tr> 
      <td><a href="?Action=<%=Action%>&Px=<%=Px%>&Page=1">第一页</a> 
          <%
if page>1 then%>
          <a href='?Action=<%=Action%>&Px=<%=Px%>&Page=<%=Page-1%>'>上一页</a> 
          <%
end if
%>
          <%
if page<rs.pagecount   then%>
          <a href='?Action=<%=Action%>&Px=<%=Px%>&Page=<%=Page+1%>'>下一页</a> 
          <%
end if
%>
          <a href="?Action=<%=Action%>&Px=<%=Px%>&Page=<%=TotalPage%>">最后一页</a> 
          页次：<font color="#ff0000"><%=Page%></font>/<%=TotalPage%>页
<%
Response.write "&nbsp;&nbsp;转到第<select id='page' onChange=""window.location=document.getElementById('page').options[document.getElementById('page').selectedIndex].value"">"
For I=1 To TotalPage
    Response.Write "<option value=?action="&action&"&px="&px&"&page="&i     
 If Page=I Then Response.Write " selected"
 Response.Write ">"& I &"</option>"
Next
Response.write "</select>页"
%>
                
</td>
    </tr>
</table></td>
    </tr>
  </table>
<%End If%>

<%If Action="monthtj" Then%>
<div class="filters">
			 <h3>最近12个月访问统计</h3>
</div>
        <table width="100%">
          <tr > 
            <td>月份</td>
            <td>显示数量</td>
            <td>IP数量</td>
            <td>在12个月内的比率及数量</td>
          </tr>
          <%
ShowTotal=0
IpTotal=0
Sql="Select Sum(MyCounter),Sum(IpCounter) From CFCount_Count_Day Where UserName='"&UserName&"' And DateDiff('m',AddDate,Date())<12"
Set Rs=Conn.Execute(Sql)
ShowTotal=Rs(0)
IpTotal=Rs(1)
If ShowTotal=0 Then ShowTotal=1
If IpTotal=0 Then IpTotal=1

For I=0 To 11
CurrMonth=Dateadd("m",-I,Date())

Sql="Select Sum(MyCounter),Sum(IpCounter) From CFCount_Count_Day Where UserName='"&UserName&"' And DateDiff('m',AddDate,'"&CurrMonth&"')=0"
Set Rs=Conn.ExeCute(Sql)
%>
          <tr > 
            <td><%=year(CurrMonth)&"-"&month(CurrMonth)%></td>
            <td><%=Rs(0)%></td>
            <td><%=Rs(1)%></td>
            <td class="td_1"><img src=images/BlueBar.gif width=<%=Rs(0)/ShowTotal*300%> height=10><%=MyRate(Rs(0),ShowTotal)%>%<br><img src=images/GreenBar.gif width=<%=Rs(1)/IpTotal*300%> height=10><%=MyRate(Rs(1),IpTotal)%>%</td>
          </tr>
          <%Next%>
          <tr> 
            <td colspan="6" >图例： <img src=images/BlueBar.gif width=30 height=10>显示数，<img src=images/GreenBar.gif width=30 height=10>IP数</td>
          </tr>
          <tr> 
            <td colspan="4"> 共<%=ShowTotal%>次浏览，<%=IpTotal%>个访问，平均每月<%=Int(ShowTotal/12)%>次浏览，平均每月<%=Int(IpTotal/12)%>个访问</td>
          </tr>
        </table>
  <%end if%>
  
  <%if Action="hourtj" then%>
  <div class="filters">
			 <h3><%=AddDate%> 每小时报表</h3>
</div>
        <table>
          <tr> 
      <form name="form1" method="post" action="jczx.asp">
              <td>选择查询日期 
                <Select  name="a" size="1" onChange="window.location=form.a.options[form.a.selectedIndex].value">
                  <%Sql="Select AddDate From CFCount_Count_Hour Where UserName='"&UserName&"' Group By AddDate Order By AddDate Desc"
Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.open Sql,Conn,1,1
While Not Rs.Eof 
 Response.Write "<option value='"&filename&"?Action="&Action&"&AddDate="&Rs("AddDate")&"'"
 If Request("AddDate")=Cstr(Rs("AddDate")) Then Response.Write " selected"
 Response.Write ">"&Rs("AddDate")&"</option>"  
Rs.MoveNext
Wend  
%>
                </select> </td>
      </form>
    </tr>
  </table>
  
    
  <%
AddDate=ChkStr(Request("AddDate"),3)
If AddDate="" Then AddDate=Date()

Sql="Select Sum(MyCounter),Sum(IpCounter) From CFCount_Count_Hour where UserName='"&UserName&"' And AddDate=#"&AddDate&"#"
Set Rs=Conn.Execute(Sql)
ShowTotal=Rs(0)
IpTotal=Rs(1)
If ShowTotal=0 Then ShowTotal=1
If IpTotal=0 Then IpTotal=1

Sql="Select AddHour,AddDate From CFCount_Count_Hour Where UserName='"&UserName&"' And AddDate=#"&AddDate&"# Group By AddHour,AddDate Order By AddHour Desc"
Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.open Sql,Conn,1,1
%>
        <table width="100%">
          <tr > 
            <td >小时</td>
            <td >显示数量</td>
            <td >访问者数量</td>
            <td >比率</td>
          </tr>
          <%
do while not rs.eof
I=I+1

Sql="Select MyCounter,IpCounter From CFCount_Count_Hour where UserName='"&UserName&"' And AddDate=#"&AddDate&"# And AddHour="&Rs("AddHour")
Set Rs1=Conn.Execute(Sql)
%>
          <tr > 
            <td><%=Rs("AddHour")%>-<%=Rs("AddHour")+1%></td>
            <td><%=Rs1(0)%><br>
            </td>
            <td><%=Rs1(1)%></td>
            <td class="td_1"><img src=images/BlueBar.gif width=<%=Rs1(0)/ShowTotal*350%> height=10><%=MyRate(Rs1(0),ShowTotal)%>%<br><img src=images/GreenBar.gif width=<%=Rs1(1)/IpTotal*350%> height=10><%=MyRate(Rs1(1),IpTotal)%>%</td>
          </tr>
          <%
rs.movenext
loop%>
          <tr> 
            <td colspan="6" >图例： <img src=images/BlueBar.gif width=30 height=10>显示数，<img src=images/GreenBar.gif width=30 height=10>IP数</td>
          </tr>
		  
          <tr> 
            <td colspan="4"><%=AddDate%>此日期：总计浏览<%=ShowTotal%>次 <%=IpTotal%>个访问者</td>
          </tr>
        </table>
  

<%end if%>
        <%if Action="backtj" then%>
<div class="filters">
			 <h3>回头率报表</h3>
</div>
        <table width="100%">
          <tr > 
            <td>次数</td>
            <td>人数</td>
            <td>比率</td>
          </tr>
<%
Sql="Select * From CFCount_Back where UserName='"&UserName&"'"
Set Rs=Conn.Execute(Sql)
TotalNum=Rs("BackNum_1")+Rs("BackNum_2")+Rs("BackNum_3")+Rs("BackNum_4")+Rs("BackNum_5")+Rs("BackNum_6")+Rs("BackNum_7")+Rs("BackNum_8")+Rs("BackNum_9")+Rs("BackNum_10")+Rs("BackNum_Higher")

For I= 1 To 10
%>
          <tr > 
            <td>第&nbsp;<font color="#FF0000"><%=I%></font>&nbsp;次访问</td>
            <td><%=Rs("BackNum_"&I)%></td>
            <td class="td_1"><img src=images/BlueBar.gif width=<%=Rs("BackNum_"&I)/TotalNum*300%> height=10><%=MyRate(Rs("BackNum_"&I),TotalNum)%>% </td>
          </tr>
          <%
Next%>
<tr > 
            <td>10次以上访问</td>
            <td><%=Rs("BackNum_Higher")%></td>
            <td class="td_1"><img src=images/BlueBar.gif width=<%=Rs("BackNum_Higher")/TotalNum*300%> height=10><%=MyRate(Rs("BackNum_Higher"),TotalNum)%>% </td>
          </tr>
        </table>
        <%end if%>


<%If Action="searchtj" Then%>
<%
AddDate=ChkStr(Request("AddDate"),3)

If AddDate="" Then  AddDate=Date()
%>
<div class="filters">
			 <h3>搜索引擎统计</h3>
</div>
        <table width="100%">

  <tr> 
<form name="form1" method="post" action="jczx.asp"><td colspan="3"> <select  name="a" size="1" onChange="window.location=form.a.options[form.a.selectedIndex].value">
        <%Sql="Select AddDate From CFCount_Search_Day Where UserName='"&UserName&"' Group By AddDate Order By AddDate Desc"
Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.open Sql,Conn,1,1
While Not Rs.Eof 
 Response.Write "<option value='?Action="&Action&"&AddDate="&Rs("AddDate")&"'"
 If Cstr(AddDate)=Cstr(Rs("AddDate")) Then Response.Write " selected"
 Response.Write ">"&Rs("AddDate")&"</option>"  
Rs.MoveNext
Wend  
%>
      </select> </td></form>
  </tr>
  <tr > 
    <td >搜索引擎英文标识</td>
    <td >PV数量</td>
    <td >IP数量</td>
  </tr>
  <%
Sql="Select * From CFCount_Search_Day Where Datediff('d',AddDate,'"&AddDate&"')=0 Order By IpCounter Desc"
Set Rs=Conn.Execute(Sql)
%>
  <%While Not Rs.Eof%>
  <tr > 
    <td><%=Rs("SiteFlag")%></td>
    <td><%=Rs("MyCounter")%></td>
    <td><%=Rs("IpCounter")%></td>
  </tr>
  <%Rs.MoveNext
Wend%>
</table>
  <%end if%>
  
<%If Action="alexatj" Then%>
<div class="filters">
			 <h3>日Alexa统计[可点击标题排序]</h3>
</div>
        <table width="100%">
          <tr >
            <td colspan="3"> 
<%
Sql="Select Sum(MyCounter) From CFCount_Alexa_Day Where UserName='"&UserName&"'"
Set Rs=Conn.Execute(Sql)
AlexaTotal=Rs(0)
If AlexaTotal=0 Then AlexaTotal=1
%>
<%
Px=ChkStr(request("Px"),1)
If Px="" Then Px="ID"
Call PxFilter(Px,"id,MyCounter,adddate")

Sql="Select * From CFCount_Alexa_Day Where UserName='"&UserName&"' Order By "&Px&" Desc"

Set Rs=Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,Conn,1,1
If Not Rs.Eof Then
If not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Then
 page=1
Else
 Page=Int(Abs(Request("page")))
End if
rs.pagesize=10
totalrs=rs.RecordCount
totalpage=rs.pageCount
mypagesize=rs.pagesize
rs.absolutepage=page
End If

%>
            </td>
          </tr>
          <tr > 
            <td><a href="?Action=<%=Action%>&Px=AddDate">日期</a></td>
            <td><a href="?Action=<%=Action%>&Px=MyCounter">安装数量</a></td>
            <td><a href="?Action=<%=Action%>&Px=MyCounter">在十天内的比率</a></td>
          </tr>
          <%While Not Rs.Eof And MyPageSize>0%>

          <tr > 
            <td><%=Rs("AddDate")%></td>
            <td><%=Rs("MyCounter")%><br>
            </td>
            <td class="td_1"><img src=images/BlueBar.gif width=<%=Rs("MyCounter")/AlexaTotal*300%> height=10><%=MyRate(Rs("MyCounter"),AlexaTotal)%>%</td>
          </tr>
          <%Rs.MoveNext
MyPageSize=MyPageSize-1
Wend%>
        </table>
  <table class="tb_3">
    <tr> 
      <td><a href="?Action=<%=Action%>&Px=<%=Px%>&Page=1">第一页</a> 
          <%
if page>1 then%>
          <a href='?Action=<%=Action%>&Px=<%=Px%>&Page=<%=Page-1%>'>上一页</a> 
          <%
end if
%>
          <%
if page<rs.pagecount   then%>
          <a href='?Action=<%=Action%>&Px=<%=Px%>&Page=<%=Page+1%>'>下一页</a> 
          <%
end if
%>
          <a href="?Action=<%=Action%>&Px=<%=Px%>&Page=<%=TotalPage%>">最后一页</a> 
          页次：<font color="#ff0000"><%=Page%></font>/<%=TotalPage%>页
                
<%
Response.write "&nbsp;&nbsp;转到第<select id='page' onChange=""window.location=document.getElementById('page').options[document.getElementById('page').selectedIndex].value"">"
For I=1 To TotalPage
    Response.Write "<option value=?action="&action&"&px="&px&"&page="&i     
 If Page=I Then Response.Write " selected"
 Response.Write ">"& I &"</option>"
Next
Response.write "</select>页"
%></td>
    </tr>
</table></td>
    </tr>
  </table>
<%End If%>


<%If Action="keywordtj" Then%>
<%
SiteFlag=ChkStr(Request("SiteFlag"),1)
AddDate=ChkStr(Request("AddDate"),3)

If AddDate="" Then  AddDate=Date()
%>
<div class="filters">
			 <h3>关键字统计[可点击标题排序]</h3>
</div>
        <table width="100%">
          <tr>
    <form name="form1" method="post" action="?Action=<%=Action%>">
      <td colspan="5"> <select  name="siteflag">
          <%Sql="Select SiteDesc,SiteFlag From CFCount_Search_Set Order By SiteFlag Desc"
Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.open Sql,Conn,1,1
While Not Rs.Eof 
 Response.Write "<option value='"&Rs("SiteFlag")&"'"
 If Cstr(SiteFlag)=Cstr(Rs("SiteFlag")) Then Response.Write " selected"
 Response.Write ">"&Rs("SiteDesc")&"</option>"  
 
 If SiteFlag="" Then
  If Instr(Rs("SiteFlag"),"baidu.com") Then
   SiteFlag = Rs("SiteFlag")
   SiteDesc = Rs("SiteDesc") 
  End If
 Else
  If Cstr(SiteFlag)=Cstr(Rs("SiteFlag")) Then
   SiteDesc = Rs("SiteDesc") 
  End if
 End If
Rs.MoveNext
Wend  
%>
        </select> <select  name="adddate">
          <%Sql="Select AddDate From CFCount_SearchKeywrod_Day Where UserName='"&UserName&"' Group By AddDate Order By AddDate Desc"
Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.open Sql,Conn,1,1
While Not Rs.Eof 
 Response.Write "<option value='"&Rs("AddDate")&"'"
 If Cstr(AddDate)=Cstr(Rs("AddDate")) Then Response.Write " selected"
 Response.Write ">"&Rs("AddDate")&"</option>"  
Rs.MoveNext
Wend  
%>
        </select> <input type="submit" name="Submit3" value="确定"> </td>
    </form>
  </tr>
  <%
Px=ChkStr(request("Px"),1)
If Px="" Then Px="IpCounter"
Call PxFilter(Px,"KeyWord,IpCounter,MyCounter,AddDate,LastTime,LastLy")


Sql="Select * From CFCount_SearchKeywrod_Day Where UserName='"&UserName&"' And SiteFlag='"&SiteFlag&"' And AddDate=#"&AddDate&"# Order By  "&Px&"  Desc"
Set Rs=Server.CreateObject("Adodb.RecordSet")
Rs.Open Sql,Conn,1,1

If Not Rs.Eof Then
If not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Then
 page=1
Else
 Page=Int(Abs(Request("page")))
End if
rs.pagesize=20
totalrs=rs.RecordCount
totalpage=rs.pageCount
mypagesize=rs.pagesize
rs.absolutepage=page
End If
%>
  <tr >
    <td><%=SiteDesc%></td>
    <td><a href="?Action=<%=Action%>&SiteFlag=<%=SiteFlag%>&AddDate=<%=AddDate%>&Px=Keyword">关键字</a></td>
    <td><a href="?Action=<%=Action%>&SiteFlag=<%=SiteFlag%>&AddDate=<%=AddDate%>&Px=IpCounter">PV数量</a></td>
    <td><a href="?Action=<%=Action%>&SiteFlag=<%=SiteFlag%>&AddDate=<%=AddDate%>&Px=MyCounter">IP数量</a></td>
    <td><a href="?Action=<%=Action%>&SiteFlag=<%=SiteFlag%>&AddDate=<%=AddDate%>&Px=AddDate">加入日期</a></td>
    <td><a href="?Action=<%=Action%>&SiteFlag=<%=SiteFlag%>&AddDate=<%=AddDate%>&Px=LastTime">最后更新时间</a></td>
  </tr>
  <%While Not Rs.Eof And MyPageSize>0%>
  <tr > 
    <td></td>
    <td>
<a href="<%=Rs("LastLy")%>" target='_blank'> 
<%
 If Instr(Rs("LastLy"),"google")>0 Or Instr(Rs("LastLy"),"yahoo")>0 Then
  KeyWord=UTF2GB(Rs("KeyWord"))  
 Else
  KeyWord=URLDecode(Rs("KeyWord"))
 End If
 If Len(Keyword)=0 Then KeyWord=Rs("KeyWord")
 Response.write Server.HTMLEncode(KeyWord)
%>
</a>
</td>
    <td><%=Rs("MyCounter")%></td>
    <td><%=Rs("IpCounter")%></td>
    <td><%=GetTurnTime(Month(Rs("AddDate")))&"-"&GetTurnTime(Day(Rs("AddDate")))%></td>
    <td><%=Rs("LastTime")%></td>
  </tr>
  <%Rs.MoveNext
		  MyPageSize=MyPageSize-1
Wend%>
</table>
		<table class="tb_3">
    <tr> 
      <td><a href="?Action=<%=Action%>&SiteFlag=<%=SiteFlag%>&AddDate=<%=AddDate%>&Px=<%=Px%>&Page=1">第一页</a> 
          <%
if page>1 then%>
          <a href='?Action=<%=Action%>&SiteFlag=<%=SiteFlag%>&AddDate=<%=AddDate%>&Px=<%=Px%>&Page=<%=Page-1%>'>上一页</a> 
          <%
end if
%>
          <%
if page<rs.pagecount   then%>
          <a href='?Action=<%=Action%>&SiteFlag=<%=SiteFlag%>&AddDate=<%=AddDate%>&Px=<%=Px%>&Page=<%=Page+1%>'>下一页</a> 
          <%
end if
%>
          <a href="?Action=<%=Action%>&SiteFlag=<%=SiteFlag%>&AddDate=<%=AddDate%>&Px=<%=Px%>&Page=<%=TotalPage%>">最后一页</a> 
          页次：<font color="#ff0000"><%=Page%></font>/<%=TotalPage%>页
                
<%
Response.write "&nbsp;&nbsp;转到第<select id='page' onChange=""window.location=document.getElementById('page').options[document.getElementById('page').selectedIndex].value"">"
For I=1 To TotalPage
    Response.Write "<option value=?Action="&Action&"&SiteFlag="&SiteFlag&"&AddDate="&AddDate&"&Px="&Px&"&Page="&I     
 If Page=I Then Response.Write " selected"
 Response.Write ">"& I &"</option>"
Next
Response.write "</select>页"
%></td>
          </tr>
</table></td>
    </tr>
  </table>
  <%end if%>
  
