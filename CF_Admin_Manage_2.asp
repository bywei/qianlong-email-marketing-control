
<%If Action="" Or Action="sysstate" Then%>
<table >
  <tr > 
              <td colspan="2">系统状态</td>
  </tr>
            <tr> 
              <td width="18%">系统名称：</td>
              <td ><%=RsSet("SysTitle")%> </td>
            </tr>
            <tr> 
              <td >系统注册状态：</td>
              <td> <%If RsSet("RegState")=0 Then%>
                目前已经禁止新用户申请计数器！ 
                <%Else%>
                目前允许新用户申请计数器，如果只想自己用可以关闭用户注册功能！ 
                <%End If%>
                [<a href="?Action=sysset">更改设置</a>]</td>
            </tr>
            <tr> 
              <td >来源统计共执行次数：</td>
              <td><%=RsSet("Store_TotalLy")%> [<a href="?Action=lyreset" onClick="{if(confirm('确定要复原来源统计数么?')){return true;}return false;};">复原为0</a>]</td>
            </tr>
            <tr> 
              <td >页面来源总数：</td>
              <td><%
Sql="Select Count(*) From CFCount_Ly_Day"
Set Rs=Conn.ExeCute(Sql)
Response.Write Rs(0)
%>
                [<a href="?Action=beforelydel" onClick="{if(confirm('确定要删除今天以前的来源么?')){return true;}return false;};">删除今天以前的来源</a>]</td>
            </tr>
            <tr> 
              <td >今天页面来源数：</td>
              <td><%
Sql="Select Count(*) From CFCount_Ly_Day Where DateDiff('d',AddDate,Date())=0"
Set Rs=Conn.ExeCute(Sql)
Response.Write Rs(0)
%> &nbsp;</td>
            </tr>
            <tr> 
              <td >页面统计总数：</td>
              <td><%
Sql="Select Count(*) From CFCount_Web_Day"
Set Rs=Conn.ExeCute(Sql)
Response.Write Rs(0)
%>
                [<a href="?Action=beforewebdel" onClick="{if(confirm('确定要删除今天以前的页面统计么?')){return true;}return false;};">删除今天以前的页面统计</a>]</td>
            </tr>
            <tr> 
              <td >今天的页面统计数：</td>
              <td><%
Sql="Select Count(*) From CFCount_Web_Day Where DateDiff('d',AddDate,Date())=0"
Set Rs=Conn.ExeCute(Sql)
Response.Write Rs(0)
%></td>
            </tr>
</table>
          <%end if%>


<%If Action="userlist" Or Action="searchuser" Then%>
           

<table >
  <tr>
    <td colspan="9">点击标题可以选择其相应的排序方式</td>
  </tr>
  <tr > 
    <td><a href="?Action=<%=Action%>&UserName=<%=UserName%>&Px=ID">ID</a></td>
    <td><a href="?Action=<%=Action%>&UserName=<%=UserName%>&Px=UserName">用户名</a></td>
    <td><a href="?Action=<%=Action%>&UserName=<%=UserName%>&Px=ParentName">账号情况</a></td>
    <td><a href="?Action=<%=Action%>&UserName=<%=UserName%>&Px=LoginTotal">登陆次数</a></td>
    <td><a href="?Action=<%=Action%>&UserName=<%=UserName%>&Px=RealIpTotal">实际计数</a></td>
    <td><a href="?Action=<%=Action%>&UserName=<%=UserName%>&Px=PageName">主页</a></td>
    <td><a href="?Action=<%=Action%>&UserName=<%=UserName%>&Px=AddTime">注册时间</a></td>
    <td><a href="?Action=<%=Action%>&UserName=<%=UserName%>&Px=UserState">帐号状态</a></td>
    <td>操作</td>
  </tr>
<%

Sql="Select Count(*) From CFCount_User" 
Set Rs=Conn.Execute(Sql)
Response.write "<b><font color=ff0000>目前账号总数："&Rs(0)&"</font></b>"

UserName=ChkStr(Request("UserName"),1)
Px=ChkStr(Request("Px"),1)

If Px="" Then Px="ID"

Call PxFilter(Px,"ID,UserName,ParentName,LoginTotal,RealIpTotal,PageName,AddTime,UserState")

Sql="Select Top 2000 * From CFCount_User Where 1=1" 

If Action="searchuser" Then
 If UserName<>"" Then Sql=Sql&" And UserName like '%"&UserName&"%'"
End If

Sql=Sql&" Order By "&px&" Desc"

Set Rs=Server.CreateObject("Adodb.RecordSet")
Rs.open Sql,Conn,1,1

If Not Rs.Eof Then

IF Not IsNumeric(Request("Page")) Or IsEmpty(Request("Page")) Then
 Page=1
Else
 Page=Int(Abs(Request("Page")))
End if

Rs.pagesize=10
TotalRs=Rs.RecordCount
TotalPage=Rs.PageCount
MyPageSize=Rs.PageSize
Rs.AbsolutePage=Page
End If

While Not Rs.Eof And MyPageSize>0%>
  <tr class="tr_2"> 
    <td><%=Rs("ID")%></td>
    <td><%=Rs("UserName")%></td>
    <td>
	<%
	If Rs("ParentName")="" Then
	 Response.write "无父账号"
    Else
	 Response.write "父账号:"&Rs("ParentName")
	End If
	%></td>
    <td><%=Rs("LoginTotal")%></td>
    <td><%=Rs("RealIpTotal")%></td>
    <td><a href="<%=Rs("PageUrl")%>" target="_blank"><%=Rs("PageName")%></a></td>
    <td><%=year(Rs("AddTime"))&"-"&month(Rs("AddTime"))&"-"&day(Rs("AddTime"))%></td>
    <td> 
        <%if rs("UserState")=-1 then
			  response.Write "有效"
			  elseif rs("UserState")=0 then
			  response.Write "<font color=#ff0000>无效</font>"
			  end if%>      </td>
    <td><a href="?Action=usergo&UserName=<%=Rs("UserName")%>" target="_blank">查看</a> 
        <a href="?Action=usermodify&ID=<%=Rs("ID")%>">修改</a> 
        <a href="?Action=userdel&UserName=<%=Rs("UserName")%>" onClick="{if(confirm('确定要删除么?')){return true;}return false;}">删除</a></td>
  </tr>
  <%rs.movenext
mypagesize=mypagesize-1
wend%>
</table>
          <table class="tb_3">
            <tr> 
              <td><a href="?Action=<%=Action%>&UserName=<%=UserName%>&Px=<%=Px%>&Page=1">第一页</a> 
                  <%
if page>1 then%>
                  <a href='?Action=<%=Action%>&UserName=<%=UserName%>&Px=<%=Px%>&Page=<%=Page-1%>'>上一页</a> 
                  <%
end if
%>
                  <%
if page<rs.pagecount   then%>
                  <a href='?Action=<%=Action%>&UserName=<%=UserName%>&Px=<%=Px%>&Page=<%=Page+1%>'>下一页</a> 
                  <%
end if
%>
                  <a href="?Action=<%=Action%>&UserName=<%=UserName%>&Px=<%=Px%>&Page=<%=TotalPage%>">最后一页</a> 
                  页次：<font color="#ff0000"><%=page%></font>/<%=totalpage%>页&nbsp;&nbsp;共<%=totalrs%>条记录&nbsp;&nbsp;每页显示<%=rs.pagesize%>条
<%
Response.write "&nbsp;&nbsp;转到第<select id='page' onChange=""window.location=document.getElementById('page').options[document.getElementById('page').selectedIndex].value"">"
For I=1 To TotalPage
    Response.Write "<option value=?Action="&Action&"&UserName="&UserName&"&Px="&Px&"&Page="&I
 If Page=I Then Response.Write " selected"
 Response.Write ">"& I &"</option>"
Next
Response.write "</select>页"
%></td>
            </tr>
          </table>
          <table class="tb_3">
            <form name="form2" method="post" action="?action=searchuser">
              <tr> 
                <td>搜索：用户名 
                    <input name="UserName" type="text" id="UserName" size="10">
                    <input type="submit" name="Submit" value="查找">
                </td>
              </tr>
            </form>
          </table>
          <%End If%>
          <%If Action="usermodify" Then%>
          <%
ID=ChkStr(Request("ID"),2)
Set Rs=Server.CreateObject("Adodb.RecordSet")
Sql="Select * From CFCount_User Where ID="&ID
Rs.Open Sql,Conn,1,1%>
<table >

  <form name="form2" method="post" action="?Action=usermodifysave&ID=<%=ID%>">
  <tr > 
      <td colspan="2">修改资料 </td>
    </tr>
    <tr> 
      <td width="150">用户名：</td>
      <td><%=Rs("UserName")%></td>
    </tr>
    <tr> 
      <td>修改为新密码：</td>
      <td><input name="Pwd_New" type="text" id="Password_New">
        (留空则原密码不会被修改)</td>
    </tr>
    <tr> 
      <td>修改为新密码找回答案：</td>
      <td><input name="PasswordAnswer_New" type="text" id="PasswordAnswer_New">
        (留空则原密码找回答案不会被修改)</td>
    </tr>
    <tr> 
      <td>帐号状态：</td>
      <td><select name="UserState" id="UserState">
          <option value="-1"<%If Rs("UserState")=-1 Then response.write" selected"%>>有效</option>
          <option value="0"<%If Rs("UserState")=0 Then response.write" selected"%>>无效</option>
        </select></td>
    </tr>
    <tr> 
      <td>管理员说明：</td>
      <td><textarea name="admindesc" cols="30" rows="6" id="notes"><%=rs("admindesc")%></textarea></td>
    </tr>
    <tr> 
      <td colspan="2">  
          <input type="submit" name="Submit3" value="修改">
          　　 
          <input type="reset" name="Submit523" value="取消" onClick="javascript:history.go(-1)">
      </td>
    </tr>
  </form>
</table>
          <%end if%>
          <%If Action="daytj" Then%>
<table >


<%

Sql="Select AddDate From CFCount_Count_Day Group By AddDate Order By AddDate Desc"
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

Sql="Select Sum(MyCounter),Sum(IPCounter) From CFCount_Count_Day"
Set Rs1=Conn.execute(Sql)
ShowTotal=Rs1(0)
IpTotal=Rs1(1)
%>
  <tr > 
              <td colspan="5">每日统计</td>
  </tr>
            <tr class="tr_2"> 
              <td >序号</td>
              <td >日期</td>
              <td >浏览数量</td>
              <td >IP数量</td>
              <td >所占此天的比例</td>
            </tr>
            <%
While Not Rs.Eof And MyPageSize>0
Sql="Select Sum(MyCounter),Sum(IPCounter) From CFCount_Count_Day where Datediff('d',AddDate,'"&Rs("AddDate")&"')=0"
Set Rs1=Conn.execute(Sql)
Counter=Rs1(0)
IpCounter=Rs1(1)


%>
            <tr class="tr_2"> 
              <td><%=Rs.RecordCount-Rs.Pagesize*(page)+mypagesize%></td>
              <td> <%=Rs("AddDate")%> </td>
              <td><%=Counter%></td>
              <td><%=IpCounter%></td>
              <td class="td_1"><img src=images/BlueBar.gif width=<%=Counter/ShowTotal*300%> height=10><%=MyRate(Counter,ShowTotal)%>%<br> <img src=images/GreenBar.gif width=<%=IpCounter/IpTotal*300%> height=10><%=MyRate(IpCounter,IpTotal)%>%</td>
            </tr>
            <%rs.movenext
mypagesize=mypagesize-1
wend%>
            <tr> 
              <td colspan="5" >图例： <img src=images/BlueBar.gif width=30 height=10>显示数，<img src=images/GreenBar.gif width=30 height=10>IP数</td>
            </tr>
            <tr> 
              <td colspan="5"> <%
Sql="Select Sum(MyCounter),Sum(IpCounter) From CFCount_Count_Day"
Set Rs1=Conn.ExeCute(Sql)

If TotalRs=0 Then
 AvgVisit=0
 AvgIp=0
Else
 AvgVisit=Int(Rs1(0)/TotalRs)
 AvgIp=Int(Rs1(1)/TotalRs)
End If
%>
                共有<%=TotalRs%>天记录，共<%=Rs1(0)%>次浏览，<%=Rs1(1)%>个IP，平均每天<%=AvgVisit%>次访问，平均每天IP<%=AvgIp%>个 </td>
            </tr>
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
                  <a href="?Action=<%=Action%>&Page=<%=totalpage%>">最后一页</a> 
                  页次：<font color="#ff0000"><%=page%></font>/<%=totalpage%>页&nbsp;&nbsp;共<%=totalrs%>条记录&nbsp;&nbsp;每页显示<%=rs.pagesize%>条
<%
Response.write "&nbsp;&nbsp;转到第<select id='page' onChange=""window.location=document.getElementById('page').options[document.getElementById('page').selectedIndex].value"">"
For I=1 To TotalPage
    Response.Write "<option value=?Action="&Action&"&Page="&i
 If Page=I Then Response.Write " selected"
 Response.Write ">"& I &"</option>"
Next
Response.write "</select>页"
%></td>

            </tr>
          </table>
          <%End If%>
          <%If Action="lylist" Then%>
<table >

            <form name="form6" method="post" action="?Action=lydel">
<tr>
                <td colspan="7" >选择查询的日期 
                  <Select  name="adddate" size="1" onChange="window.location=form.adddate.options[form.adddate.selectedIndex].value">
                    <%Sql="Select AddDate From CFCount_Ly_Day Group By AddDate Order By AddDate Desc"
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
              </tr>
              <%
Ip=Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If Ip= "" Then Ip= Request.ServerVariables("REMOTE_ADDR")

Px=ChkStr(Request("Px"),1)
AddDate=ChkStr(Request("AddDate"),3)

If Px="" Then  Px="MyCounter"
If AddDate="" Then AddDate=Cstr(Date)

Call PxFilter(Px,"id,ip,lyhead,addtime,MyCounter,lasttime")

If AddDate<>"" Then
 If IsDate(Cdate(AddDate))=False Then Call AlertBack("输入了错误的时间",1)
End if

Sql="Select Sum(MyCounter) From CFCount_Ly_Day where Datediff('d',AddDate,'"&AddDate&"')=0"
Set Rs=Conn.execute(Sql)
ThisDayTotal=Rs(0)
If ThisDayTotal=0 Then ThisDayTotal=1

Sql="Select * From CFCount_Ly_Day where Datediff('d',AddDate,'"&AddDate&"')=0"

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
              <tr class="tr_2"> 
                <td ><a href="?Action=<%=Action%>&AddDate=<%=AddDate%>&Px=ID">序号</a></td>
                <td ><a href="?Action=<%=Action%>&AddDate=<%=AddDate%>&Px=LyHead">来源网站</a></td>
                <td ><a href="?Action=<%=Action%>&AddDate=<%=AddDate%>&Px=MyCounter">来源数量</a></td>
                <td ><a href="?Action=<%=Action%>&AddDate=<%=AddDate%>&Px=MyCounter">所占此天的比例</a></td>
                <td ><a href="?Action=<%=Action%>&AddDate=<%=AddDate%>&Px=AddTime">上站时间</a></td>
                <td ><a href="?Action=<%=Action%>&AddDate=<%=AddDate%>&Px=Ip">最后IP地址</a></td>
                <td ><a href="?Action=<%=Action%>&AddDate=<%=AddDate%>&Px=LastTime">最后访问时间</a></td>
              </tr>
              <%
While Not Rs.Eof And MyPageSize>0
%>
              <tr class="tr_2"> 
                <td><%=Rs.RecordCount-Rs.Pagesize*(page)+mypagesize%></td>
                <td> 
                    <%If Rs("LyHead")="" Then
  Response.Write "直接从浏览器输入"
 Else
  Response.Write "<a href=http://"&Rs("LyHead")&" target='_blank'>"&Rs("LyHead")&"</a>&nbsp;[<a href="&Rs("Ly")&" target='_blank'>详细</a>]"
 End If%>
                </td>
                <td><%=Rs("MyCounter")%></td>
                <td class="td_1"><img src=images/BlueBar.gif width=<%=Rs("MyCounter")/ThisDayTotal*150%> height=10><%=MyRate(Rs("MyCounter"),ThisDayTotal)%>%</td>
                <td><%=GetTurnTime(Hour(Rs("AddTime")))&":"&GetTurnTime(Minute(Rs("AddTime")))&":"&GetTurnTime(Second(Rs("AddTime")))%></td>
                <td> 
                    <%If Ip=Rs("Ip") Then Response.Write "<font color=ff0000>"%>
                    <%=Rs("Ip")%> 
                    <%If Ip=Rs("Ip") Then Response.Write "</font>"%>
                </td>
                <td><%=GetTurnTime(Hour(Rs("LastTime")))&":"&GetTurnTime(Minute(Rs("LastTime")))&":"&GetTurnTime(Second(Rs("LastTime")))%></td>
              </tr>
              <%rs.movenext
mypagesize=mypagesize-1
wend%>
            </form>
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
                  <a href="?Action=<%=Action%>&AddDate=<%=AddDate%>&Px=<%=Px%>&Page=<%=totalpage%>">最后一页</a> 
                  页次：<font color="#ff0000"><%=page%></font>/<%=totalpage%>页&nbsp;&nbsp;共<%=totalrs%>条记录&nbsp;&nbsp;每页显示<%=rs.pagesize%>条
<%
Response.write "&nbsp;&nbsp;转到第<select id='page' onChange=""window.location=document.getElementById('page').options[document.getElementById('page').selectedIndex].value"">"
For I=1 To TotalPage
    Response.Write "<option value=?Action="&Action&"&AddDate="&AddDate&"&Px="&Px&"&Page="&i
 If Page=I Then Response.Write " selected"
 Response.Write ">"& I &"</option>"
Next
Response.write "</select>页"
%></td>

            </tr>
          </table>
          <%End If%>
          <%If Action="sysset" Then%>
<SCRIPT>

function show_1(){
t_2_1.style.display = "none";
t_2_2.style.display = "";
t_2_3.style.display = "none";
}
function show_2(){
t_2_1.style.display = "";
t_2_2.style.display = "";
t_2_3.style.display = "";
}
</SCRIPT>
<%
Set Rs=Server.CreateObject("Adodb.RecordSet")
Sql="Select * From CFCount_Admin"
Rs.Open Sql,Conn,1,1%>
          
<table >


  <form name="form2" method="post" action="?action=syssetmodifysave">
  <tr > 
      <td colspan="2">系统设置</td>
    </tr>
    <tr> 
      <td width="22%">计数器系统名称：</td>
      <td><input name="SysTitle" type="text" id="mc" value="<%=Rs("SysTitle")%>" size="30"></td>
    </tr>
    <tr> 
      <td>计数器简称：</td>
      <td><input name="TjTextName" type="text" id="mc" value="<%=Rs("TjTextName")%>" size="30">
        (8个字以内)</td>
    </tr>
    <tr> 
      <td>网友是否能申请计数器：</td>
      <td> <input type="radio" name="RegState" value="-1" <%If Rs("RegState")=-1 Then%>checked<%End If%>>
        是 　　 
        <input name="RegState" type="radio" value="0" <%If Rs("RegState")=0 Then%>checked<%End If%>>
        否<br> &nbsp;注：如果计数器只想自己用，请注册用户后把申请功能关闭。<br> <%If Rs("RegState")=0 Then%> <font color="#FF0000">提示：</font>目前已经禁止申请计数器！ 
        <%End If%> </td>
    </tr>
    <tr> 
      <td>是否记录来源页面,IP信息：</td>
      <td><input type="radio" name="LyKeep" value="-1" <%If Rs("LyKeep")=-1 Then%>checked<%End If%>>
        是 　　 
        <input name="LyKeep" type="radio" value="0" <%If Rs("LyKeep")=0 Then%>checked<%End If%>>
        否<br> <%if rs("LyKeep")=0 then%> <font color="#FF0000">提示：</font>目前已经禁止记录IP，来源信息！ 
        <%end if%> </td>
    </tr>
    <tr> 
      <td>是否记录在线人数：</td>
      <td><input type="radio" name="OnlineKeep" value="-1" <%if rs("OnlineKeep")=-1 then%>checked<%end if%>>
        是 　　 
        <input name="OnlineKeep" type="radio" value="0" <%if rs("OnlineKeep")=0 then%>checked<%end if%>>
        否<br> <%if rs("OnlineKeep")=0 then%> <font color="#FF0000">提示：</font>目前已经禁止记录在线人数信息！ 
        <%end if%> </td>
    </tr>
    <tr> 
      <td>是否记录网站页面浏览次数：</td>
      <td><input type="radio" name="WebKeep" value="-1" <%if rs("WebKeep")=-1 then%>checked<%end if%>>
        是 　　 
        <input name="WebKeep" type="radio" value="0" <%if rs("WebKeep")=0 then%>checked<%end if%>>
        否<br> <%if rs("WebKeep")=0 then%> <font color="#FF0000">提示：</font>目前已经禁止记录网站页面浏览次数！ 
        <%end if%> </td>
    </tr>
    <tr> 
      <td>计数图片样式数量：</td>
      <td> <input name="StyleTotal" type="text" id="stylenum" value="<%=rs("StyleTotal")%>" size="30">
        个</td>
    </tr>
    <tr> 
      <td>计数器Logo图片地址：</td>
      <td><input name="LogoUrl" type="text" id="LogoUrl" value="<%=rs("LogoUrl")%>" size="30">
        大小480*60</td>
    </tr>
    <tr> 
      <td>显示IP地址的IP地域：</td>
      <td class="td_1"><input type="radio" name="IpArea" value="-1"<%if rs("IpArea")=-1 then Response.write " checked"%>>
        是 
        <input type="radio" name="IpArea" value="0"<%if rs("IpArea")=0 then Response.write " checked"%>>
        否<br>
        注：由于IP数据库较大,默认为空,需要显示具体IP地区位置的话请上js.tb11.com/js/上下载IP数据库后,再选择此选项为&quot;是&quot;</td>
    </tr>
    <tr>
      <td>系统皮肤：</td>
      <td><input type="radio" name="skintype" value="0"<%if rs("skintype")=0 then Response.write " checked"%> onclick="document.getElementById('skincolor').value='CACACA|F3F9FE';">
        默认 
        <input type="radio" name="skintype" value="1"<%if rs("skintype")=1 then Response.write " checked"%>  onclick="document.getElementById('skincolor').value='43A102|E4F9D5';">
        样式1 
        <input type="radio" name="skintype" value="2" <%if rs("skintype")=2 then Response.write " checked"%>  onclick="document.getElementById('skincolor').value='A73800|F6BDA0';">
        样式2 
        <input type="radio" name="skintype" value="3"<%if rs("skintype")=3 then Response.write " checked"%>  onclick="document.getElementById('skincolor').value='FFA500|FBD38B';">
        样式3 
        <input type="radio" name="skintype" value="4"<%if rs("skintype")=4 then Response.write " checked"%>  onclick="document.getElementById('skincolor').value='4C4C4C|E7E4E4';">
        样式4 </td>
    </tr>
    <tr id="a_1"> 
      <td>自定义皮肤颜色：</td>
      <td><input name="skincolor" type="text" id="skincolor" value="<%=rs("SkinColor")%>" size="30">
        (输入颜色代码，默认值为：BFE7F1)</td>
    </tr>
    <tr> 
      <td>发送邮件组件：</td>
      <td><input type="radio" name="emailsendtype" value="1"<%if rs("emailsendtype")=1 then%> checked<%end if%> onclick="show_1()">
        CDONTS组件 
        <%
on error resume next
set TestObj=server.CreateObject ("CDONTS.NewMail")
If -2147221005 <> Err then
 Response.write "(支持)"
Else
  Response.write "(不支持)"
End If
%> <input type="radio" name="emailsendtype" value="2"<%if rs("emailsendtype")=2 then%> checked<%end if%> onclick="show_2()">
        JMAIL组件 
        <%
on error resume next
set TestObj=server.CreateObject ("JMail.SmtpMail")
If -2147221005 <> Err then
 Response.write "(支持)"
Else
  Response.write "(不支持)"
End If
%></td>
    </tr>
    <tr id="t_2_1"> 
      <td>服务器地址：</td>
      <td><input name="SmtpAddress" type="text" value="<%=rs("SmtpAddress")%>" size="30">
      (如smtp.a.com)</td>
    </tr>
    <tr id="t_2_2"> 
      <td>发送邮件帐号名称：</td>
      <td><input name="SmtpUser" type="text" value="<%=rs("SmtpUser")%>" size="30">
      (如a@a.com)</td>
    </tr>
    <tr id="t_2_3"> 
      <td>发送邮件帐号的密码：</td>
      <td><input name="SmtpPassword" type="password" value="<%=rs("SmtpPassword")%>" size="30">
      <br>
      如：你有一个邮件b@a.com，密码为c，发信服务器地址设置设置为：<br>smtp.a.com发送邮件帐号名称设置为b@a.com，发送邮件帐号的密<br>码设置为c，其它邮件类似</td>
    </tr>
    <tr> 
      <td></td>  <td>  
          <input type="submit" name="Submit32" value="修改">
          　　 
          <input type="reset" name="Submit5232" value="取消">
      </td>
    </tr>
  </form>
</table>
          <%End If%>
		  
<%If Action="otherset" Then%>

          
<table >

  <form name="form2" method="post" action="?action=othersetmodifysave">
  <tr > 
      <td colspan="2">其它设置</td>
    </tr>
    <tr> 
      <td>系统安全码：</td>
      <td><input name="SysCode" type="text" id="SysCode" value="<%=GetMySet("SysCode")%>" size="25">
        *务必在这里设置8个以上任意字符加强安全</td>
    </tr>
    <tr> 
      <td>小时统计保留天数：</td>
      <td><input name="HourCountKeepDay" type="text" id="HourCountKeepDay" value="<%=GetMySet("HourCountKeepDay")%>" size="25">
        *默认2天</td>
    </tr>
    <tr> 
      <td>网页浏览记录保留天数：</td>
      <td><input name="WebKeepDay" type="text" id="WebKeepDay" value="<%=GetMySet("WebKeepDay")%>" size="25">
        *默认2天</td>
    </tr>	
    <tr> 
      <td>来源记录保留天数：</td>
      <td><input name="LyKeepDay" type="text" id="LyKeepDay" value="<%=GetMySet("LyKeepDay")%>" size="25">
        *默认2天</td>
    </tr>
    <tr> 
      <td>搜索引擎记录保留天数：</td>
      <td><input name="SearchKeepDay" type="text" id="SearchKeepDay" value="<%=GetMySet("SearchKeepDay")%>" size="25">
        *默认7天</td>
    </tr>
    <tr> 
      <td>搜索关键字记录保留天数：</td>
      <td><input name="KeywordKeepDay" type="text" id="KeywordKeepDay" value="<%=GetMySet("KeywordKeepDay")%>" size="25">
        *默认2天</td>
    </tr>
    <tr> 
      <td>天数记录保留天数：</td>
      <td><input name="CountKeepDay" type="text" id="CountKeepDay" value="<%=GetMySet("CountKeepDay")%>" size="25">
        *默认60天</td>
    </tr>
	<tr> 
      <td></td>  <td>  
          <input type="submit" name="Submit32" value="修改">
          　　 
          <input type="reset" name="Submit5232" value="取消">
      </td>
    </tr>
  </form>
</table>
          <%End If%>
          <%If Action="searchset" Then%>
          <%

Sql="Select * From CFCount_Search_Set Order By ID Desc"
Set Rs=Server.CreateObject("Adodb.RecordSet")
Rs.open Sql,Conn,1,1
%>
<table >
  <tr > 
              <td>序号</td>
              <td>搜索引擎中文名</td>
              <td>搜索引擎英文标识</td>
              <td>搜索关键字的参数</td>
              <td>操作</td>
  </tr>
            <%While Not Rs.Eof
			I=I+1%>
            <tr class="tr_2"> 
              <td ><%=I%></td>
              <td ><a href="http://www.<%=Rs("SiteFlag")%>" target=_blank><%=Rs("SiteDesc")%></a></td>
              <td ><a href="http://www.<%=Rs("SiteFlag")%>" target=_blank><%=Rs("SiteFlag")%></a></td>
              <td ><%=Rs("KeywordFlag")%></td>
              <td > 
                  <a href="?Action=searchmodify&ID=<%=Rs("ID")%>">修改</a> 
                  <a href="?Action=searchdel&SiteFlag=<%=Rs("SiteFlag")%>" onClick="{if(confirm('确定要删除么?')){return true;}return false;}">删除</a> 
              </td>
            </tr>
            <%
AllSearch=AllSearch&Rs("SiteFlag")&","&Rs("KeywordFlag")&"|"
rs.movenext
wend

AllSearch = StrReverse(Mid(StrReverse(AllSearch), 2))

Sql="Update CFCount_Admin Set AllSearch='"&AllSearch&"'"
Conn.ExeCute Sql
%>
</table>

<table >
  <tr > 
              <td colspan="4">增加自定义引擎</td>
  </tr>
            <form name="form5" method="post" action="?Action=searchaddsave">
              <tr class="tr_2"> 
                <td>搜索引擎中文名</td>
                <td width="191">搜索引擎英文标识</td>
                <td width="191">搜索关键字的参数</td>
                <td></td>
              </tr>
              <tr class="tr_2"> 
                <td>  
                    <input name="SiteDesc" type="text" id="SiteDesc">
                </td>
                <td> 
                    <input name="SiteFlag" type="text" id="SiteFlag">
                </td>
                <td> 
                    <input name="KeywordFlag" type="text" id="KeywordFlag">
                </td>
                <td><input type="submit" name="Submit" value="增　加"></td>
              </tr>
              <tr> 
                <td colspan="4"  class="td_1">比如：<br>
        百度网页搜索&quot;abcd&quot;IE地址栏为:http://www.baidu.com/s?wd=abcd&amp;cl=3<br>
        百度贴吧搜索&quot;abcd&quot;IE地址栏为:http://post.baidu.com/f?kw=abcd<br>
        百度知道搜索&quot;abcd&quot;IE地址栏为:http://zhidao.baidu.com/q?ct=17&amp;pn=0&amp;tn=ikaslist&amp;rn=10&amp;word=abcd&amp;t=51<br>
        这要我们就可以这样填写：搜索引擎中文名填写&quot;百度&quot;,搜索引擎英文标识填写&quot;baidu.com&quot;，标识可以填写多个，各个标识间用英文分号;分隔，搜索关键字的参数填写&quot;wd;kw;word&quot;,关键字参数用英文分号;分隔，以此类推</td>
              </tr>
            </form>
</table>
          <%End If%>
   
          <%If Action="searchmodify" Then%>
          <%
ID=ChkStr(Request("ID"),2)
Set Rs=Server.CreateObject("Adodb.RecordSet")
Sql="Select * From CFCount_Search_Set Where ID="&ID
Rs.Open Sql,Conn,1,1%>
<table >
              <tr > 
              <td colspan="4">修改搜索引擎</td>
  </tr>
            <form name="form5" method="post" action="?Action=searchmodifysave&ID=<%=ID%>">
              <tr class="tr_2"> 
                <td>搜索引擎中文名</td>
                <td width="191">搜索引擎英文标识</td>
                <td width="191">搜索关键字的参数</td>
                <td></td>
              </tr>
              <tr class="tr_2"> 
                <td> 
                    <input name="SiteDesc" type="text" id="SiteDesc" value="<%=Rs("SiteDesc")%>">
                </td>
                <td>
                    <input name="SiteFlag" type="text" id="SiteFlag" value="<%=Rs("SiteFlag")%>">
                </td>
                <td> 
                    <input name="KeywordFlag" type="text" id="KeywordFlag" value="<%=Rs("KeywordFlag")%>">
                </td>
                <td><input type="submit" name="Submit" value="修　改"></td>
              </tr>
              <tr> 
                <td colspan="4" class="td_1">如何增加或修改例子，例如你在百度上搜索网页&quot;康佑&quot;,在结果页的地址栏里的网址是http://www.baidu.com/s?wd=%B3%CB%B7%E7&amp;cl=3<br>
                  这样我们就这样填写：搜索引擎中文名填写&quot;百度&quot;,搜索引擎英文标识填写&quot;baidu.com&quot;,搜索关键字的参数填写&quot;wd&quot;</td>
              </tr>
            </form>
</table>
          <%End If%>
		  
		  
<%if Action="imglist" then%>
<table >
  <tr > 
 <td>网店类图片管理</td>
</tr>
<tr> 
 <td><iframe style="top:2px" ID="UploadFiles" src="CF_Upfile.asp?UserType=0&CheckCode=<%=RsSet("AdminCode")%>" frameborder=0 scrolling=no width="450" height="25"></iframe></td>
</tr>


  <tr> 
    <td>

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
 

Response.write "文件大小：" & CInt(Rs("FileSizeNum")/1024) &"k<br>"

Response.write "使用人数："
Sql="Select Count(*) From CFCount_User Where ImgFileName='"&Rs("FileName")&"'"
Set Rs2=Conn.Execute(Sql)
Response.write Rs2(0)
Rs2.Close()
Response.write "<br>"

Response.write "上传时间："&Rs("AddTime")&"<br>"

Response.write "操作：<a href='?Action=picdel&FileName="&Rs("FileName")&"'onClick=""{if(confirm('确定要删除么?')){return true;}return false;};"">删除</a>"
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
        </table>
		
    </td>
  </tr>

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

<%end if%>


<%If Action="pwdmodify" Then%>
<%
Set Rs = Server.CreateObject("ADODB.Recordset")
Sql="Select * From CFCount_Admin"
Rs.open Sql,Conn,1,1
%>
  
<table >

  
  <form name="form2" method="post" action="?Action=pwdmodifysave">
  <tr > 
      <td colspan="2">输入原管理员密码</td>
    </tr>
    <tr>
      <td width="16%">原管理员用户名：</td>
      <td><%=Rs("Admin")%></td>
    </tr>
    <tr> 
      <td>请输入原管理密码：</td>
      <td><input name="Pwd_Old" type="password" id="Pwd_Old" value="" size="15"></td>
    </tr>
  <tr > 
      <td colspan="2">新管理员名称和密码</td>
    </tr>
    <tr> 
      <td>新管理员名称：</td>
      <td><input name="Admin" type="text" id="admin" value="<%=Rs("Admin")%>" size="15"></td>
    </tr>
    <tr> 
      <td>新密码：</td>
      <td><input name="Pwd" type="password" id="Pwd" size="15"></td>
    </tr>
    <tr> 
      <td>新密码确认： </td>
      <td><input name="Pwd2" type="password" id="Pwd2" size="15"></td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><input type="submit" name="Submit323" value="修改">
        　　 
        <input type="reset" name="Submit5233" value="取消"> </td>
    </tr>
  </form>
</table>
<%End If%>


<%If Action="dbdo" Then%>         
  
<table >
  <tr > 
    <td>执行Sql语句</td>
  </tr>
  <form name="form2" method="post" action="?action=sqlexesave">
    <tr> 
      <td><textarea name="sql" cols="80" rows="5" id="sqlcmd"></textarea> 
      </td>
    </tr>
    <tr>
      <td>请输入管理员密码确认执行：
        <input type="password" name="Pwd"></td>
    </tr>
    <tr>
      <td><input type="submit" name="Submit" value="确定执行以上Sql语句" onclick="{if(confirm('确定执行以上Sql语句么?')){return true;}return false;};">
        注：多行语句执行以回车换行，每一行要求为一完整语句</td>
    </tr>
  </form>
  <tr >  
    <td >数据库日志压缩</td>
  </tr>
  <tr> 
    <td><input type="button" name="Submit2" value="点击压缩数据库" onclick="{if(confirm('确定要压缩数据库么?')){window.location='?Action=dbys';return true;}return false;};">
      注：任何时候可以操作，但每次操作不要间隔太短，如一分钟内操作好几次</td>
  </tr>
  <tr > 
    <td> 数据库备份</td>
  </tr>
  <tr> 
    <td><input type="button" name="Submit2" value="点击备份数据库" onclick="{if(confirm('确定要备份数据库么?此操作可能需要花费比较长的时间')){window.location='?Action=dbbackup';return true;}return false;};">
      注：条件是程序和数据库放在一台服务器上，并且不是在虚拟主机上，放在虚拟主机上请手工自己备份,每次操作不要间隔太短，如一分钟内操作好几次，备份数据库放在程序的data目录里，以backup开头后面是备份的时间的文件格式命名，数据库大的话可此操作需要花费比较长的时间</td>
  </tr>
  <tr > 
    <td  class="td_1"><%
foldername=server.mappath("data")
path=foldername
If Instr(path,"\data")=0 Then Call AlertBack("不能浏览这个目录！",1)
ShowFolderList(path) 



Function ShowFolderList(folderspec) 
temp=request.ServerVariables("HTTP_REFERER") 
temp=left(temp,Instrrev(temp,"/")) 
temp1=len(folderspec)-len(server.MapPath("./"))-1 
if temp1>0 then 
temp1=right(folderspec,cint(temp1)) 
elseif temp1=-1 then 
temp1="" 
end if 
tempurl=temp+replace(temp1,"\","/")+"/" 
Set fso = CreateObject("Scripting.FileSystemObject") 
upfolderspec=fso.GetParentfoldername(folderspec&"\") 
%> <table >
        <tr> 
          <td>名称</td>
          <td>大小</td>
          <td>类型</td>
          <td>修改时间</td>
        </tr>
        <% 
'列出目录 
Set f = fso.GetFolder(folderspec) 
Set fc = f.SubFolders 
For Each f1 in fc 
%>
        <tr> 
          <td><%= f1.name%></td>
          <td>
              <%MySize=f1.size/1024
	  If MySize<1024 Then Response.write MySize&"K"
	  If MySize>=1024 Then Response.write MySize/1024&"M"%>
          </td>
          <td>文件夹</td>
          <td><%= f1.datelastmodified%></td>
        </tr>
        <% 
Next 
foldername=StrReverse(left(StrReverse(folderspec),instr(StrReverse(folderspec),"\")-1))&"/"
'列出文件 
Set fc = f.Files 
For Each f1 in fc 
%>
        <tr> 
          <td><%= f1.name%>&nbsp;&nbsp; <%If Instr(Lcase(f1.name),"backup")>0 Then%>
            [<a href="?Action=dbdel&DbMc=<%= f1.name%>" onClick="{if(confirm('确定要删除么?')){return true;}return false;}">删除</a>] 
            <%End If%></td>
          <td><%MySize=f1.size/1024
	  If MySize<1024 Then Response.write MySize&"K"
	  If MySize>=1024 Then Response.write MySize/1024&"M"%>
          </td>
          <td><div >&nbsp;</td>
          <td><%= f1.datelastmodified%></td>
        </tr>
        <% 
Next 
set fso=nothing 
%>
      </table>
      <%End Function %></td>
  </tr>
</table>
<%End If%>