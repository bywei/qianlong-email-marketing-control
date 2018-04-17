
<%
UserName=ChkStr(Request.QueryString("UserName"),1)
Ip=Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If Ip= "" Then Ip= Request.ServerVariables("REMOTE_ADDR")

Set RsUser= Server.CreateObject("Adodb.RecordSet")
Sql="Select * From CFCount_User Where UserName='"&UserName&"'"
RsUser.Open Sql,Conn,1,1
%>


<table class="tb_2">
  <tr class="tr_1"> 
    <td><a href="Index.asp?UserName=<%=Username%>">[<%=RsUser("PageName")%>] 的站长请点击这里登录</a></td>
  </tr>
  <form name="form1" method="post" action="?Action=viewuserlogin&UserName=<%=UserName%>">
    <tr> 
      <td class="td_2">独立查看密码： 
          <input name="Pwd_View" type="password" id="Pwd_View">
          <input type="submit" name="Submit" value="查看统计">
        </td>
    </tr>
  </form>
</table>

<%If RsUser("GbookState")=-1 Then%>
<script>
function gbookcheck(){
 if(document.gbook_form.content.value==""||document.gbook_form.content.value=="请输入留言内容"){
  alert("请输入留言内容!");
  document.gbook_form.content.focus();
  return false;
 }
 if(document.gbook_form.contact.value==""||document.gbook_form.contact.value=="请输入联系方式"){
  alert("请输入联系方式!");
  document.gbook_form.contact.focus();
  return false;
 }
 
document.gbook_form.submit();
}
</script>
<table class="tb_2 tb_underline">
  <form name="gbook_form" method="post" action="Gbook.asp?Action=gbookaddsave&UserName=<%=UserName%>" target="_blank">
    <tr class="tr_1"> 
      <td>给[<%=RsUser("PageName")%>]站长留言</td>
    </tr>
	<tr class="tr_2"> 
      <td>留言内容：<br /> 
<textarea name='content' id='content' cols='22' rows='8' style="overflow:auto;border: 1px solid #CACACA;" onmouseover="if(document.getElementById('content').value=='请输入留言内容')document.getElementById('content').value='';" onmouseout="if(document.getElementById('content').value=='')document.getElementById('content').value='请输入留言内容';">请输入留言内容</textarea>
        </td>
		</tr>
		<tr class="tr_2"> 
      <td>联系方式：<br /> 
<textarea name='contact' id='contact' cols='22' rows='6' style="overflow:auto;border: 1px solid #CACACA;" onmouseover="if(document.getElementById('contact').value=='请输入联系方式')document.getElementById('contact').value='';" onmouseout="if(document.getElementById('contact').value=='')document.getElementById('contact').value='请输入联系方式';">请输入联系方式</textarea>
        </td>
    </tr>
		<tr class="tr_2"> 
      <td><span style="cursor:hand"><img src='images/gbook_send.gif' onclick="gbookcheck()" ></span>
        </td>
    </tr>
  </form>
</table>
<%End If%> 

<%If RsUser("TjOpen")=-1 Then'统计开放时%>


<%If RsUser("SurveyOpen")=-1 Then%>
<%
Sql="Select Top 1 MyCounter,AddDate From CFCount_Count_Day Where UserName='"&UserName&"' Order By MyCounter Desc"
Set RsTopShow=Conn.Execute(Sql)

Sql="Select Top 1 IpCounter,AddDate From CFCount_Count_Day Where UserName='"&UserName&"' Order By IpCounter Desc"
Set RsTopIp=Conn.Execute(Sql)

Sql="Select Top 1 MyCounter,AddDate From CFCount_Count_Day Where UserName='"&UserName&"' And DateDiff('d','"&RsUser("AddTime")&"',AddDate)>0 And DateDiff('d',AddDate,Date())>0 Order By MyCounter"
Set RsLowShow=Conn.Execute(Sql)

Sql="Select Top 1 IpCounter,AddDate From CFCount_Count_Day Where UserName='"&UserName&"' And DateDiff('d','"&RsUser("AddTime")&"',AddDate)>0 And DateDiff('d',AddDate,Date())>0 Order By IpCounter"
Set RsLowIp=Conn.Execute(Sql)

%>
<table class="tb_2 tb_underline">
  <tr class="tr_1"> 
            <td colspan="4"> 统计概况</td>
          </tr>
          <tr > 
            <td width="20%" >网站名称：</td>
            <td colspan="3" ><b><%=RsUser("PageName")%></b></td>
          </tr>
          <tr > 
            <td >站点地址：</td>
            <td colspan="3" ><a href="<%=RsUser("PageUrl")%>" target="_blank"><%=RsUser("PageUrl")%></a></td>
          </tr>
          <tr > 
            <td >开始统计于：</td>
            <td colspan="3" ><%=RsUser("AddTime")%></td>
          </tr>
          <tr > 
            <td >已统计天数：</td>
            <td colspan="3" ><%=DateDiff("d",RsUser("AddTime"),Now())%>天</td>
          </tr>
          <tr > 
            <td >在线人数：</td>
            <td colspan="3" ><%
If IsEmpty(Application("CFCountOnline_"&UserName)) Then
 OnlineTotal=0
Else
 Myarray=Split(Application("CFCountOnline_"&UserName),"|")
 OnlineTotal=Ubound(Myarray)
End If
Response.Write OnlineTotal&"人"%></td>
          </tr>
          <tr > 
            <td ></td>
            <td width="62" >浏览量 
              </td>
            <td width="84" >访问量 
              </td>
            <td width="407" ></td>
          </tr>
          <tr > 
            <td >总量：</td>
            <td ><%=RsUser("RealShowTotal")%></td>
            <td ><%=RsUser("RealIpTotal")%></td>
            <td ></td>
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
            <td >最高访问量： <%=RsTopIp("IpCounter")%> IP (发生在: <%=RsTopIp("AddDate")%>) </td>
          </tr>
          <%
Sql="Select MyCounter,IpCounter From CFCount_Count_Day Where UserName='"&UserName&"' And DateDiff('d',AddDate,Date())=1"
Set Rs=Conn.Execute(Sql)
%>
          <tr > 
            <td >昨日流量：</td>
            <td ><%=Rs("MyCounter")%></td>
            <td ><%=Rs("IpCounter")%></td>
            <td >最高浏览量： <%=RsTopShow("MyCounter")%> PV (发生在: <%=RsTopShow("AddDate")%>)</td>
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
Set Rs2=Conn.Execute(Sql)
%>
          <tr > 
            <td >今年累计：</td>
            <td ><%=Rs2(0)%></td>
            <td ><%=Rs2(1)%></td>
            <td >最低访问量： <%=RsLowIp("IpCounter")%> IP (发生在: <%=RsLowIp("AddDate")%>)</td>
          </tr>
          <tr > 
            <td >平均每日：</td>
            <td ><%=Int(RsUser("RealShowTotal")/DateDiff("d",RsUser("AddTime"),Now())+1)%></td>
            <td ><%=Int(RsUser("RealIpTotal")/DateDiff("d",RsUser("AddTime"),Now())+1)%></td>
            <td >最低浏览量： <%=RsLowShow("MyCounter")%> Pv (发生在: <%=RsLowShow("AddDate")%>)</td>
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
<%End If%>

<%If RsUser("TodayLyOpen")=-1 Then%>
<%
Myarray=Split(Application("CFCountLy_"&UserName),"|")
%>

<table class="tb_2 tb_underline">
  <tr class="tr_1"> 
            <td colspan="8" >最新访问记录</td>
          </tr>
          <tr> 
            <td > <p >序号</td>
            <td >IP</td>
            <td >访问次数</td>
            <td >访问时间</td>
            <td >来源页面</td>
            <td >最后访问时间</td>
            <td >最后浏览页面</td>
            <td >共浏览几页</td>
          </tr>
          <%
		  If Ubound(Myarray)>10 Then
		   K=Ubound(Myarray)-10
		  Else
		   K=0
		  End if
		  
		  For I=Ubound(Myarray)-1 TO K Step -1
		  J=J+1%>
          <%
		  If MyArray(I)="" Then MyArray(I)="\"
		  MyArray_2=Split(MyArray(I),"\")%>
          <tr> 
            <td ><%=J%></td>
            <td > <%If Ip=MyArray_2(0) Then Response.Write "<font color=ff0000>"%> <%=MyArray_2(0)%> <%If Ip=MyArray_2(0) Then Response.Write "</font>"%> <%If RsSet("IpArea")=-1 Then
				  Response.write "<br>"&GetIpArea(MyArray_2(0))
				 End If%> </td>
            <td ><%=MyArray_2(1)%></td>
            <td ><%=GetTurnTime(Hour(Cdate(MyArray_2(2))))&":"&GetTurnTime(Minute(Cdate(MyArray_2(2))))&":"&GetTurnTime(Second(Cdate(MyArray_2(2))))%></td>
            <td > <%If MyArray_2(3)="" Then%>
              直接从浏览器输入 
              <%Else%> <a href="http://<%=BreakUrl(MyArray_2(3),1)%>" target='_blank'><%=BreakUrl(MyArray_2(3),1)%></a>[<a href="<%=MyArray_2(3)%>" target='_blank'>详细</a>] 
              <%End If%></td>
            <td ><%=GetTurnTime(Hour(Cdate(MyArray_2(4))))&":"&GetTurnTime(Minute(Cdate(MyArray_2(4))))&":"&GetTurnTime(Second(Cdate(MyArray_2(4))))%></td>
            <td > <%If MyArray_2(5)="" Then%>
              直接从浏览器输入 
              <%Else%> <a href="<%=MyArray_2(5)%>" target='_blank'><%=BreakUrl(MyArray_2(5),2)%></a>[<a href="<%=MyArray_2(5)%>" target='_blank'>详细</a>] 
              <%End If%></td>
            <td ><%=MyArray_2(6)%></td>
          </tr>
          <%
Next%>
        </table>
        <br>

<table class="tb_2 tb_underline">


          <%
Ip=Request.ServerVariables("HTTP_X_FORWARDED_FOR")
If Ip= "" Then Ip= Request.ServerVariables("REMOTE_ADDR")

Sql="Select Sum(MyCounter) From CFCount_Ly_Day where UserName='"&UserName&"' And Datediff('d',AddDate,Date())=0"
Set Rs=Conn.execute(Sql)
ThisDayTotal=Rs(0)
If ThisDayTotal=0 Then ThisDayTotal=1

Sql="Select * From CFCount_Ly_Day where UserName='"&UserName&"' And Datediff('d',AddDate,Date())=0 Order By MyCounter Desc"
Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.open Sql,Conn,1,1
%>
  <tr class="tr_1"> 
            <td colspan="7"><%=Date()%> 来源统计</td>
          </tr>
          <tr> 
            <td >序号</td>
            <td >来源网站</td>
            <td >来源数量</td>
            <td >所占此天的比例</td>
            <td >上站时间</td>
            <td >最后IP地址</td>
            <td >最后访问时间</td>
          </tr>
<%
I=0
While Not Rs.Eof
I=I+1
%>
          <tr> 
            <td class="td_1"><%=I%></td>
            <td class="td_1"> 
                <%If Rs("LyHead")="" Then
  Response.Write "直接从浏览器输入"
 Else
  Response.Write "<a href=http://"&Rs("LyHead")&" target='_blank'>"&Rs("LyHead")&"</a>&nbsp;[<a href="&Rs("Ly")&" target='_blank'>详细</a>]"
 End If%>
              </td>
            <td class="td_1"><%=Rs("MyCounter")%></td>
            <td class="td_1"><img src=images/BlueBar.gif width=<%=Rs("MyCounter")/ThisDayTotal*150%> height=10><%=MyRate(Rs("MyCounter"),ThisDayTotal)%>%</td>
            <td class="td_1"><%=GetTurnTime(Hour(Rs("AddTime")))&":"&GetTurnTime(Minute(Rs("AddTime")))&":"&GetTurnTime(Second(Rs("AddTime")))%></td>
            <td class="td_1"> 
                <%If Ip=Rs("Ip") Then Response.Write "<font color=ff0000>"%>
                <%=Rs("Ip")%> 
                <%If Ip=Rs("Ip") Then Response.Write "</font>"%>

              </td>
            <td class="td_1"><%=GetTurnTime(Hour(Rs("LastTime")))&":"&GetTurnTime(Minute(Rs("LastTime")))&":"&GetTurnTime(Second(Rs("LastTime")))%></td>
          </tr>
          <%rs.movenext
wend%></form>
        </table>

<%End If%>
        <%If RsUser("OnlineOpen")=-1 Then%><br>

        <%
Myarray=Split(Application("CFCountOnline_"&UserName),"|")
%>

<table class="tb_2 tb_underline">
  <tr class="tr_1"> 
            <td colspan="8" >最新在线</td>
          </tr>
          <tr> 
            <td > <p >序号</td>
            <td >IP</td>
            <td >访问次数</td>
            <td >访问时间</td>
            <td >来源页面</td>
            <td >最后访问时间</td>
            <td >最后浏览页面</td>
            <td >共浏览几页</td>
          </tr>
          <%
		  If Ubound(Myarray)>10 Then
		   K=Ubound(Myarray)-10
		  Else
		   K=0
		  End if
		  
		  For I=Ubound(Myarray)-1 TO K Step -1
		  J=J+1%>
          <%
		  If MyArray(I)="" Then MyArray(I)="\"
		  MyArray_2=Split(MyArray(I),"\")%>
          <tr> 
            <td ><%=J%></td>
            <td > <%If Ip=MyArray_2(0) Then Response.Write "<font color=ff0000>"%> <%=MyArray_2(0)%> <%If Ip=MyArray_2(0) Then Response.Write "</font>"%> <%If RsSet("IpArea")=-1 Then
				  Response.write "<br>"&GetIpArea(MyArray_2(0))
				 End If%>
</td>
            <td ><%=MyArray_2(1)%></td>
            <td ><%=GetTurnTime(Hour(Cdate(MyArray_2(2))))&":"&GetTurnTime(Minute(Cdate(MyArray_2(2))))&":"&GetTurnTime(Second(Cdate(MyArray_2(2))))%></td>
            <td > <%If MyArray_2(3)="" Then%>
              直接从浏览器输入 
              <%Else%> <a href="http://<%=BreakUrl(MyArray_2(3),1)%>" target='_blank'><%=BreakUrl(MyArray_2(3),1)%></a>[<a href="<%=MyArray_2(3)%>" target='_blank'>详细</a>] 
              <%End If%></td>
            <td ><%=GetTurnTime(Hour(Cdate(MyArray_2(4))))&":"&GetTurnTime(Minute(Cdate(MyArray_2(4))))&":"&GetTurnTime(Second(Cdate(MyArray_2(4))))%></td>
            <td > <%If MyArray_2(5)="" Then%>
              直接从浏览器输入 
              <%Else%> <a href="<%=MyArray_2(5)%>" target='_blank'><%=BreakUrl(MyArray_2(5),2)%></a>[<a href="<%=MyArray_2(5)%>" target='_blank'>详细</a>] 
              <%End If%></td>
            <td ><%=MyArray_2(6)%></td>
          </tr>
          <%
Next%></table>
        <%End If%>
		
        <%If RsUser("TodayHourOpen")=-1 Then%>
  <%
AddDate=Date()

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

<table class="tb_2 tb_underline">
  <tr class="tr_1"> 
            <td colspan="4"><%=AddDate%> 每小时报表</td>
          </tr>
          <tr> 
            <td>小时</td>
            <td width="134">显示数量</td>
            <td width="135">IP数量</td>
            <td>比率</td>
          </tr>
          <%
do while not rs.eof
I=I+1

Sql="Select MyCounter,IpCounter From CFCount_Count_Hour where UserName='"&UserName&"' And AddDate=#"&AddDate&"# And AddHour="&Rs("AddHour")
Set Rs1=Conn.Execute(Sql)
%>
          <tr> 
            <td><%=Rs("AddHour")%>-<%=Rs("AddHour")+1%></td>
            <td><%=Rs1(0)%><br>
              </td>
            <td><%=Rs1(1)%></td>
            <td><img src=images/BlueBar.gif width=<%=Rs1(0)/ShowTotal*350%> height=10><%=MyRate(Rs1(0),ShowTotal)%>%<br><img src=images/GreenBar.gif width=<%=Rs1(1)/IpTotal*350%> height=10><%=MyRate(Rs1(1),IpTotal)%>%</td>
          </tr>
          <%
rs.movenext
loop%>
          <tr> 
            <td colspan="6" >图例： <img src=images/BlueBar.gif width=30 height=10>显示数，<img src=images/GreenBar.gif width=30 height=10>IP数</td>
          </tr>
		  
          <tr> 
            <td colspan="4"><%=AddDate%>此日期：总计浏览<%=ShowTotal%>次 
                IP<%=IpTotal%>个</td>
          </tr>
        </table>
        <%End If%>
        <%If RsUser("TodayHourOpen")=-1 Then%>

<table class="tb_2 tb_underline">
  <tr class="tr_1"> 
            <td colspan="4">日访问统计
<%
Action=Request("Action")
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
<%
Sql="Select Sum(MyCounter),Sum(IpCounter) From CFCount_Count_Day Where UserName='"&UserName&"'"
Set Rs1=Conn.Execute(Sql)
ShowTotal=Rs1(0)
IpTotal=Rs1(1)
If ShowTotal=0 Then ShowTotal=1
If IpTotal=0 Then IpTotal=1
%>
                <br>
              </td>
          </tr>
          <tr> 
            <td>日期</td>
            <td>显示数量</td>
            <td>IP数量</td>
            <td>比率</td>
          </tr>
          <%While Not Rs.Eof And MyPageSize>0%>

          <tr> 
            <td><%=Rs("AddDate")%></td>
            <td><%=Rs("MyCounter")%><br>
              </td>
            <td><%=Rs("IpCounter")%></td>
            <td><img src=images/BlueBar.gif width=<%=Rs("MyCounter")/ShowTotal*300%> height=10><%=MyRate(Rs("MyCounter"),ShowTotal)%>%<br>
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
              共有<%=TotalRs%>天记录，共<%=Rs1(0)%>次浏览，<%=Rs1(1)%>个IP，平均每天<%=AvgVisit%>次访问，平均每天IP<%=AvgIp%>个</td>
          </tr>
          <tr>
            <td colspan="4">
<%
Sql="Select Top 1 IpCounter,AddDate From CFCount_Count_Day Where UserName='"&UserName&"' Order By IpCounter Desc"
Set Rs1=Conn.ExeCute(Sql)
BestHightIp=Rs1("IpCounter")
BestHightIpDate=Rs1("AddDate")

Sql="Select Top 1 IpCounter,AddDate From CFCount_Count_Day Where UserName='"&UserName&"' Order By IpCounter"
Set Rs1=Conn.ExeCute(Sql)
BestLowIp=Rs1("IpCounter")
BestLowIpDate=Rs1("AddDate")

Sql="Select Top 1 MyCounter,AddDate From CFCount_Count_Day Where UserName='"&UserName&"' And DateDiff('d',AddDate,Date())>0 Order By MyCounter Desc"
Set Rs1=Conn.ExeCute(Sql)
BestHightVisit=Rs1("MyCounter")
BestHightVisitDate=Rs1("AddDate")

Sql="Select Top 1 MyCounter,AddDate From CFCount_Count_Day Where UserName='"&UserName&"' And DateDiff('d',AddDate,Date())>0 Order By MyCounter"
Set Rs1=Conn.ExeCute(Sql)
BestLowVisit=Rs1("MyCounter")
BestLowVisitDate=Rs1("AddDate")
%></td>
          </tr>
        </table>
  <table class="tb_3">
    <tr> 
      <td><a href="?UserName=<%=UserName%>&Page=1">第一页</a> 
          <%
if page>1 then%>
          <a href='?UserName=<%=UserName%>&Page=<%=Page-1%>'>上一页</a> 
          <%
end if
%>
          <%
if page<rs.pagecount   then%>
          <a href='?UserName=<%=UserName%>&Page=<%=Page+1%>'>下一页</a> 
          <%
end if
%>
          <a href="?UserName=<%=UserName%>&Page=<%=TotalPage%>">最后一页</a> 
          页次：<font color="#ff0000"><%=Page%></font>/<%=TotalPage%>页
                
<%
Response.write "&nbsp;&nbsp;转到第<select id='page' onChange=""window.location=document.getElementById('page').options[document.getElementById('page').selectedIndex].value"">"
For I=1 To TotalPage
    Response.Write "<option value=?UserName="&UserName&"&page="&i     
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

<%End If%>
