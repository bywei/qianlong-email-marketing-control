
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
    <td><a href="Index.asp?UserName=<%=Username%>">[<%=RsUser("PageName")%>] ��վ�����������¼</a></td>
  </tr>
  <form name="form1" method="post" action="?Action=viewuserlogin&UserName=<%=UserName%>">
    <tr> 
      <td class="td_2">�����鿴���룺 
          <input name="Pwd_View" type="password" id="Pwd_View">
          <input type="submit" name="Submit" value="�鿴ͳ��">
        </td>
    </tr>
  </form>
</table>

<%If RsUser("GbookState")=-1 Then%>
<script>
function gbookcheck(){
 if(document.gbook_form.content.value==""||document.gbook_form.content.value=="��������������"){
  alert("��������������!");
  document.gbook_form.content.focus();
  return false;
 }
 if(document.gbook_form.contact.value==""||document.gbook_form.contact.value=="��������ϵ��ʽ"){
  alert("��������ϵ��ʽ!");
  document.gbook_form.contact.focus();
  return false;
 }
 
document.gbook_form.submit();
}
</script>
<table class="tb_2 tb_underline">
  <form name="gbook_form" method="post" action="Gbook.asp?Action=gbookaddsave&UserName=<%=UserName%>" target="_blank">
    <tr class="tr_1"> 
      <td>��[<%=RsUser("PageName")%>]վ������</td>
    </tr>
	<tr class="tr_2"> 
      <td>�������ݣ�<br /> 
<textarea name='content' id='content' cols='22' rows='8' style="overflow:auto;border: 1px solid #CACACA;" onmouseover="if(document.getElementById('content').value=='��������������')document.getElementById('content').value='';" onmouseout="if(document.getElementById('content').value=='')document.getElementById('content').value='��������������';">��������������</textarea>
        </td>
		</tr>
		<tr class="tr_2"> 
      <td>��ϵ��ʽ��<br /> 
<textarea name='contact' id='contact' cols='22' rows='6' style="overflow:auto;border: 1px solid #CACACA;" onmouseover="if(document.getElementById('contact').value=='��������ϵ��ʽ')document.getElementById('contact').value='';" onmouseout="if(document.getElementById('contact').value=='')document.getElementById('contact').value='��������ϵ��ʽ';">��������ϵ��ʽ</textarea>
        </td>
    </tr>
		<tr class="tr_2"> 
      <td><span style="cursor:hand"><img src='images/gbook_send.gif' onclick="gbookcheck()" ></span>
        </td>
    </tr>
  </form>
</table>
<%End If%> 

<%If RsUser("TjOpen")=-1 Then'ͳ�ƿ���ʱ%>


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
            <td colspan="4"> ͳ�Ƹſ�</td>
          </tr>
          <tr > 
            <td width="20%" >��վ���ƣ�</td>
            <td colspan="3" ><b><%=RsUser("PageName")%></b></td>
          </tr>
          <tr > 
            <td >վ���ַ��</td>
            <td colspan="3" ><a href="<%=RsUser("PageUrl")%>" target="_blank"><%=RsUser("PageUrl")%></a></td>
          </tr>
          <tr > 
            <td >��ʼͳ���ڣ�</td>
            <td colspan="3" ><%=RsUser("AddTime")%></td>
          </tr>
          <tr > 
            <td >��ͳ��������</td>
            <td colspan="3" ><%=DateDiff("d",RsUser("AddTime"),Now())%>��</td>
          </tr>
          <tr > 
            <td >����������</td>
            <td colspan="3" ><%
If IsEmpty(Application("CFCountOnline_"&UserName)) Then
 OnlineTotal=0
Else
 Myarray=Split(Application("CFCountOnline_"&UserName),"|")
 OnlineTotal=Ubound(Myarray)
End If
Response.Write OnlineTotal&"��"%></td>
          </tr>
          <tr > 
            <td ></td>
            <td width="62" >����� 
              </td>
            <td width="84" >������ 
              </td>
            <td width="407" ></td>
          </tr>
          <tr > 
            <td >������</td>
            <td ><%=RsUser("RealShowTotal")%></td>
            <td ><%=RsUser("RealIpTotal")%></td>
            <td ></td>
          </tr>
          <tr > 
            <td >����������</td>
            <%
Sql="Select MyCounter,IpCounter From CFCount_Count_Day Where UserName='"&UserName&"' And DateDiff('d',AddDate,Date())=0"
Set Rs=Conn.Execute(Sql)
TodayShow=Rs("MyCounter")
TodayIp=Rs("IpCounter")
%>
            <td ><%=Rs("MyCounter")%></td>
            <td ><%=Rs("IpCounter")%></td>
            <td >��߷������� <%=RsTopIp("IpCounter")%> IP (������: <%=RsTopIp("AddDate")%>) </td>
          </tr>
          <%
Sql="Select MyCounter,IpCounter From CFCount_Count_Day Where UserName='"&UserName&"' And DateDiff('d',AddDate,Date())=1"
Set Rs=Conn.Execute(Sql)
%>
          <tr > 
            <td >����������</td>
            <td ><%=Rs("MyCounter")%></td>
            <td ><%=Rs("IpCounter")%></td>
            <td >���������� <%=RsTopShow("MyCounter")%> PV (������: <%=RsTopShow("AddDate")%>)</td>
          </tr>
          <%
Sql="Select Sum(MyCounter),Sum(IpCounter) From CFCount_Count_Day Where UserName='"&UserName&"' And DateDiff('m',AddDate,Date())=0"
Set Rs=Conn.Execute(Sql)
%>
          <tr > 
            <td >�����ۼƣ�</td>
            <td ><%=Rs(0)%></td>
            <td ><%=Rs(1)%></td>
            <td >&nbsp;</td>
          </tr>
          <%
Sql="Select Sum(MyCounter),Sum(IpCounter) From CFCount_Count_Day Where UserName='"&UserName&"' And DateDiff('yyyy',AddDate,Date())=0"
Set Rs2=Conn.Execute(Sql)
%>
          <tr > 
            <td >�����ۼƣ�</td>
            <td ><%=Rs2(0)%></td>
            <td ><%=Rs2(1)%></td>
            <td >��ͷ������� <%=RsLowIp("IpCounter")%> IP (������: <%=RsLowIp("AddDate")%>)</td>
          </tr>
          <tr > 
            <td >ƽ��ÿ�գ�</td>
            <td ><%=Int(RsUser("RealShowTotal")/DateDiff("d",RsUser("AddTime"),Now())+1)%></td>
            <td ><%=Int(RsUser("RealIpTotal")/DateDiff("d",RsUser("AddTime"),Now())+1)%></td>
            <td >���������� <%=RsLowShow("MyCounter")%> Pv (������: <%=RsLowShow("AddDate")%>)</td>
          </tr>
          <tr > 
            <td >Ԥ�ƽ��գ�</td>
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
            <td colspan="8" >���·��ʼ�¼</td>
          </tr>
          <tr> 
            <td > <p >���</td>
            <td >IP</td>
            <td >���ʴ���</td>
            <td >����ʱ��</td>
            <td >��Դҳ��</td>
            <td >������ʱ��</td>
            <td >������ҳ��</td>
            <td >�������ҳ</td>
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
              ֱ�Ӵ���������� 
              <%Else%> <a href="http://<%=BreakUrl(MyArray_2(3),1)%>" target='_blank'><%=BreakUrl(MyArray_2(3),1)%></a>[<a href="<%=MyArray_2(3)%>" target='_blank'>��ϸ</a>] 
              <%End If%></td>
            <td ><%=GetTurnTime(Hour(Cdate(MyArray_2(4))))&":"&GetTurnTime(Minute(Cdate(MyArray_2(4))))&":"&GetTurnTime(Second(Cdate(MyArray_2(4))))%></td>
            <td > <%If MyArray_2(5)="" Then%>
              ֱ�Ӵ���������� 
              <%Else%> <a href="<%=MyArray_2(5)%>" target='_blank'><%=BreakUrl(MyArray_2(5),2)%></a>[<a href="<%=MyArray_2(5)%>" target='_blank'>��ϸ</a>] 
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
            <td colspan="7"><%=Date()%> ��Դͳ��</td>
          </tr>
          <tr> 
            <td >���</td>
            <td >��Դ��վ</td>
            <td >��Դ����</td>
            <td >��ռ����ı���</td>
            <td >��վʱ��</td>
            <td >���IP��ַ</td>
            <td >������ʱ��</td>
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
  Response.Write "ֱ�Ӵ����������"
 Else
  Response.Write "<a href=http://"&Rs("LyHead")&" target='_blank'>"&Rs("LyHead")&"</a>&nbsp;[<a href="&Rs("Ly")&" target='_blank'>��ϸ</a>]"
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
            <td colspan="8" >��������</td>
          </tr>
          <tr> 
            <td > <p >���</td>
            <td >IP</td>
            <td >���ʴ���</td>
            <td >����ʱ��</td>
            <td >��Դҳ��</td>
            <td >������ʱ��</td>
            <td >������ҳ��</td>
            <td >�������ҳ</td>
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
              ֱ�Ӵ���������� 
              <%Else%> <a href="http://<%=BreakUrl(MyArray_2(3),1)%>" target='_blank'><%=BreakUrl(MyArray_2(3),1)%></a>[<a href="<%=MyArray_2(3)%>" target='_blank'>��ϸ</a>] 
              <%End If%></td>
            <td ><%=GetTurnTime(Hour(Cdate(MyArray_2(4))))&":"&GetTurnTime(Minute(Cdate(MyArray_2(4))))&":"&GetTurnTime(Second(Cdate(MyArray_2(4))))%></td>
            <td > <%If MyArray_2(5)="" Then%>
              ֱ�Ӵ���������� 
              <%Else%> <a href="<%=MyArray_2(5)%>" target='_blank'><%=BreakUrl(MyArray_2(5),2)%></a>[<a href="<%=MyArray_2(5)%>" target='_blank'>��ϸ</a>] 
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
            <td colspan="4"><%=AddDate%> ÿСʱ����</td>
          </tr>
          <tr> 
            <td>Сʱ</td>
            <td width="134">��ʾ����</td>
            <td width="135">IP����</td>
            <td>����</td>
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
            <td colspan="6" >ͼ���� <img src=images/BlueBar.gif width=30 height=10>��ʾ����<img src=images/GreenBar.gif width=30 height=10>IP��</td>
          </tr>
		  
          <tr> 
            <td colspan="4"><%=AddDate%>�����ڣ��ܼ����<%=ShowTotal%>�� 
                IP<%=IpTotal%>��</td>
          </tr>
        </table>
        <%End If%>
        <%If RsUser("TodayHourOpen")=-1 Then%>

<table class="tb_2 tb_underline">
  <tr class="tr_1"> 
            <td colspan="4">�շ���ͳ��
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
            <td>����</td>
            <td>��ʾ����</td>
            <td>IP����</td>
            <td>����</td>
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
            <td colspan="4" >ͼ���� <img src=images/BlueBar.gif width=30 height=10>��ʾ����<img src=images/GreenBar.gif width=30 height=10>IP��</td>
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
              ����<%=TotalRs%>���¼����<%=Rs1(0)%>�������<%=Rs1(1)%>��IP��ƽ��ÿ��<%=AvgVisit%>�η��ʣ�ƽ��ÿ��IP<%=AvgIp%>��</td>
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
      <td><a href="?UserName=<%=UserName%>&Page=1">��һҳ</a> 
          <%
if page>1 then%>
          <a href='?UserName=<%=UserName%>&Page=<%=Page-1%>'>��һҳ</a> 
          <%
end if
%>
          <%
if page<rs.pagecount   then%>
          <a href='?UserName=<%=UserName%>&Page=<%=Page+1%>'>��һҳ</a> 
          <%
end if
%>
          <a href="?UserName=<%=UserName%>&Page=<%=TotalPage%>">���һҳ</a> 
          ҳ�Σ�<font color="#ff0000"><%=Page%></font>/<%=TotalPage%>ҳ
                
<%
Response.write "&nbsp;&nbsp;ת����<select id='page' onChange=""window.location=document.getElementById('page').options[document.getElementById('page').selectedIndex].value"">"
For I=1 To TotalPage
    Response.Write "<option value=?UserName="&UserName&"&page="&i     
 If Page=I Then Response.Write " selected"
 Response.Write ">"& I &"</option>"
Next
Response.write "</select>ҳ"
%></td>
          </tr>
        </table></td>
    </tr>
  </table>
<%End If%>

<%End If%>
