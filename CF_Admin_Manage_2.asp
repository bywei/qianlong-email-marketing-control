
<%If Action="" Or Action="sysstate" Then%>
<table >
  <tr > 
              <td colspan="2">ϵͳ״̬</td>
  </tr>
            <tr> 
              <td width="18%">ϵͳ���ƣ�</td>
              <td ><%=RsSet("SysTitle")%> </td>
            </tr>
            <tr> 
              <td >ϵͳע��״̬��</td>
              <td> <%If RsSet("RegState")=0 Then%>
                Ŀǰ�Ѿ���ֹ���û������������ 
                <%Else%>
                Ŀǰ�������û���������������ֻ���Լ��ÿ��Թر��û�ע�Ṧ�ܣ� 
                <%End If%>
                [<a href="?Action=sysset">��������</a>]</td>
            </tr>
            <tr> 
              <td >��Դͳ�ƹ�ִ�д�����</td>
              <td><%=RsSet("Store_TotalLy")%> [<a href="?Action=lyreset" onClick="{if(confirm('ȷ��Ҫ��ԭ��Դͳ����ô?')){return true;}return false;};">��ԭΪ0</a>]</td>
            </tr>
            <tr> 
              <td >ҳ����Դ������</td>
              <td><%
Sql="Select Count(*) From CFCount_Ly_Day"
Set Rs=Conn.ExeCute(Sql)
Response.Write Rs(0)
%>
                [<a href="?Action=beforelydel" onClick="{if(confirm('ȷ��Ҫɾ��������ǰ����Դô?')){return true;}return false;};">ɾ��������ǰ����Դ</a>]</td>
            </tr>
            <tr> 
              <td >����ҳ����Դ����</td>
              <td><%
Sql="Select Count(*) From CFCount_Ly_Day Where DateDiff('d',AddDate,Date())=0"
Set Rs=Conn.ExeCute(Sql)
Response.Write Rs(0)
%> &nbsp;</td>
            </tr>
            <tr> 
              <td >ҳ��ͳ��������</td>
              <td><%
Sql="Select Count(*) From CFCount_Web_Day"
Set Rs=Conn.ExeCute(Sql)
Response.Write Rs(0)
%>
                [<a href="?Action=beforewebdel" onClick="{if(confirm('ȷ��Ҫɾ��������ǰ��ҳ��ͳ��ô?')){return true;}return false;};">ɾ��������ǰ��ҳ��ͳ��</a>]</td>
            </tr>
            <tr> 
              <td >�����ҳ��ͳ������</td>
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
    <td colspan="9">����������ѡ������Ӧ������ʽ</td>
  </tr>
  <tr > 
    <td><a href="?Action=<%=Action%>&UserName=<%=UserName%>&Px=ID">ID</a></td>
    <td><a href="?Action=<%=Action%>&UserName=<%=UserName%>&Px=UserName">�û���</a></td>
    <td><a href="?Action=<%=Action%>&UserName=<%=UserName%>&Px=ParentName">�˺����</a></td>
    <td><a href="?Action=<%=Action%>&UserName=<%=UserName%>&Px=LoginTotal">��½����</a></td>
    <td><a href="?Action=<%=Action%>&UserName=<%=UserName%>&Px=RealIpTotal">ʵ�ʼ���</a></td>
    <td><a href="?Action=<%=Action%>&UserName=<%=UserName%>&Px=PageName">��ҳ</a></td>
    <td><a href="?Action=<%=Action%>&UserName=<%=UserName%>&Px=AddTime">ע��ʱ��</a></td>
    <td><a href="?Action=<%=Action%>&UserName=<%=UserName%>&Px=UserState">�ʺ�״̬</a></td>
    <td>����</td>
  </tr>
<%

Sql="Select Count(*) From CFCount_User" 
Set Rs=Conn.Execute(Sql)
Response.write "<b><font color=ff0000>Ŀǰ�˺�������"&Rs(0)&"</font></b>"

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
	 Response.write "�޸��˺�"
    Else
	 Response.write "���˺�:"&Rs("ParentName")
	End If
	%></td>
    <td><%=Rs("LoginTotal")%></td>
    <td><%=Rs("RealIpTotal")%></td>
    <td><a href="<%=Rs("PageUrl")%>" target="_blank"><%=Rs("PageName")%></a></td>
    <td><%=year(Rs("AddTime"))&"-"&month(Rs("AddTime"))&"-"&day(Rs("AddTime"))%></td>
    <td> 
        <%if rs("UserState")=-1 then
			  response.Write "��Ч"
			  elseif rs("UserState")=0 then
			  response.Write "<font color=#ff0000>��Ч</font>"
			  end if%>      </td>
    <td><a href="?Action=usergo&UserName=<%=Rs("UserName")%>" target="_blank">�鿴</a> 
        <a href="?Action=usermodify&ID=<%=Rs("ID")%>">�޸�</a> 
        <a href="?Action=userdel&UserName=<%=Rs("UserName")%>" onClick="{if(confirm('ȷ��Ҫɾ��ô?')){return true;}return false;}">ɾ��</a></td>
  </tr>
  <%rs.movenext
mypagesize=mypagesize-1
wend%>
</table>
          <table class="tb_3">
            <tr> 
              <td><a href="?Action=<%=Action%>&UserName=<%=UserName%>&Px=<%=Px%>&Page=1">��һҳ</a> 
                  <%
if page>1 then%>
                  <a href='?Action=<%=Action%>&UserName=<%=UserName%>&Px=<%=Px%>&Page=<%=Page-1%>'>��һҳ</a> 
                  <%
end if
%>
                  <%
if page<rs.pagecount   then%>
                  <a href='?Action=<%=Action%>&UserName=<%=UserName%>&Px=<%=Px%>&Page=<%=Page+1%>'>��һҳ</a> 
                  <%
end if
%>
                  <a href="?Action=<%=Action%>&UserName=<%=UserName%>&Px=<%=Px%>&Page=<%=TotalPage%>">���һҳ</a> 
                  ҳ�Σ�<font color="#ff0000"><%=page%></font>/<%=totalpage%>ҳ&nbsp;&nbsp;��<%=totalrs%>����¼&nbsp;&nbsp;ÿҳ��ʾ<%=rs.pagesize%>��
<%
Response.write "&nbsp;&nbsp;ת����<select id='page' onChange=""window.location=document.getElementById('page').options[document.getElementById('page').selectedIndex].value"">"
For I=1 To TotalPage
    Response.Write "<option value=?Action="&Action&"&UserName="&UserName&"&Px="&Px&"&Page="&I
 If Page=I Then Response.Write " selected"
 Response.Write ">"& I &"</option>"
Next
Response.write "</select>ҳ"
%></td>
            </tr>
          </table>
          <table class="tb_3">
            <form name="form2" method="post" action="?action=searchuser">
              <tr> 
                <td>�������û��� 
                    <input name="UserName" type="text" id="UserName" size="10">
                    <input type="submit" name="Submit" value="����">
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
      <td colspan="2">�޸����� </td>
    </tr>
    <tr> 
      <td width="150">�û�����</td>
      <td><%=Rs("UserName")%></td>
    </tr>
    <tr> 
      <td>�޸�Ϊ�����룺</td>
      <td><input name="Pwd_New" type="text" id="Password_New">
        (������ԭ���벻�ᱻ�޸�)</td>
    </tr>
    <tr> 
      <td>�޸�Ϊ�������һش𰸣�</td>
      <td><input name="PasswordAnswer_New" type="text" id="PasswordAnswer_New">
        (������ԭ�����һش𰸲��ᱻ�޸�)</td>
    </tr>
    <tr> 
      <td>�ʺ�״̬��</td>
      <td><select name="UserState" id="UserState">
          <option value="-1"<%If Rs("UserState")=-1 Then response.write" selected"%>>��Ч</option>
          <option value="0"<%If Rs("UserState")=0 Then response.write" selected"%>>��Ч</option>
        </select></td>
    </tr>
    <tr> 
      <td>����Ա˵����</td>
      <td><textarea name="admindesc" cols="30" rows="6" id="notes"><%=rs("admindesc")%></textarea></td>
    </tr>
    <tr> 
      <td colspan="2">  
          <input type="submit" name="Submit3" value="�޸�">
          ���� 
          <input type="reset" name="Submit523" value="ȡ��" onClick="javascript:history.go(-1)">
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
              <td colspan="5">ÿ��ͳ��</td>
  </tr>
            <tr class="tr_2"> 
              <td >���</td>
              <td >����</td>
              <td >�������</td>
              <td >IP����</td>
              <td >��ռ����ı���</td>
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
              <td colspan="5" >ͼ���� <img src=images/BlueBar.gif width=30 height=10>��ʾ����<img src=images/GreenBar.gif width=30 height=10>IP��</td>
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
                ����<%=TotalRs%>���¼����<%=Rs1(0)%>�������<%=Rs1(1)%>��IP��ƽ��ÿ��<%=AvgVisit%>�η��ʣ�ƽ��ÿ��IP<%=AvgIp%>�� </td>
            </tr>
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
                  <a href="?Action=<%=Action%>&Page=<%=totalpage%>">���һҳ</a> 
                  ҳ�Σ�<font color="#ff0000"><%=page%></font>/<%=totalpage%>ҳ&nbsp;&nbsp;��<%=totalrs%>����¼&nbsp;&nbsp;ÿҳ��ʾ<%=rs.pagesize%>��
<%
Response.write "&nbsp;&nbsp;ת����<select id='page' onChange=""window.location=document.getElementById('page').options[document.getElementById('page').selectedIndex].value"">"
For I=1 To TotalPage
    Response.Write "<option value=?Action="&Action&"&Page="&i
 If Page=I Then Response.Write " selected"
 Response.Write ">"& I &"</option>"
Next
Response.write "</select>ҳ"
%></td>

            </tr>
          </table>
          <%End If%>
          <%If Action="lylist" Then%>
<table >

            <form name="form6" method="post" action="?Action=lydel">
<tr>
                <td colspan="7" >ѡ���ѯ������ 
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
 If IsDate(Cdate(AddDate))=False Then Call AlertBack("�����˴����ʱ��",1)
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
                <td colspan="7"><%=AddDate%> ��Դͳ��[�ɵ����������]</td>
              </tr>
              <tr class="tr_2"> 
                <td ><a href="?Action=<%=Action%>&AddDate=<%=AddDate%>&Px=ID">���</a></td>
                <td ><a href="?Action=<%=Action%>&AddDate=<%=AddDate%>&Px=LyHead">��Դ��վ</a></td>
                <td ><a href="?Action=<%=Action%>&AddDate=<%=AddDate%>&Px=MyCounter">��Դ����</a></td>
                <td ><a href="?Action=<%=Action%>&AddDate=<%=AddDate%>&Px=MyCounter">��ռ����ı���</a></td>
                <td ><a href="?Action=<%=Action%>&AddDate=<%=AddDate%>&Px=AddTime">��վʱ��</a></td>
                <td ><a href="?Action=<%=Action%>&AddDate=<%=AddDate%>&Px=Ip">���IP��ַ</a></td>
                <td ><a href="?Action=<%=Action%>&AddDate=<%=AddDate%>&Px=LastTime">������ʱ��</a></td>
              </tr>
              <%
While Not Rs.Eof And MyPageSize>0
%>
              <tr class="tr_2"> 
                <td><%=Rs.RecordCount-Rs.Pagesize*(page)+mypagesize%></td>
                <td> 
                    <%If Rs("LyHead")="" Then
  Response.Write "ֱ�Ӵ����������"
 Else
  Response.Write "<a href=http://"&Rs("LyHead")&" target='_blank'>"&Rs("LyHead")&"</a>&nbsp;[<a href="&Rs("Ly")&" target='_blank'>��ϸ</a>]"
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
              <td><a href="?Action=<%=Action%>&AddDate=<%=AddDate%>&Px=<%=Px%>&Page=1">��һҳ</a> 
                  <%
if page>1 then%>
                  <a href='?Action=<%=Action%>&AddDate=<%=AddDate%>&Px=<%=Px%>&Page=<%=Page-1%>'>��һҳ</a> 
                  <%
end if
%>
                  <%
if page<rs.pagecount   then%>
                  <a href='?Action=<%=Action%>&AddDate=<%=AddDate%>&Px=<%=Px%>&Page=<%=Page+1%>'>��һҳ</a> 
                  <%
end if
%>
                  <a href="?Action=<%=Action%>&AddDate=<%=AddDate%>&Px=<%=Px%>&Page=<%=totalpage%>">���һҳ</a> 
                  ҳ�Σ�<font color="#ff0000"><%=page%></font>/<%=totalpage%>ҳ&nbsp;&nbsp;��<%=totalrs%>����¼&nbsp;&nbsp;ÿҳ��ʾ<%=rs.pagesize%>��
<%
Response.write "&nbsp;&nbsp;ת����<select id='page' onChange=""window.location=document.getElementById('page').options[document.getElementById('page').selectedIndex].value"">"
For I=1 To TotalPage
    Response.Write "<option value=?Action="&Action&"&AddDate="&AddDate&"&Px="&Px&"&Page="&i
 If Page=I Then Response.Write " selected"
 Response.Write ">"& I &"</option>"
Next
Response.write "</select>ҳ"
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
      <td colspan="2">ϵͳ����</td>
    </tr>
    <tr> 
      <td width="22%">������ϵͳ���ƣ�</td>
      <td><input name="SysTitle" type="text" id="mc" value="<%=Rs("SysTitle")%>" size="30"></td>
    </tr>
    <tr> 
      <td>��������ƣ�</td>
      <td><input name="TjTextName" type="text" id="mc" value="<%=Rs("TjTextName")%>" size="30">
        (8��������)</td>
    </tr>
    <tr> 
      <td>�����Ƿ��������������</td>
      <td> <input type="radio" name="RegState" value="-1" <%If Rs("RegState")=-1 Then%>checked<%End If%>>
        �� ���� 
        <input name="RegState" type="radio" value="0" <%If Rs("RegState")=0 Then%>checked<%End If%>>
        ��<br> &nbsp;ע�����������ֻ���Լ��ã���ע���û�������빦�ܹرա�<br> <%If Rs("RegState")=0 Then%> <font color="#FF0000">��ʾ��</font>Ŀǰ�Ѿ���ֹ����������� 
        <%End If%> </td>
    </tr>
    <tr> 
      <td>�Ƿ��¼��Դҳ��,IP��Ϣ��</td>
      <td><input type="radio" name="LyKeep" value="-1" <%If Rs("LyKeep")=-1 Then%>checked<%End If%>>
        �� ���� 
        <input name="LyKeep" type="radio" value="0" <%If Rs("LyKeep")=0 Then%>checked<%End If%>>
        ��<br> <%if rs("LyKeep")=0 then%> <font color="#FF0000">��ʾ��</font>Ŀǰ�Ѿ���ֹ��¼IP����Դ��Ϣ�� 
        <%end if%> </td>
    </tr>
    <tr> 
      <td>�Ƿ��¼����������</td>
      <td><input type="radio" name="OnlineKeep" value="-1" <%if rs("OnlineKeep")=-1 then%>checked<%end if%>>
        �� ���� 
        <input name="OnlineKeep" type="radio" value="0" <%if rs("OnlineKeep")=0 then%>checked<%end if%>>
        ��<br> <%if rs("OnlineKeep")=0 then%> <font color="#FF0000">��ʾ��</font>Ŀǰ�Ѿ���ֹ��¼����������Ϣ�� 
        <%end if%> </td>
    </tr>
    <tr> 
      <td>�Ƿ��¼��վҳ�����������</td>
      <td><input type="radio" name="WebKeep" value="-1" <%if rs("WebKeep")=-1 then%>checked<%end if%>>
        �� ���� 
        <input name="WebKeep" type="radio" value="0" <%if rs("WebKeep")=0 then%>checked<%end if%>>
        ��<br> <%if rs("WebKeep")=0 then%> <font color="#FF0000">��ʾ��</font>Ŀǰ�Ѿ���ֹ��¼��վҳ����������� 
        <%end if%> </td>
    </tr>
    <tr> 
      <td>����ͼƬ��ʽ������</td>
      <td> <input name="StyleTotal" type="text" id="stylenum" value="<%=rs("StyleTotal")%>" size="30">
        ��</td>
    </tr>
    <tr> 
      <td>������LogoͼƬ��ַ��</td>
      <td><input name="LogoUrl" type="text" id="LogoUrl" value="<%=rs("LogoUrl")%>" size="30">
        ��С480*60</td>
    </tr>
    <tr> 
      <td>��ʾIP��ַ��IP����</td>
      <td class="td_1"><input type="radio" name="IpArea" value="-1"<%if rs("IpArea")=-1 then Response.write " checked"%>>
        �� 
        <input type="radio" name="IpArea" value="0"<%if rs("IpArea")=0 then Response.write " checked"%>>
        ��<br>
        ע������IP���ݿ�ϴ�,Ĭ��Ϊ��,��Ҫ��ʾ����IP����λ�õĻ�����js.tb11.com/js/������IP���ݿ��,��ѡ���ѡ��Ϊ&quot;��&quot;</td>
    </tr>
    <tr>
      <td>ϵͳƤ����</td>
      <td><input type="radio" name="skintype" value="0"<%if rs("skintype")=0 then Response.write " checked"%> onclick="document.getElementById('skincolor').value='CACACA|F3F9FE';">
        Ĭ�� 
        <input type="radio" name="skintype" value="1"<%if rs("skintype")=1 then Response.write " checked"%>  onclick="document.getElementById('skincolor').value='43A102|E4F9D5';">
        ��ʽ1 
        <input type="radio" name="skintype" value="2" <%if rs("skintype")=2 then Response.write " checked"%>  onclick="document.getElementById('skincolor').value='A73800|F6BDA0';">
        ��ʽ2 
        <input type="radio" name="skintype" value="3"<%if rs("skintype")=3 then Response.write " checked"%>  onclick="document.getElementById('skincolor').value='FFA500|FBD38B';">
        ��ʽ3 
        <input type="radio" name="skintype" value="4"<%if rs("skintype")=4 then Response.write " checked"%>  onclick="document.getElementById('skincolor').value='4C4C4C|E7E4E4';">
        ��ʽ4 </td>
    </tr>
    <tr id="a_1"> 
      <td>�Զ���Ƥ����ɫ��</td>
      <td><input name="skincolor" type="text" id="skincolor" value="<%=rs("SkinColor")%>" size="30">
        (������ɫ���룬Ĭ��ֵΪ��BFE7F1)</td>
    </tr>
    <tr> 
      <td>�����ʼ������</td>
      <td><input type="radio" name="emailsendtype" value="1"<%if rs("emailsendtype")=1 then%> checked<%end if%> onclick="show_1()">
        CDONTS��� 
        <%
on error resume next
set TestObj=server.CreateObject ("CDONTS.NewMail")
If -2147221005 <> Err then
 Response.write "(֧��)"
Else
  Response.write "(��֧��)"
End If
%> <input type="radio" name="emailsendtype" value="2"<%if rs("emailsendtype")=2 then%> checked<%end if%> onclick="show_2()">
        JMAIL��� 
        <%
on error resume next
set TestObj=server.CreateObject ("JMail.SmtpMail")
If -2147221005 <> Err then
 Response.write "(֧��)"
Else
  Response.write "(��֧��)"
End If
%></td>
    </tr>
    <tr id="t_2_1"> 
      <td>��������ַ��</td>
      <td><input name="SmtpAddress" type="text" value="<%=rs("SmtpAddress")%>" size="30">
      (��smtp.a.com)</td>
    </tr>
    <tr id="t_2_2"> 
      <td>�����ʼ��ʺ����ƣ�</td>
      <td><input name="SmtpUser" type="text" value="<%=rs("SmtpUser")%>" size="30">
      (��a@a.com)</td>
    </tr>
    <tr id="t_2_3"> 
      <td>�����ʼ��ʺŵ����룺</td>
      <td><input name="SmtpPassword" type="password" value="<%=rs("SmtpPassword")%>" size="30">
      <br>
      �磺����һ���ʼ�b@a.com������Ϊc�����ŷ�������ַ��������Ϊ��<br>smtp.a.com�����ʼ��ʺ���������Ϊb@a.com�������ʼ��ʺŵ���<br>������Ϊc�������ʼ�����</td>
    </tr>
    <tr> 
      <td></td>  <td>  
          <input type="submit" name="Submit32" value="�޸�">
          ���� 
          <input type="reset" name="Submit5232" value="ȡ��">
      </td>
    </tr>
  </form>
</table>
          <%End If%>
		  
<%If Action="otherset" Then%>

          
<table >

  <form name="form2" method="post" action="?action=othersetmodifysave">
  <tr > 
      <td colspan="2">��������</td>
    </tr>
    <tr> 
      <td>ϵͳ��ȫ�룺</td>
      <td><input name="SysCode" type="text" id="SysCode" value="<%=GetMySet("SysCode")%>" size="25">
        *�������������8�����������ַ���ǿ��ȫ</td>
    </tr>
    <tr> 
      <td>Сʱͳ�Ʊ���������</td>
      <td><input name="HourCountKeepDay" type="text" id="HourCountKeepDay" value="<%=GetMySet("HourCountKeepDay")%>" size="25">
        *Ĭ��2��</td>
    </tr>
    <tr> 
      <td>��ҳ�����¼����������</td>
      <td><input name="WebKeepDay" type="text" id="WebKeepDay" value="<%=GetMySet("WebKeepDay")%>" size="25">
        *Ĭ��2��</td>
    </tr>	
    <tr> 
      <td>��Դ��¼����������</td>
      <td><input name="LyKeepDay" type="text" id="LyKeepDay" value="<%=GetMySet("LyKeepDay")%>" size="25">
        *Ĭ��2��</td>
    </tr>
    <tr> 
      <td>���������¼����������</td>
      <td><input name="SearchKeepDay" type="text" id="SearchKeepDay" value="<%=GetMySet("SearchKeepDay")%>" size="25">
        *Ĭ��7��</td>
    </tr>
    <tr> 
      <td>�����ؼ��ּ�¼����������</td>
      <td><input name="KeywordKeepDay" type="text" id="KeywordKeepDay" value="<%=GetMySet("KeywordKeepDay")%>" size="25">
        *Ĭ��2��</td>
    </tr>
    <tr> 
      <td>������¼����������</td>
      <td><input name="CountKeepDay" type="text" id="CountKeepDay" value="<%=GetMySet("CountKeepDay")%>" size="25">
        *Ĭ��60��</td>
    </tr>
	<tr> 
      <td></td>  <td>  
          <input type="submit" name="Submit32" value="�޸�">
          ���� 
          <input type="reset" name="Submit5232" value="ȡ��">
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
              <td>���</td>
              <td>��������������</td>
              <td>��������Ӣ�ı�ʶ</td>
              <td>�����ؼ��ֵĲ���</td>
              <td>����</td>
  </tr>
            <%While Not Rs.Eof
			I=I+1%>
            <tr class="tr_2"> 
              <td ><%=I%></td>
              <td ><a href="http://www.<%=Rs("SiteFlag")%>" target=_blank><%=Rs("SiteDesc")%></a></td>
              <td ><a href="http://www.<%=Rs("SiteFlag")%>" target=_blank><%=Rs("SiteFlag")%></a></td>
              <td ><%=Rs("KeywordFlag")%></td>
              <td > 
                  <a href="?Action=searchmodify&ID=<%=Rs("ID")%>">�޸�</a> 
                  <a href="?Action=searchdel&SiteFlag=<%=Rs("SiteFlag")%>" onClick="{if(confirm('ȷ��Ҫɾ��ô?')){return true;}return false;}">ɾ��</a> 
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
              <td colspan="4">�����Զ�������</td>
  </tr>
            <form name="form5" method="post" action="?Action=searchaddsave">
              <tr class="tr_2"> 
                <td>��������������</td>
                <td width="191">��������Ӣ�ı�ʶ</td>
                <td width="191">�����ؼ��ֵĲ���</td>
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
                <td><input type="submit" name="Submit" value="������"></td>
              </tr>
              <tr> 
                <td colspan="4"  class="td_1">���磺<br>
        �ٶ���ҳ����&quot;abcd&quot;IE��ַ��Ϊ:http://www.baidu.com/s?wd=abcd&amp;cl=3<br>
        �ٶ���������&quot;abcd&quot;IE��ַ��Ϊ:http://post.baidu.com/f?kw=abcd<br>
        �ٶ�֪������&quot;abcd&quot;IE��ַ��Ϊ:http://zhidao.baidu.com/q?ct=17&amp;pn=0&amp;tn=ikaslist&amp;rn=10&amp;word=abcd&amp;t=51<br>
        ��Ҫ���ǾͿ���������д������������������д&quot;�ٶ�&quot;,��������Ӣ�ı�ʶ��д&quot;baidu.com&quot;����ʶ������д�����������ʶ����Ӣ�ķֺ�;�ָ��������ؼ��ֵĲ�����д&quot;wd;kw;word&quot;,�ؼ��ֲ�����Ӣ�ķֺ�;�ָ����Դ�����</td>
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
              <td colspan="4">�޸���������</td>
  </tr>
            <form name="form5" method="post" action="?Action=searchmodifysave&ID=<%=ID%>">
              <tr class="tr_2"> 
                <td>��������������</td>
                <td width="191">��������Ӣ�ı�ʶ</td>
                <td width="191">�����ؼ��ֵĲ���</td>
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
                <td><input type="submit" name="Submit" value="�ޡ���"></td>
              </tr>
              <tr> 
                <td colspan="4" class="td_1">������ӻ��޸����ӣ��������ڰٶ���������ҳ&quot;����&quot;,�ڽ��ҳ�ĵ�ַ�������ַ��http://www.baidu.com/s?wd=%B3%CB%B7%E7&amp;cl=3<br>
                  �������Ǿ�������д������������������д&quot;�ٶ�&quot;,��������Ӣ�ı�ʶ��д&quot;baidu.com&quot;,�����ؼ��ֵĲ�����д&quot;wd&quot;</td>
              </tr>
            </form>
</table>
          <%End If%>
		  
		  
<%if Action="imglist" then%>
<table >
  <tr > 
 <td>������ͼƬ����</td>
</tr>
<tr> 
 <td><iframe style="top:2px" ID="UploadFiles" src="CF_Upfile.asp?UserType=0&CheckCode=<%=RsSet("AdminCode")%>" frameborder=0 scrolling=no width="450" height="25"></iframe></td>
</tr>


  <tr> 
    <td>

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
 

Response.write "�ļ���С��" & CInt(Rs("FileSizeNum")/1024) &"k<br>"

Response.write "ʹ��������"
Sql="Select Count(*) From CFCount_User Where ImgFileName='"&Rs("FileName")&"'"
Set Rs2=Conn.Execute(Sql)
Response.write Rs2(0)
Rs2.Close()
Response.write "<br>"

Response.write "�ϴ�ʱ�䣺"&Rs("AddTime")&"<br>"

Response.write "������<a href='?Action=picdel&FileName="&Rs("FileName")&"'onClick=""{if(confirm('ȷ��Ҫɾ��ô?')){return true;}return false;};"">ɾ��</a>"
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
        </table>
		
    </td>
  </tr>

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
      <td colspan="2">����ԭ����Ա����</td>
    </tr>
    <tr>
      <td width="16%">ԭ����Ա�û�����</td>
      <td><%=Rs("Admin")%></td>
    </tr>
    <tr> 
      <td>������ԭ�������룺</td>
      <td><input name="Pwd_Old" type="password" id="Pwd_Old" value="" size="15"></td>
    </tr>
  <tr > 
      <td colspan="2">�¹���Ա���ƺ�����</td>
    </tr>
    <tr> 
      <td>�¹���Ա���ƣ�</td>
      <td><input name="Admin" type="text" id="admin" value="<%=Rs("Admin")%>" size="15"></td>
    </tr>
    <tr> 
      <td>�����룺</td>
      <td><input name="Pwd" type="password" id="Pwd" size="15"></td>
    </tr>
    <tr> 
      <td>������ȷ�ϣ� </td>
      <td><input name="Pwd2" type="password" id="Pwd2" size="15"></td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
      <td><input type="submit" name="Submit323" value="�޸�">
        ���� 
        <input type="reset" name="Submit5233" value="ȡ��"> </td>
    </tr>
  </form>
</table>
<%End If%>


<%If Action="dbdo" Then%>         
  
<table >
  <tr > 
    <td>ִ��Sql���</td>
  </tr>
  <form name="form2" method="post" action="?action=sqlexesave">
    <tr> 
      <td><textarea name="sql" cols="80" rows="5" id="sqlcmd"></textarea> 
      </td>
    </tr>
    <tr>
      <td>���������Ա����ȷ��ִ�У�
        <input type="password" name="Pwd"></td>
    </tr>
    <tr>
      <td><input type="submit" name="Submit" value="ȷ��ִ������Sql���" onclick="{if(confirm('ȷ��ִ������Sql���ô?')){return true;}return false;};">
        ע���������ִ���Իس����У�ÿһ��Ҫ��Ϊһ�������</td>
    </tr>
  </form>
  <tr >  
    <td >���ݿ���־ѹ��</td>
  </tr>
  <tr> 
    <td><input type="button" name="Submit2" value="���ѹ�����ݿ�" onclick="{if(confirm('ȷ��Ҫѹ�����ݿ�ô?')){window.location='?Action=dbys';return true;}return false;};">
      ע���κ�ʱ����Բ�������ÿ�β�����Ҫ���̫�̣���һ�����ڲ����ü���</td>
  </tr>
  <tr > 
    <td> ���ݿⱸ��</td>
  </tr>
  <tr> 
    <td><input type="button" name="Submit2" value="����������ݿ�" onclick="{if(confirm('ȷ��Ҫ�������ݿ�ô?�˲���������Ҫ���ѱȽϳ���ʱ��')){window.location='?Action=dbbackup';return true;}return false;};">
      ע�������ǳ�������ݿ����һ̨�������ϣ����Ҳ��������������ϣ������������������ֹ��Լ�����,ÿ�β�����Ҫ���̫�̣���һ�����ڲ����ü��Σ��������ݿ���ڳ����dataĿ¼���backup��ͷ�����Ǳ��ݵ�ʱ����ļ���ʽ���������ݿ��Ļ��ɴ˲�����Ҫ���ѱȽϳ���ʱ��</td>
  </tr>
  <tr > 
    <td  class="td_1"><%
foldername=server.mappath("data")
path=foldername
If Instr(path,"\data")=0 Then Call AlertBack("����������Ŀ¼��",1)
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
          <td>����</td>
          <td>��С</td>
          <td>����</td>
          <td>�޸�ʱ��</td>
        </tr>
        <% 
'�г�Ŀ¼ 
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
          <td>�ļ���</td>
          <td><%= f1.datelastmodified%></td>
        </tr>
        <% 
Next 
foldername=StrReverse(left(StrReverse(folderspec),instr(StrReverse(folderspec),"\")-1))&"/"
'�г��ļ� 
Set fc = f.Files 
For Each f1 in fc 
%>
        <tr> 
          <td><%= f1.name%>&nbsp;&nbsp; <%If Instr(Lcase(f1.name),"backup")>0 Then%>
            [<a href="?Action=dbdel&DbMc=<%= f1.name%>" onClick="{if(confirm('ȷ��Ҫɾ��ô?')){return true;}return false;}">ɾ��</a>] 
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