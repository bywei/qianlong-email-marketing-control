
<%
If Action="subaddsave" Or Action="subgoto" Or Action="submodifysave" Or Action="subdel" Or Action="jsqsetsave" Or Action="jsqquicksetsave" Or Action="datamodifysave" Or Action="pwdmodifysave" Or Action="passwordanswermodifysave" Or Action="stylemodifysave" Or Action="imgmodifysave" Or Action="jsqresetsave" Or Action="userdelsave" Or Action="gbookmodifysave" Or Action="gbookdel" Then
 If Session("CFCountUser_View")<>"" Then'Ϊ�����鿴��ʱ
  Call AlertBack("�����鿴���޷��Դ˲���",1)
 End If
End If

If Action="subaddsave" Then
 SubUserName=GoBack(ChkStr(Request("SubUserName"),1),"���������˺��û���!")
 Call CheckInput_Letter(SubUserName)
 
 Pwd=GoBack(ChkStr(Request("Pwd"),1),"���������˺�����!")
 Pwd2=GoBack(ChkStr(Request("Pwd2"),1),"�������ظ����˺�����!")
 
 If Pwd<>Pwd2 Then Call AlertBack("�ظ���������!",1)
 
 PageName=GoBack(Server.HtmlEncode(ChkStr(Request("PageName"),1)),"������վ������!")
 PageUrl=GoBack(Server.HtmlEncode(ChkStr(Request("PageUrl"),1)),"����������!")

 Set Rs=Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From CFCount_User Where UserName='"&SubUserName&"'"
 Rs.Open Sql,Conn,3,2
 If Not Rs.Eof Then
  Call AlertBack("���û����Ѿ�������ע��!",1)
 Else
  Rs.AddNew
  Rs("UserName")=SubUserName
  Rs("Pwd")=Md5(Pwd,1)
  Rs("ParentName")=UserName
  Rs("PageName")=PageName
  Rs("PageUrl")=PageUrl
  Rs.Update
 End If
 
 Call AlertBack("ע��ɹ�!",2)
End If


If Request("Action")="subgoto" Then
 SubUserName=ChkStr(Request("SubUserName"),1)
 Set Rs= Server.CreateObject("ADODB.Recordset")
 Sql="select * from CFCount_User where UserName='"&SubUserName&"'"
 Rs.Open Sql,Conn,1,1
 If Rs("ParentName")<>UserName Then Call AlertBack("û��Ȩ��",1)
 
 Session("CFCountSubAdmin")=UserName
 Session("CFCountUser")=SubUserName
 Response.Redirect("Manage.asp")
End If

If Request("Action")="parentgoto" Then
 Session("CFCountUser")=Session("CFCountSubAdmin")
 Session("CFCountSubAdmin")=""
 Response.Redirect("Manage.asp")
End If

If Action="submodifysave" Then

 SubUserName=ChkStr(Request("SubUserName"),1)
 
 PageName=GoBack(Server.HtmlEncode(ChkStr(Request("PageName"),1)),"������վ������!")
 PageUrl=GoBack(Server.HtmlEncode(ChkStr(Request("PageUrl"),1)),"����������!")

 Set Rs=Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From CFCount_User Where UserName='"&SubUserName&"' And ParentName='"&UserName&"'"
 Rs.Open Sql,Conn,3,2
 Rs("PageName")=PageName
 Rs("PageUrl")=PageUrl
 Rs.Update
 
 Call AlertBack("�޸ĳɹ�!",2)
End If

If Action="subdel" Then '�˺�ɾ��
 SubUserName=ChkStr(Request("SubUserName"),1)
 Set Rs= Server.CreateObject("ADODB.Recordset")
 Sql="select * from CFCount_User where UserName='"&SubUserName&"'"
 Rs.Open Sql,Conn,1,1
 If Rs("ParentName")<>UserName Then Call AlertBack("û��Ȩ��",1)

 Filename=SubUserName&".jpg"
 Filename_2=SubUserName&".gif"

 Set Fs = Server.CreateObject("Scripting.FileSystemObject")
 If Fs.FileExists(Server.mappath("Upload/"&Filename)) Then
  Set Os = Fs.GetFile(Server.mappath("Upload/"&Filename))
  Os.Delete
 End If

 If Fs.FileExists(Server.mappath("Upload/"&Filename_2)) Then
  Set Os = Fs.GetFile(Server.mappath("Upload/"&Filename_2))
  Os.Delete
 End If

 Sql="Delete From CFCount_Alexa_Day Where UserName='"&SubUserName&"'"
 Conn.ExeCute Sql
 Sql="Delete From CFCount_Back Where UserName='"&SubUserName&"'"
 Conn.ExeCute Sql
 Sql="Delete From CFCount_Count_Day Where UserName='"&SubUserName&"'"
 Conn.ExeCute Sql
 Sql="Delete From CFCount_Count_Hour Where UserName='"&SubUserName&"'"
 Conn.ExeCute Sql
 Sql="Delete From CFCount_Ly_Day Where UserName='"&SubUserName&"'"
 Conn.ExeCute Sql
 Sql="Delete From CFCount_User Where UserName='"&SubUserName&"'"
 Conn.ExeCute Sql
 Sql="Delete From CFCount_Search_Day Where UserName='"&SubUserName&"'"
 Conn.ExeCute Sql
 Sql="Delete From CFCount_SearchKeywrod_Day Where UserName='"&SubUserName&"'"
 Conn.ExeCute Sql

 Sql="Delete From CFCount_Web_Day Where UserName='"&UserName&"'"
 Conn.ExeCute Sql

 Call AlertBack("������˺��Ѿ���ȫɾ����",1)
End If
 

If Action="jsqsetsave" Then '����������
 ShowTotal=GoBack(ChkStr(Request("ShowTotal"),2),"�����������ʾ����!")
 PicNum=GoBack(ChkStr(Request("PicNum"),2),"�������������ʾλ��!")
 Adddate=GoBack(ChkStr(Request("Adddate"),3),"�������������ʼ����ʱ��!")
 CounterShow=CInt(Request("CounterShow"))
 ShowType=Int(Request("ShowType"))
 CounterHiddenPic=Int(Request("CounterHiddenPic"))
 CounterSite=Int(Request("CounterSite"))
 ImgCounterShow=Int(Request("ImgCounterShow"))
 ImgShowType=Int(Request("ImgShowType"))
 Online=Int(Request("Online"))
 OnlineTime=Request("OnlineTime")
 OnlineShow=Int(Request("OnlineShow"))
 TodayShow=Int(Request("TodayShow"))
 TodayIpShow=Int(Request("TodayIpShow"))
 IpShow=Int(Request("IpShow"))
 VisitShow=Int(Request("VisitShow"))

 TjOpen=Int(Request("TjOpen"))
 SurveyOpen=Int(Request("SurveyOpen"))
 TodayLyOpen=Int(Request("TodayLyOpen"))
 OnlineOpen=Int(Request("OnlineOpen"))
 TodayHourOpen=Int(Request("TodayHourOpen"))
 EveryDayOpen=Int(Request("EveryDayOpen"))
 
 GbookState=Int(Request("GbookState"))
 
 If Int(Picnum)<1 Or Int(Picnum)>8 Then Call AlertBack("�������������ʾλ����1-8֮��!",1)
 If Onlinetime="" Then Call AlertBack("�������¼���������ļ��ʱ��!",1)
 If onlinetime<1 Or onlinetime>180 Then Call AlertBack("ֻ������180��������,һ��������30����!",1)

 Set Rs=Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From CFCount_User Where UserName='"&UserName&"'"
 Rs.Open Sql,Conn,3,2
 Rs("PicNum")=PicNum
 Rs("ShowTotal")=ShowTotal
 Rs("AddDate")=AddDate
 Rs("CounterShow")=CounterShow
 Rs("ShowType")=ShowType
 Rs("CounterHiddenPic")=CounterHiddenPic
 Rs("CounterSite")=CounterSite
 Rs("ImgCounterShow")=ImgCounterShow
 Rs("ImgShowType")=ImgShowType
 Rs("Online")=Online
 Rs("OnlineTime")=OnlineTime
 Rs("OnlineShow")=OnlineShow
 Rs("TodayShow")=TodayShow
 Rs("TodayIpShow")=TodayIpShow
 Rs("IpShow")=IpShow
 Rs("VisitShow")=VisitShow
 Rs("TjOpen")=TjOpen
 Rs("SurveyOpen")=SurveyOpen
 Rs("TodayLyOpen")=TodayLyOpen
 Rs("OnlineOpen")=OnlineOpen
 Rs("TodayHourOpen")=TodayHourOpen
 Rs("EveryDayOpen")=EveryDayOpen
 Rs("GbookState")=GbookState
 Rs.Update
 
 Response.Cookies("CFCountCookie")=""
 Call AlertUrl("�޸ĳɹ�!","?Action=jsqset")
End If

If Action="jsqquicksetsave" Then '������������ʽ����
 QuickStyle=Int(Request("QuickStyle"))
 TjOpen=Int(Request("TjOpen"))
 
 If QuickStyle=0 Then Call AlertBack("��ѡ��һ�ֻ�����ʽ!",1)

 If QuickStyle=1 Or QuickStyle=4 Then
  CounterShow=-1
  ShowType=1
  CounterHiddenPic=2
  CounterSite=3
 End if

 If QuickStyle=2 Or QuickStyle=5 Then
  CounterShow=0
  ShowType=1
  CounterHiddenPic=2
  CounterSite=3
 End if
 
 If QuickStyle=3 Or QuickStyle=6 Then
  CounterShow=0
  ShowType=2
  CounterHiddenPic=1
  CounterSite=3
 End if
 
 If QuickStyle=7 Then
  CounterShow=0
  ShowType=2
  CounterHiddenPic=0
  CounterSite=3
 End if
 
 If QuickStyle=4 Or QuickStyle=5 Or QuickStyle=6 Then
  TjOpen=-1
  TjOpen_2=-1
 Else
  TjOpen_2=0
 End If
 
  Set Rs=Server.CreateObject("Adodb.RecordSet")
  Sql="Select * From CFCount_User Where UserName='"&UserName&"'"
  Rs.Open Sql,Conn,3,2
  Rs("CounterShow")=CounterShow
  Rs("ShowType")=ShowType
  Rs("CounterHiddenPic")=CounterHiddenPic
  Rs("CounterSite")=CounterSite
  Rs("OnlineShow")=TjOpen_2
  Rs("TodayShow")=TjOpen_2
  Rs("TodayIpShow")=TjOpen_2
  Rs("IpShow")=TjOpen_2
  Rs("VisitShow")=TjOpen_2
  Rs("TjOpen")=TjOpen
  Rs("SurveyOpen")=TjOpen
  Rs("TodayLyOpen")=TjOpen
  Rs("OnlineOpen")=TjOpen
  Rs("TodayHourOpen")=TjOpen
  Rs("EveryDayOpen")=TjOpen
  Rs.Update
 
 Call AlertUrl("���óɹ���ת��Ԥ��������ҳ��!","?Action=getcode")
End If

If Action="datamodifysave" Then
 Pwd=Goback(ChkStr(Request("Pwd"),1),"����������ȷ��")
 Email=GoBack(ChkStr(Request("Email"),1),"������E-mail!")
 QQ=ChkStr(Request("QQ"),1)
 PageName=GoBack(ChkStr(Request("PageName"),1),"��������ҳ����!")
 PageUrl=GoBack(ChkStr(Request("PageUrl"),1),"��������ҳ��ַ!")

 Set Rs=Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From CFCount_User Where UserName='"&UserName&"'"
 Rs.Open Sql,Conn,3,2
 If Rs("Pwd")<>Md5(Pwd,1) Then Call AlertBack("������������",1)

 Rs("Email")=Email
 Rs("QQ")=QQ
 Rs("PageName")=PageName
 Rs("PageUrl")=PageUrl
 Rs.Update
 
 Call AlertBack("�޸ĳɹ�!",1)
End If

If Action="pwdmodifysave" Then '�û��޸�����
 Pwd=goback(ChkStr(Request("Pwd"),1),"�������¹�������")
 Pwd2=goback(ChkStr(Request("Pwd2"),1),"������ȷ���¹�������")

 Pwd_View=goback(ChkStr(Request("Pwd_View"),1),"�������²鿴����")
 Pwd_View2=goback(ChkStr(Request("Pwd_View2"),1),"������ȷ���²鿴����")
 If Pwd_View<>Pwd_View2 Then Call AlertBack("�²鿴���벻һ��",1)
 If Pwd<>Pwd2 Then Call AlertBack("�¹������벻һ��",1)
 Pwd_Old=goback(ChkStr(Request("Pwd_Old"),1),"������ɹ�������ȷ��")
 
 Set Rs= Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From CFCount_User Where UserName='"&UserName&"'"
 Rs.Open Sql,Conn,3,2
 If Md5(Pwd_Old,1)<>Rs("Pwd") Then Call AlertBack("����ľɹ��������д���:",1)
 Rs("Pwd")=Md5(Pwd,1)
 Rs("Pwd_View")=Md5(Pwd_View,1)
 Rs.update
 
 Set Rs= Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From Users Where UserName='"&UserName&"'"
 Rs.Open Sql,Conn,3,2
 Rs("Pwd")=Md5(Pwd,1)
 Rs("UserPwd")=Pwd
 Rs.update

 Call AlertBack("�����޸ĳɹ�",1)
End If

If Action="passwordanswermodifysave" Then '�û��޸���ʾ����ʹ�
 
 Set rs = Server.CreateObject("ADODB.Recordset")
 sql="select * from CFCount_User where UserName='"&UserName&"'"
 rs.open sql,conn,1,1
 If Rs("PasswordAsk")<>"" Then
  PasswordAnswer_Old=GoBack(ChkStr(Request("PasswordAnswer_Old"),1),"������ԭ����ش��")
  if Rs("PasswordAnswer")<>Md5(PasswordAnswer_Old,1) Then Call AlertBack("ԭ����ش���д���",1)
 End If

 PasswordAsk_New=GoBack(ChkStr(Request("PasswordAsk_New"),1),"��������������ʾ����")
 PasswordAnswer_New=GoBack(ChkStr(Request("PasswordAnswer_New"),1),"������������ش��")

 Set Rs= Server.CreateObject("ADODB.Recordset")
 Sql="select * from CFCount_User where UserName='"&UserName&"'"
 Rs.Open Sql,Conn,3,2
 Rs("PasswordAsk")=PasswordAsk_New
 Rs("PasswordAnswer")=Md5(PasswordAnswer_New,1)
 Rs.update
 Call AlertBack("�޸ĳɹ�,���μǣ�",1)
End If

If Action="stylemodifysave" Then
 Style=GoBack(ChkStr(Request("Style"),2),"��ѡ��һ����ϲ����ʽ!")
 CounterShow=Int(Request("CounterShow"))

 Set Rs=Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From CFCount_User where UserName='"&UserName&"'"
 Rs.Open Sql,Conn,3,2
 Rs("CounterShow")=CounterShow
 Rs("style")=style
 Rs.Update

 Call AlertUrl("�޸ĳɹ�!","?Action=stylelist")
End If


If Action="imgmodifysave" Then
 Assort=Int(Request("Assort"))
 FileName_2=ChkStr(Request("FileName_2"),1)

 If Assort=2 And FileName_2="" Then
  Call AlertBack("��ѡ��ͼƬ",1)
 Else
  FileName=FileName_2
 End If

 If Assort=3 Then FileName="default.jpg"

 Set Rs=Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From CFCount_User where UserName='"&UserName&"'"
 Rs.Open Sql,Conn,3,2
 Rs("ImgFileName")=FileName
 Rs.Update

 Call AlertUrl("�޸ĳɹ�!","Manage.asp?Action=imglist")
End If

If Action="jsqresetsave" Then '����������
 Pwd=Goback(ChkStr(Request("Pwd"),1),"����������ȷ��")

 Set Rs= Server.CreateObject("ADODB.Recordset")
 Sql="select * from CFCount_User where UserName='"&UserName&"'"
 Rs.Open Sql,Conn,1,1
 If Rs("Pwd")<>Md5(Pwd,1) Then Call AlertBack("������������",1)

 If Int(Request("ly"))=-1 Then
  Sql="Delete From CFCount_Ly_Day Where UserName='"&UserName&"'"
  Conn.ExeCute Sql
 End if

 If Int(Request("daytj"))=-1 Then
  Sql="Delete From CFCount_Count_Day Where UserName='"&UserName&"'"
  Conn.ExeCute Sql
 End if

 If Int(Request("hourtj"))=-1 Then
  Sql="Delete From CFCount_Count_Hour Where UserName='"&UserName&"'"
  Conn.ExeCute Sql
 End if

 If Int(Request("backtj"))=-1 Then
  Sql="Delete From CFCount_Back Where UserName='"&UserName&"'"
  Conn.ExeCute Sql
 End if

 If Int(Request("alexatj"))=-1 Then
  Sql="Delete From CFCount_Alexa_Day Where UserName='"&UserName&"'"
  Conn.ExeCute Sql
 End if

 If Int(Request("searchtj"))=-1 Then
  Sql="Delete From CFCount_Search_Day Where UserName='"&UserName&"'"
  Conn.ExeCute Sql
 End if

 If Int(Request("keywordtj"))=-1 Then
  Sql="Delete From CFCount_SearchKeywrod_Day Where UserName='"&UserName&"'"
  Conn.ExeCute Sql
 End if

 If Int(Request("webtj"))=-1 Then
  Sql="Delete From CFCount_Web_Day Where UserName='"&UserName&"'"
  Conn.ExeCute Sql
 End if

 If Int(Request("RealShowTotal"))=-1 Then
  Sql="Update CFCount_User Set RealShowTotal=0 Where UserName='"&UserName&"'"
  Conn.ExeCute Sql
 End if

 If Int(Request("RealIpTotal"))=-1 Then
  Sql="Update CFCount_User Set RealIpTotal=0 Where UserName='"&UserName&"'"
  Conn.ExeCute Sql
 End if

 Call Alertback("����Ѿ����óɹ���",1)
End If

If Action="userdelsave" Then '�˺�ɾ��
 Pwd=Goback(ChkStr(Request("Pwd"),1),"����������ȷ��")
 Set Rs= Server.CreateObject("ADODB.Recordset")
 Sql="select * from CFCount_User where UserName='"&UserName&"'"
 Rs.Open Sql,Conn,1,1
 If Rs("Pwd")<>Md5(Pwd,1) Then Call AlertBack("������������",1)

 Filename=UserName&".jpg"
 Filename_2=UserName&".gif"

 Set Fs = Server.CreateObject("Scripting.FileSystemObject")
 If Fs.FileExists(Server.mappath("Upload/"&Filename)) Then
  Set Os = Fs.GetFile(Server.mappath("Upload/"&Filename))
  Os.Delete
 End If

 If Fs.FileExists(Server.mappath("Upload/"&Filename_2)) Then
  Set Os = Fs.GetFile(Server.mappath("Upload/"&Filename_2))
  Os.Delete
 End If

 Sql="Delete From CFCount_Alexa_Day Where UserName='"&UserName&"'"
 Conn.ExeCute Sql
 Sql="Delete From CFCount_Back Where UserName='"&UserName&"'"
 Conn.ExeCute Sql
 Sql="Delete From CFCount_Count_Day Where UserName='"&UserName&"'"
 Conn.ExeCute Sql
 Sql="Delete From CFCount_Count_Hour Where UserName='"&UserName&"'"
 Conn.ExeCute Sql
 Sql="Delete From CFCount_Ly_Day Where UserName='"&UserName&"'"
 Conn.ExeCute Sql
 Sql="Delete From CFCount_User Where UserName='"&UserName&"'"
 Conn.ExeCute Sql
 Sql="Delete From CFCount_Search_Day Where UserName='"&UserName&"'"
 Conn.ExeCute Sql
 Sql="Delete From CFCount_SearchKeywrod_Day Where UserName='"&UserName&"'"
 Conn.ExeCute Sql
 Sql="Delete From CFCount_Web_Day Where UserName='"&UserName&"'"
 Conn.ExeCute Sql

 Call AlertUrl("����˺��Ѿ���ȫɾ����","?Action=logout")
End If


If Action="gbookmodifysave" Then
 ID=ChkStr(Request("ID"),1)
 
 Content=GoBack(ChkStr(Request("Content"),1),"��������������!")
 Contact=GoBack(ChkStr(Request("Contact"),1),"��������ϵ��ʽ!")

 Set Rs=Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From CFCount_Gbook Where UserName='"&UserName&"' And ID="&ID
 Rs.Open Sql,Conn,3,2
 Rs("Content")=Content
 Rs("Contact")=Contact
 Rs.Update
 
 Call AlertBack("�޸ĳɹ�!",2)
End If

If Action="gbookdel" Then
 ID=ChkStr(Request("ID"),1)
 
 Sql="Delete From CFCount_Gbook Where UserName='"&UserName&"' And ID="&ID
 Conn.Execute Sql
 
 Call AlertBack("ɾ���ɹ�!",1)
End If

If Action="urlgo" Then
 Url=URLDecode(Request("Url"))
 Call GotoUrl(Url)
End If


If Action="userlogout" Then
 Session("CFCountUser")=""
 Session("CFCountUser_View")=""
 Response.Cookies("CFCountUserCooKie")=""
 Response.Cookies("CFCountUserCooKie").Expires=Dateadd("n",-60,Now())
 Response.Redirect "Index.asp"
End If
%>
