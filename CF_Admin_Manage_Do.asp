
<%
If Action="lyreset" Then
 Sql="Update CFCount_Admin Set Store_TotalLy=0"
 Conn.ExeCute Sql
 Call AlertBack("��ԭ�ɹ�",1)
End If

If Action="beforelydel" Then
 Sql="Delete From CFCount_Ly_Day Where DateDiff('d',AddDate,Now())>0"
 Conn.ExeCute Sql
 Call AlertBack("ɾ���ɹ�",1)
End If

If Action="beforewebdel" Then
 Sql="Delete From CFCount_Web_Day Where DateDiff('d',AddDate,Now())>0"
 Conn.ExeCute Sql
 Call AlertBack("ɾ���ɹ�",1)
End If

If Action="userdel" Then
 UserName=ChkStr(Request("UserName"),1)

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

 Call AlertBack("ɾ���ɹ�!",1)
End If

If Action="usermodifysave" Then
 ID=ChkStr(Request("ID"),2)
 Pwd_New=ChkStr(Request("Pwd_New"),1)
 PasswordAnswer_New=ChkStr(Request("PasswordAnswer_New"),1)
 UserState=ChkStr(Request("UserState"),2)
 AdminDesc=ChkStr(Request("AdminDesc"),1)


 Set Rs=Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From CFCount_User Where ID="&ID
 Rs.Open Sql,Conn,3,2
 If Pwd_New<>"" Then Rs("Pwd")=Md5(Pwd_New,1)
 If PasswordAnswer_New<>"" Then
  Rs("PassWordAsk")="��Ĵ���?"
  Rs("PasswordAnswer")=Md5(PasswordAnswer_New,1)
 End If
 Rs("UserState")=UserState
 Rs("AdminDesc")=AdminDesc
 Rs.Update

 Call AlertBack("�޸��û����ϳɹ�!",2)
End If

If Action="syssetmodifysave" Then
 SysTitle=GoBack(ChkStr(Request("SysTitle"),1),"���������������!")
 SysTitle=Replace(SysTitle,"""","")'�滻��˫����
 SysTitle=Replace(SysTitle,"'","")'�滻��������
 
 TjTextName=GoBack(Left(ChkStr(Request("TjTextName"),1),8),"���������������!")
 TjTextName=Replace(TjTextName,"""","")'�滻��˫����
 TjTextName=Replace(TjTextName,"'","")'�滻��������
 
 RegState=ChkStr(Request("RegState"),2)
 LyKeep=ChkStr(Request("LyKeep"),2)
 OnlineKeep=ChkStr(Request("OnlineKeep"),2)
 WebKeep=ChkStr(Request("WebKeep"),2)
 StyleTotal=GoBack(ChkStr(Request("StyleTotal"),2),"������ͼƬ��ʽ����!")
 LogoUrl=GoBack(ChkStr(Request("LogoUrl"),1),"�����������ϵͳ�ϵ�LogoͼƬ��ַ!")
 IpArea=ChkStr(Request("IpArea"),2)
 SkinType=ChkStr(Request("SkinType"),2)
 SkinColor=ChkStr(Request("SkinColor"),1)
 emailsendtype=ChkStr(Request("emailsendtype"),2)
 SmtpAddress=ChkStr(Request("SmtpAddress"),1)
 SmtpUser=ChkStr(Request("SmtpUser"),1)
 SmtpPassword=ChkStr(Request("SmtpPassword"),1)
 


 Set Rs=Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From CFCount_Admin"
 Rs.Open Sql,conn,3,2
 Rs("SysTitle")=SysTitle
 Rs("TjTextName")=TjTextName
 Rs("RegState")=RegState
 Rs("LyKeep")=LyKeep
 Rs("OnlineKeep")=OnlineKeep
 Rs("WebKeep")=WebKeep
 Rs("StyleTotal")=StyleTotal
 Rs("LogoUrl")=LogoUrl
 Rs("IpArea")=IpArea
 Rs("SkinType")=SkinType
 Rs("SkinColor")=SkinColor
 Rs("emailsendtype")=emailsendtype
 Rs("SmtpAddress")=SmtpAddress
 Rs("SmtpUser")=SmtpUser
 Rs("SmtpPassword")=SmtpPassword
 Rs.Update

 Call AlertBack("�޸Ĺ���Ա���ϳɹ�!",1)
End if

If Action="othersetmodifysave" Then

 MyStr="SysCode="&Goback(ChkStr(Request("SysCode"),1),"������ϵͳ��ȫ��")&"|||"
 MyStr=MyStr&"HourCountKeepDay="&Goback(ChkStr(Request("HourCountKeepDay"),1),"������Сʱͳ�Ʊ�������")&"|||"
 MyStr=MyStr&"WebKeepDay="&Goback(ChkStr(Request("WebKeepDay"),1),"��������ҳ�����¼��������")&"|||"
 MyStr=MyStr&"LyKeepDay="&Goback(ChkStr(Request("LyKeepDay"),1),"��������Դ��¼��������")&"|||"
 MyStr=MyStr&"SearchKeepDay="&Goback(ChkStr(Request("SearchKeepDay"),1),"���������������¼��������")&"|||"
 MyStr=MyStr&"KeywordKeepDay="&Goback(ChkStr(Request("KeywordKeepDay"),1),"������KeywordKeepDay")&"|||"
 MyStr=MyStr&"CountKeepDay="&Goback(ChkStr(Request("CountKeepDay"),1),"������������¼��������")&"|||"

 Sql="Update CFCount_Admin set OtherSet='"&Mystr&"'"
 Conn.Execute Sql
 
 Application("CFCount_MySet") = Empty'���Asp�������

 Call AlertBack("�޸ĳɹ�",1)
End If

If Action="searchaddsave" Then
 SiteDesc=GoBack(ChkStr(Lcase(Request("SiteDesc")),1),"��������������������")
 SiteFlag=GoBack(ChkStr(Lcase(Request("SiteFlag")),1),"��������������Ӣ�ı�ʶ")
 KeywordFlag=GoBack(ChkStr(Lcase(Request("KeywordFlag")),1),"�����������ؼ��ֵĲ���")
 SiteFlag=Replace(SiteFlag,",")
 SiteFlag=Replace(SiteFlag,"|")
 KeywordFlag=Replace(KeywordFlag,",")
 KeywordFlag=Replace(KeywordFlag,"|")


 Set Rs=Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From CFCount_Search_Set Where SiteFlag='"&SiteFlag&"'"
 Rs.Open Sql,Conn,3,2
 If Not.Rs.Eof Then
  Call AlertBack("�Ѿ�����,��������!",1)
 Else
  Rs.AddNew
  Rs("SiteDesc")=SiteDesc
  Rs("SiteFlag")=SiteFlag
  Rs("KeywordFlag")=KeywordFlag
  Rs.Update
 End If

 Call AlertBack("����ɹ�!",1)
End If


If Action="searchmodifysave" Then
 ID=ChkStr(Request("ID"),2)
 SiteDesc=GoBack(ChkStr(Lcase(Request("SiteDesc")),1),"��������������������")
 SiteFlag=GoBack(ChkStr(Lcase(Request("SiteFlag")),1),"��������������Ӣ�ı�ʶ")
 KeywordFlag=GoBack(ChkStr(Lcase(Request("KeywordFlag")),1),"�����������ؼ��ֵĲ���")

 Sql="Select Count(*) From CFCount_Search_Set Where SiteFlag='"&SiteFlag&"' And ID<>"&ID
 Set Rs=Server.CreateObject("Adodb.RecordSet")
 Rs.Open Sql,Conn,1,1
 If Rs(0)>0 Then Call AlertBack("�Ѿ�����,��������!",1)

 Set Rs=Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From CFCount_Search_Set Where ID="&ID
 Rs.Open Sql,Conn,3,2
 Rs("SiteDesc")=SiteDesc
 Rs("SiteFlag")=SiteFlag
 Rs("KeywordFlag")=KeywordFlag
 Rs.Update

 Call AlertBack("�޸ĳɹ�!",2)
End If

If Action="searchdel" Then
 SiteFlag=ChkStr(Request("SiteFlag"),1)
 Sql="Delete From CFCount_Search_Set Where SiteFlag='"&SiteFlag&"'"
 Conn.ExeCute Sql
 Sql="Delete From CFCount_Search_Day Where SiteFlag='"&SiteFlag&"'"
 Conn.ExeCute Sql

 Call AlertBack("ɾ���ɹ�",1)
End If  '�Զ���ɾ��

If Action="picdel" Then
 Filename=ChkStr(Request("Filename"),1)

 Set Fs = Server.CreateObject("Scripting.FileSystemObject")
 If Fs.FileExists(Server.mappath("Upload/"&Filename)) Then
  Set Os = Fs.GetFile(Server.mappath("Upload/"&Filename))
  Os.Delete
 End If

 Set Rs = Server.CreateObject("ADODB.Recordset")
 Sql="delete from CFCount_Upfile where Filename='"&Filename&"'"
 Rs.open sql,conn,3,2

 Call AlertBack("�ɹ�ɾ��!",1)
End if

If Action="pwdmodifysave" Then
 Pwd_Old=GoBack(ChkStr(Trim(Request("Pwd_Old")),1),"������ԭ����Ա����")
 Admin=GoBack(ChkStr(Trim(Request("Admin")),1),"���������Ա����")
 Pwd=GoBack(ChkStr(Trim(Request("Pwd")),1),"����������")
 Pwd2=GoBack(ChkStr(Trim(Request("Pwd2")),1),"�������ظ�����")

 If Pwd<>Pwd2 Then Call AlertBack("��������벻һ�£�����������һ��!",1)
 
 Set Rs= Server.CreateObject("Adodb.RecordSet")
 Sql="Select * From CFCount_Admin"
 Rs.Open Sql,Conn,3,2
 If Md5(Pwd_Old,1)<>Rs("Pwd") Then Call AlertBack("�����������д���",1)

 Rs("Admin")=Admin
 Rs("Pwd")=Md5(Pwd,1)
 Rs.Update

 Call AlertBack("�޸ĳɹ�",1)
End If

If Action="alllydel" Then
 Set Rs= Server.CreateObject("Adodb.RecordSet")
 Sql="Delete From CFCount_Ly_Day"
 Rs.Open Sql,Conn,3,2
 Call AlertBack("ɾ���ɹ�!",1)
End if

If Action="adminlogout" Then
 Session("CFCountAdmin")=""
 Response.Cookies("CFCountAdminCookie")=""
 Response.Redirect "Admin.asp"
End If

If Action="usergo" Then
 Session("CFCountUser")=Request("UserName")	
 Response.Redirect("Manage.asp")
End If

If Action="sqlexesave" Then
 If AdminOnlineExecSql=0 Then Call AlertBack("Conn.asp��������ϵͳ��ֹ����ִ��Sql���",1)
 
 Pwd=GoBack(ChkStr(Trim(Request("Pwd")),1),"������ȷ������")

 Set Rs= Server.CreateObject("Adodb.RecordSet")
 Sql="Select Pwd From CFCount_Admin"
 Rs.Open Sql,Conn,3,2
 If Md5(Pwd,1)<>Rs("Pwd") Then Call AlertBack("ȷ�����������д���",1)
 
 Sql = Request.Form("Sql")
 Conn.Execute(Sql)
  
 Call AlertBack("ִ�гɹ�",1)
End If

If Action="dbys" Then
 YsDbName="backup_"&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&"_"&Md5(timer(),1)&".mdb"
 TgDbName="backup_"&year(now)&month(now)&day(now)&hour(now)&minute(now)&(second(now)+1)&"_"&Md5(timer(),1)&".mdb"

 Set Fs = Server.CreateObject("Scripting.FileSystemObject")
 Fs.copyfile Server.MapPath(DbPath),Server.MapPath("data\"&YsDbName)
 
  Const JET_3X = 4
  Set Engine = CreateObject("JRO.JetEngine")
  on error resume next
  Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("data\"&YsDbName) , "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("data\"&TgDbName)

  Fs.CopyFile Server.MapPath("data\"&TgDbName),Server.MapPath(DbPath)
  
 Call AlertBack("ѹ���ɹ�",1)
End If

If Action="dbbackup" Then
 Set Fs = Server.CreateObject("Scripting.FileSystemObject")
 Fs.copyfile Server.MapPath(DbPath),Server.MapPath("data\backup_"&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&"_"&Md5(timer(),1)&".mdb")
 Call AlertBack("�������ݿ�ɹ�!",1)
End If '���ݿⱸ�ݽ���


If Action="dbdel" Then
 DbMc=ChkStr(Request("DbMc"),1)
 
If Lcase(StrReverse(Left(StrReverse(DbMc),4)))<>".mdb" Then Call Alertback("ֻ��ɾ����չ��Ϊmdb�����ݿ��ļ�",1)

 Set Fs = Server.CreateObject("Scripting.FileSystemObject")
 Set Os = Fs.GetFile(Server.mappath("data/"&DbMc))
 Os.Delete


 Call AlertBack("�ɹ�ɾ��!",1)
End If
%>