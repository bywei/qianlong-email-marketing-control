
<%
DbPath="data/cfcount.mdb" '���ݿ��������ã�ʹ��Ĭ��ֵ�Ļ�����Ҫ���������SysCode�Ա�֤��ȫ,��ð�qqcf_data���Ĭ�����ݿ����ƻ�һ�����ƣ����ú��������д��


On Error Resume Next
ConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(DbPath)
Set Conn = Server.CreateObject ("Adodb.Connection")
Conn.Open ConnStr

If Err Then
 err.Clear
 Set Conn = Nothing
 Response.Write "���ݿ����ӳ������������ִ���"
 Response.End
End If

InjGuard=0 '�Ƿ���ͨ�÷�ע�룬0Ϊ��������1Ϊ������ע�������Ѿ�������ע�봦����,����һ�����ùرվͿ�����
AdminOnlineExecSql=0 '�Ƿ�����������Ա��̨���ݿ����������ִ��Sql��0Ϊ��1Ϊ�ǣ�Ĭ��Ϊ��

If InjGuard=1 Then
SQL_injdata = "'|'"
SQL_inj = split(SQL_Injdata,"|") 

If Request.QueryString<>"" Then 
 For Each SQL_Get In Request.QueryString 
  For SQL_Data=0 To Ubound(SQL_inj) 
   if instr(Lcase(Request.QueryString(SQL_Get)),Sql_Inj(Sql_DATA))>0 Then 
    Response.Write "��ַ�а����Ƿ��ַ�"&Sql_Inj(Sql_DATA)
    Response.End 
   End If
  Next
 Next
End If

If Request.Form<>"" Then 
 For Each Sql_Post In Request.Form 
  For SQL_Data=0 To Ubound(SQL_inj) 
   if instr(Lcase(Request.Form(Sql_Post)),Sql_Inj(Sql_DATA))>0 Then 
    Response.Write "���а����Ƿ��ַ�"&Sql_Inj(Sql_DATA)
    Response.end 
   End If 
  Next 
 Next 
End If

If Request.Cookies<>"" Then
 For Each Sql_Cookies In Request.Cookies
  For SQL_Data=0 To Ubound(SQL_inj)
   if instr(Lcase(Request.Cookies(Sql_Cookies)),Sql_Inj(Sql_DATA))>0 Then
    Response.Write "Cookies�а����Ƿ��ַ�"&Sql_Inj(Sql_DATA)&",���� IEѡ��-���� ��ɾ��Cookiess��ִ�в���!"
    Response.end
   End If
  Next
 Next
End If
End If

Set RsSet=Server.CreateObject("Adodb.RecordSet")
Sql="Select * From CFCount_Admin"
RsSet.Open Sql,Conn,1,1
%>