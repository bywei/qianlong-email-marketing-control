
<!--#include FILE="Conn.asp"-->
<!--#Include File="CF_MyFunCtion.asp"-->
<!--#include FILE="CF_UpFile_Class.asp"-->
<%If Request("Action")="" Then%>
<body leftmargin="0" topmargin="0">
<form name="form1" method="post" action="?Action=upsave" enctype="multipart/form-data">
<input type=file name="img">
<input type=hidden name="UserType" value="<%=Request("UserType")%>">
<input type=hidden name="CheckCode" value="<%=Request("CheckCode")%>">
<input type=submit name="submit" value="上传">
</form>
<%End If%>


<%If Request("Action")="upsave" Then%>
<%

 Const UpFileType="jpg|gif" '允许的上传文件类型
 Const MaxFileSize="50" '允许的最大上传文件多少K
 FormName="img"'上传的表单名称

 If Right(SavePath,1)<>"/" Then SavePath=SavePath&"/" '在目录后加(/) 

 set upfile=new upfile_class ''建立上传对象
 upfile.NoAllowExt="asp;exe;htm;html;aspx;cs;vb;js;"	'设置上传类型的黑名单
 upfile.GetData (10240000)   '取得上传数据,限制最大上传10M
 UserType = CInt(upfile.form("UserType"))
 CheckCode = Replace(upfile.form("CheckCode"),"'","''")'过滤单引号


 Set File=UpFile.File(FormName)    
 If file.filesize<100 Then
  Response.Write "<script language='javascript'>" & VbCRlf
  Response.Write "alert('请先选择你要上传的文件！');" & VbCrlf
  Response.Write "history.go(-1);" & vbCrlf
  Response.Write "</script>" & VbCRLF
  Response.End
 End If
 If File.filesize>(MaxFileSize*1024) Then
  Response.Write "<script language='javascript'>" & VbCRlf
  Response.Write "alert('上传的文件大小超过限制！');" & VbCrlf
  Response.Write "history.go(-1);" & vbCrlf
  Response.Write "</script>" & VbCRLF
  Response.End
 End If
		
 FileExt=Lcase(File.FileExt)
 FileSizeNum=File.filesize
 ForumUpload=Split(UpFileType,"|")
 For I=0 To Ubound(ForumUpload)
  If fileEXT=Trim(Forumupload(i)) Then
   EnableUpload=True
   Exit For
  End If
 Next
 If FileExt="asp" Or FileExt="asa" or FileExt="aspx" Then EnableUpload=False
 If EnableUpload=False Then
  Response.Write "<script language='javascript'>" & VbCRlf
  Response.Write "alert('对不起,不支持这类文件上传!');" & VbCrlf
  Response.Write "history.go(-1);" & vbCrlf
  Response.Write "</script>" & VbCRLF
  Response.End
 End If	


 If UserType=0 Then'管理员上传时
  FileName="sys-img-"&DateDiff("s","2000-1-1",Now)
  FileName_2=FileName&"."&FileExt 
  FilePath="upload/"&FileName_2  

  If RsSet("AdminCode")<>CheckCode Then
   Response.write "验证码有误"
   Response.End
  End If


  UpFile.SaveToFile FormName,Server.MapPath(FilePath)
  Set UpFile=nothing
 
  set rs=server.createobject("adodb.recordset")
  sql="select Top 1 * from CFCount_Upfile"
  rs.open sql,conn,3,2
  rs.addnew
  rs("FileSizeNum")=FileSizeNum
  rs("FileName")=FileName_2
  rs.update  

  Response.Write "<script language='javascript'>" & VbCRlf
  Response.Write "alert('上传成功!');" & VbCrlf
  Response.Write "top.location.href = 'CF_Admin_Manage.asp?Action=imglist';" & vbCrlf
  Response.Write "</script>" & VbCRLF  
 ElseIf UserType=1 Then
     Set Rs=Server.CreateObject("Adodb.RecordSet")
     Sql="Select * From CFCount_User where UserCode='"&CheckCode&"'"
     Rs.Open Sql,Conn,3,2
	 
	 If Rs.Eof Then
	  Response.write "验证码有误或没有上传权限"
	  Response.End
	 End If
     
	 FileName=Rs("UserName")
	 FileName_2=FileName&"."&FileExt
	 FilePath="upload/"&FileName_2  

    
     UpFile.SaveToFile FormName,Server.MapPath(FilePath)
     Set UpFile=nothing
	 
     Rs("ImgFileName")=FileName_2
     Rs.Update
  
 
  Response.Write "<script language='javascript'>" & VbCRlf
  Response.Write "alert('上传成功!');" & VbCrlf
  Response.Write "top.location.href = 'Manage.asp?Action=imglist';" & vbCrlf
  Response.Write "</script>" & VbCRLF
 End If
%>
<%End If%>