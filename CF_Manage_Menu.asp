<%
Set Rs= Server.CreateObject("Adodb.RecordSet")
Sql="Select * From CFCount_User Where UserName='"&UserName&"'"
Rs.Open Sql,Conn,1,1
%> 

<div class="sidebar">
	<div id="leftnav1" class="menu">
<ul class="items"> 
<%If Rs("ParentName")<>"" And Session("CFCountSubAdmin")<>"" Then
 Response.write "<li><input type='submit' name='Submit' value='���ظ��˺�' onclick=""location.href='?Action=parentgoto'""></li>"
End If
If Session("CFCountUser_View")<>"" Then
 Response.write "<li><b>�����鿴��</b></li>"
End If
%>
			<li>
					<h3 id="curnavbanner">�ۺϱ���</h3>
				</li>
				<li>
					<a href=?Action=tjsurvey>ͳ�Ƹſ�</a>
				</li>
				<li>
					<a href=?Action=lastvisit>���·���</a>
				</li>
				<li>
					<a href=?Action=lylist>��Դ��ϸ</a>
				</li>
				<li>
					<a href=?Action=onlinetj>�����Ķ�</a>
				</li>
				<li>
					<a href=?Action=webtj>ҳ��ͳ��</a>
				</li>
				<li>
					<a href=?Action=urlcounttj>����ͳ��</a>
				</li>
				<li>
					<h3>ͳ�Ʊ���</h3>
				</li>
				<li>
					<a href="?Action=hourtj">��Сʱͳ��</a>
				</li>
				<li>
					<a href="?Action=daytj">������ͳ��</a>
				</li>
				<li>
					<a href="?Action=monthtj">���·�ͳ��</a>
				</li>
				<li>
					<a href="?Action=backtj">��ͷ��ͳ��</a>
				</li>
				<li>
					<a href="?Action=alexatj">Alext��ͳ��</a>
				</li>
				<li>
					<h3>��������ͳ��</h3>
				</li>
				<li>
					<a href="?Action=searchtj">����ͳ��</a>
				</li>
				<li>
					 <a href="?Action=keywordtj">�ؼ���ͳ��</a>
				</li>
                <%
                  If Session("CFCountUser_View")="" Then'���Ƕ����鿴��ʱ
                %>
                <li>
					 <h3>����������</h3>
				</li>
                <li>
					 <a href="?Action=jsqset">��������</a>
				</li>
                <li>
					 <a href="?Action=jsqquickset">ͳ����ʽ</a>
				</li>
                <li>
					 <a href="?Action=stylelist">��վͳ����ʽ</a>
				</li>
                <li>
					 <a href="?Action=imglist">����ͳ����ʽ</a>
				</li>
                <li>
					 <a href="?Action=jsqreset">ͳ������</a>
				</li>
                
                <%End If%>
	</ul>
	</div>

	<div id="leftnav3" class="menu">
		<div class="corner"></div>
	        <ul class="items">
             <%
                If Session("CFCountUser_View")="" Then'���Ƕ����鿴��ʱ
              %>
                <li><h3>�޸ĸ�������</h3></li>
				<li><a href="?Action=pwdmodify">�޸�����</a></li>
				<li><a href="?Action=datamodify">�޸�����</a></li>
		        <li><a href="?Action=passwordanswermodify">���뱣��</a></li>
                <li><a href="?Action=userdel">�˺�ɾ��</a></li>
                <li><a href="?Action=gbooklist">�����б�</a></li>
                  <%End If%>
                <li><a href="?Action=getcode">ͳ�ƴ���</a></li>
                 <li><a href="?Action=userlogout">�˳�����</a></li>
			</ul>
  </div>



</div>