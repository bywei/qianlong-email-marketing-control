<%
Set Rs= Server.CreateObject("Adodb.RecordSet")
Sql="Select * From CFCount_User Where UserName='"&UserName&"'"
Rs.Open Sql,Conn,1,1
%> 

<div class="sidebar">
	<div id="leftnav1" class="menu">
<ul class="items"> 
<%If Rs("ParentName")<>"" And Session("CFCountSubAdmin")<>"" Then
 Response.write "<li><input type='submit' name='Submit' value='返回父账号' onclick=""location.href='?Action=parentgoto'""></li>"
End If
If Session("CFCountUser_View")<>"" Then
 Response.write "<li><b>独立查看者</b></li>"
End If
%>
			<li>
					<h3 id="curnavbanner">综合报告</h3>
				</li>
				<li>
					<a href=?Action=tjsurvey>统计概况</a>
				</li>
				<li>
					<a href=?Action=lastvisit>最新访问</a>
				</li>
				<li>
					<a href=?Action=lylist>来源明细</a>
				</li>
				<li>
					<a href=?Action=onlinetj>正在阅读</a>
				</li>
				<li>
					<a href=?Action=webtj>页面统计</a>
				</li>
				<li>
					<a href=?Action=urlcounttj>链接统计</a>
				</li>
				<li>
					<h3>统计报表</h3>
				</li>
				<li>
					<a href="?Action=hourtj">按小时统计</a>
				</li>
				<li>
					<a href="?Action=daytj">按天数统计</a>
				</li>
				<li>
					<a href="?Action=monthtj">按月份统计</a>
				</li>
				<li>
					<a href="?Action=backtj">回头率统计</a>
				</li>
				<li>
					<a href="?Action=alexatj">Alext条统计</a>
				</li>
				<li>
					<h3>搜索引擎统计</h3>
				</li>
				<li>
					<a href="?Action=searchtj">数量统计</a>
				</li>
				<li>
					 <a href="?Action=keywordtj">关键字统计</a>
				</li>
                <%
                  If Session("CFCountUser_View")="" Then'不是独立查看者时
                %>
                <li>
					 <h3>计数器设置</h3>
				</li>
                <li>
					 <a href="?Action=jsqset">参数设置</a>
				</li>
                <li>
					 <a href="?Action=jsqquickset">统计样式</a>
				</li>
                <li>
					 <a href="?Action=stylelist">网站统计样式</a>
				</li>
                <li>
					 <a href="?Action=imglist">邮箱统计样式</a>
				</li>
                <li>
					 <a href="?Action=jsqreset">统计重置</a>
				</li>
                
                <%End If%>
	</ul>
	</div>

	<div id="leftnav3" class="menu">
		<div class="corner"></div>
	        <ul class="items">
             <%
                If Session("CFCountUser_View")="" Then'不是独立查看者时
              %>
                <li><h3>修改个人资料</h3></li>
				<li><a href="?Action=pwdmodify">修改密码</a></li>
				<li><a href="?Action=datamodify">修改资料</a></li>
		        <li><a href="?Action=passwordanswermodify">密码保护</a></li>
                <li><a href="?Action=userdel">账号删除</a></li>
                <li><a href="?Action=gbooklist">留言列表</a></li>
                  <%End If%>
                <li><a href="?Action=getcode">统计代码</a></li>
                 <li><a href="?Action=userlogout">退出管理</a></li>
			</ul>
  </div>



</div>