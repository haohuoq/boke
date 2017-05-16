        <%ST(A)%>
			<div id="Content_ContentList" class="content-width"><a name="body" accesskey="B" href="#body"></a>
				<div class="pageContent">
					<div style="float:right;width:auto"><img border="0" src="images/Cprevious.gif" alt=""/><strong>上一篇:</strong> <a href="?id=27" accesskey=",">最简单的ASP连接数据库代码</a><br><img border="0" src="images/Cnext.gif" alt=""/><strong>下一篇:</strong> <a href="?id=29" accesskey=".">最简单的ASP读取数据库内容代码(带简单翻页功能)</a><br></div> 
					<img src="images/icons/3.gif" style="margin:0px 2px -4px 0px" alt=""/> <strong><a href="default.asp?cateID=4" title="查看所有【资源共享】的日志">资源共享</a></strong> <a href="feed.asp?cateID=4" target="_blank" title="订阅所有【资源共享】的日志" accesskey="O"><img border="0" src="images/rss.png" alt="订阅所有【资源共享】的日志" style="margin-bottom:-1px"/></a>
				</div>
				<div class="Content">
					<div class="Content-top"><div class="ContentLeft"></div><div class="ContentRight"></div>
					<h1 class="ContentTitle"><strong>最简单的ASP读取数据库内容代码</strong></h1>
					<h2 class="ContentAuthor">作者:jinbenli 日期:2011-07-29 17:29:05</h2>
				</div>
			    <div class="Content-Info">
					<div class="InfoOther">字体大小: <a href="javascript:SetFont('12px')" accesskey="1">小</a> <a href="javascript:SetFont('14px')" accesskey="2">中</a> <a href="javascript:SetFont('16px')" accesskey="3">大</a></div>
					<div class="InfoAuthor"><img src="images/weather/hn2_sunny.gif" style="margin:0px 2px -6px 0px" alt=""/><img src="images/weather/hn2_t_sunny.gif" alt=""/> <img src="images/level3.gif" style="margin:0px 2px -1px 0px" alt=""/><$EditAndDel$></div>
				</div>
				<div id="logPanel" class="Content-body">
					&lt;!-- #include file=&#34;conn.asp&#34;--&gt;&nbsp;&nbsp;<br/>&lt;% &#39;<a href="http://jinbenli.com" target="_blank">jinbenli</a>.com<br/>Set Rs=Server.Cr&#101;ateObject(&#34;Adodb.RecordSet&#34;)<br/>Sql=&#34;Sel&#101;ct&nbsp;&nbsp;* From [news]&#34;<br/>Rs.Open Sql,Conn,1,1<br/>If Not (Rs.Eof o&#114; Rs.Bof) Then<br/>Do While Not Rs.Eof <br/>%&gt;<br/><br/>内容列表：<br/>&lt;ul&gt;&lt;li&gt;&lt;%=rs(&#34;title&#34;)%&gt;&lt;/li&gt;&lt;/ul&gt;<br/><br/>&lt;% <br/>&#39;循环输出内容 <a href="http://jinbenli.com" target="_blank">jinbenli</a>.com<br/>rs.movenext<br/>loop<br/>end if&#160;&#160;&#160;&#160;<br/>rs.close<br/>set rs=nothing<br/>%&gt;<br/>
					<br/><br/>
				</div>
				<div class="Content-body">
					
					<img src="images/From.gif" style="margin:0px 2px -4px 0px" alt=""/><strong>文章来自:</strong> <a href="http://jinbenli.com/" target="_blank">本站原创</a><br/>
					<img src="images/icon_trackback.gif" style="margin:4px 2px -4px 0px" alt=""/><strong>引用通告:</strong> <a href="trackback.asp?tbID=28&amp;action=view" target="_blank">查看所有引用</a> | <a href="javascript:;" title="获得引用文章的链接" onclick="getTrackbackURL(28)">我要引用此文章</a><br/>
					<img src="images/tag.gif" style="margin:4px 2px -4px 0px" alt=""/><strong>Tags:</strong> <a href="default.asp?tag=asp%E4%BB%A3%E7%A0%81">asp代码</a><a href="http://technorati.com/tag/asp代码" rel="tag" style="display:none">asp代码</a> <br/>
					<img src="images/notify.gif" style="margin:4px 2px -4px 0px" alt=""/><strong>相关日志:</strong>
                    <div class="Content-body" id="related_tag" style="margin-left:25px"></div>
                    <script language="javascript" type="text/javascript">Related(28, 1, "related_tag");</script>
				</div>
				<div class="Content-bottom"><div class="ContentBLeft"></div><div class="ContentBRight"></div>评论: 0 | <a href="trackback.asp?tbID=28&amp;action=view" target="_blank">引用: 0</a> | 查看次数: <$log_ViewNums$></div>
			</div>
		</div>
