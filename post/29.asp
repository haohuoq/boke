        <%ST(A)%>
			<div id="Content_ContentList" class="content-width"><a name="body" accesskey="B" href="#body"></a>
				<div class="pageContent">
					<div style="float:right;width:auto"><img border="0" src="images/Cprevious.gif" alt=""/><strong>上一篇:</strong> <a href="?id=28" accesskey=",">最简单的ASP读取数据库内容代码</a><br><img border="0" src="images/Cnext.gif" alt=""/><strong>下一篇:</strong> <a href="?id=30" accesskey=".">《初恋这件小事》片尾曲 会有那么的一天</a><br></div> 
					<img src="images/icons/3.gif" style="margin:0px 2px -4px 0px" alt=""/> <strong><a href="default.asp?cateID=4" title="查看所有【资源共享】的日志">资源共享</a></strong> <a href="feed.asp?cateID=4" target="_blank" title="订阅所有【资源共享】的日志" accesskey="O"><img border="0" src="images/rss.png" alt="订阅所有【资源共享】的日志" style="margin-bottom:-1px"/></a>
				</div>
				<div class="Content">
					<div class="Content-top"><div class="ContentLeft"></div><div class="ContentRight"></div>
					<h1 class="ContentTitle"><strong>最简单的ASP读取数据库内容代码(带简单翻页功能)</strong></h1>
					<h2 class="ContentAuthor">作者:jinbenli 日期:2011-07-29 17:39:14</h2>
				</div>
			    <div class="Content-Info">
					<div class="InfoOther">字体大小: <a href="javascript:SetFont('12px')" accesskey="1">小</a> <a href="javascript:SetFont('14px')" accesskey="2">中</a> <a href="javascript:SetFont('16px')" accesskey="3">大</a></div>
					<div class="InfoAuthor"><img src="images/weather/hn2_sunny.gif" style="margin:0px 2px -6px 0px" alt=""/><img src="images/weather/hn2_t_sunny.gif" alt=""/> <img src="images/level3.gif" style="margin:0px 2px -1px 0px" alt=""/><$EditAndDel$></div>
				</div>
				<div id="logPanel" class="Content-body">
					&lt;!-- #include file=&#34;conn.asp&#34;--&gt; <br/>&lt;!--翻页代码--&gt;<br/>&lt;% dim ThisURL,ThisPage&nbsp;&nbsp; &#39;定义变量<br/>ThisURL=&#34;<a href="http://" target="_blank" rel="external">http://</a>&#34;&amp;request.ServerVariables(&#34;SERVER_NAME&#34;)&amp;request.ServerVariables(&#34;URL&#34;) &#39;取得当前页URL<br/><br/>if not isempty(request(&#34;page&#34;)) then&nbsp;&nbsp; &#39;如果 传递过来的 page 值为空值 则<br/>thispage = request(&#34;page&#34;)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &#39;thispage 的值就是page的值&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <br/>else&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &#39;否则&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <br/>thispage = 1&nbsp;&nbsp;&nbsp;&nbsp; &#39;thispage 的值为 1&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br/>end if&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &#39;退出<br/><br/>&#39;创建记录集&#160;&#160;&#160;&#160;&nbsp;&nbsp;<br/>Set Rs=Server.Cr&#101;ateObject(&#34;Adodb.RecordSet&#34;)<br/>Sql=&#34;Sel&#101;ct * From [news] &#34;&nbsp;&nbsp; <br/>Rs.Open Sql,Conn,1,1&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &#39;打开news表<br/><br/>rs.pagesize = 5&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&#39;每页显示5条记录<br/>rs.absolutepage = thispage&nbsp;&nbsp; &#39;将thispage 转换成rs.absolutepage<br/><br/>for ipage=1 to rs.pagesize&nbsp;&nbsp; &#39;循环从1到 rs.pagesize ,前面我们定义了rs.pagesize=5&nbsp;&nbsp; <br/>if rs.eof then exit for&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&#39;如果表中没有记录 则退出循环,否则运行下面代码<br/>%&gt;<br/><br/>&lt;!--读取内容--&gt;<br/>&lt;ul&gt;&lt;li&gt;&lt;%=rs(&#34;title&#34;)%&gt;&lt;/li&gt;&lt;/ul&gt;<br/><br/><br/>&lt;%<br/>&#39;<a href="http://jinbenli.com" target="_blank">jinbenli</a>.com<br/>rs.movenext<br/>next<br/>%&gt;<br/><br/>&lt;!--翻页--&gt;<br/>页次:&lt;%=thispage%&gt;/&lt;%=rs.pagecount%&gt;&nbsp;&nbsp; <br/>&lt;%for i=1 to rs.pagecount%&gt;<br/>&lt;a href=&#34;&lt;%=thisURL%&gt;?page=&lt;%=i%&gt;&#34; target=&#34;_self&#34;&gt;[&lt;%=i%&gt;]&lt;/a&gt; <br/>&lt;%next%&gt;<br/>共&lt;%=rs.recordcount%&gt;条信息<br/><br/>&lt;!--关闭记录集--&gt;<br/>&lt;%<br/>Rs.Close<br/>Set Rs=Nothing<br/>%&gt;
					<br/><br/>
				</div>
				<div class="Content-body">
					
					<img src="images/From.gif" style="margin:0px 2px -4px 0px" alt=""/><strong>文章来自:</strong> <a href="http://jinbenli.com/" target="_blank">本站原创</a><br/>
					<img src="images/icon_trackback.gif" style="margin:4px 2px -4px 0px" alt=""/><strong>引用通告:</strong> <a href="trackback.asp?tbID=29&amp;action=view" target="_blank">查看所有引用</a> | <a href="javascript:;" title="获得引用文章的链接" onclick="getTrackbackURL(29)">我要引用此文章</a><br/>
					<img src="images/tag.gif" style="margin:4px 2px -4px 0px" alt=""/><strong>Tags:</strong> <a href="default.asp?tag=asp%E4%BB%A3%E7%A0%81">asp代码</a><a href="http://technorati.com/tag/asp代码" rel="tag" style="display:none">asp代码</a> <br/>
					<img src="images/notify.gif" style="margin:4px 2px -4px 0px" alt=""/><strong>相关日志:</strong>
                    <div class="Content-body" id="related_tag" style="margin-left:25px"></div>
                    <script language="javascript" type="text/javascript">Related(29, 1, "related_tag");</script>
				</div>
				<div class="Content-bottom"><div class="ContentBLeft"></div><div class="ContentBRight"></div>评论: 0 | <a href="trackback.asp?tbID=29&amp;action=view" target="_blank">引用: 0</a> | 查看次数: <$log_ViewNums$></div>
			</div>
		</div>
