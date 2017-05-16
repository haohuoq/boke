        <%ST(A)%>
			<div id="Content_ContentList" class="content-width"><a name="body" accesskey="B" href="#body"></a>
				<div class="pageContent">
					<div style="float:right;width:auto"><img border="0" src="images/Cprevious.gif" alt=""/><strong>上一篇:</strong> <a href="?id=25" accesskey=",">最简单的ASP生成静态文件代码</a><br><img border="0" src="images/Cnext.gif" alt=""/><strong>下一篇:</strong> <a href="?id=27" accesskey=".">最简单的ASP连接数据库代码</a><br></div> 
					<img src="images/icons/3.gif" style="margin:0px 2px -4px 0px" alt=""/> <strong><a href="default.asp?cateID=4" title="查看所有【资源共享】的日志">资源共享</a></strong> <a href="feed.asp?cateID=4" target="_blank" title="订阅所有【资源共享】的日志" accesskey="O"><img border="0" src="images/rss.png" alt="订阅所有【资源共享】的日志" style="margin-bottom:-1px"/></a>
				</div>
				<div class="Content">
					<div class="Content-top"><div class="ContentLeft"></div><div class="ContentRight"></div>
					<h1 class="ContentTitle"><strong>最简单的ASP生成静态文件代码(带静态生成模版)</strong></h1>
					<h2 class="ContentAuthor">作者:jinbenli 日期:2011-07-28 15:59:46</h2>
				</div>
			    <div class="Content-Info">
					<div class="InfoOther">字体大小: <a href="javascript:SetFont('12px')" accesskey="1">小</a> <a href="javascript:SetFont('14px')" accesskey="2">中</a> <a href="javascript:SetFont('16px')" accesskey="3">大</a></div>
					<div class="InfoAuthor"><img src="images/weather/hn2_sunny.gif" style="margin:0px 2px -6px 0px" alt=""/><img src="images/weather/hn2_t_sunny.gif" alt=""/> <img src="images/level3.gif" style="margin:0px 2px -1px 0px" alt=""/><$EditAndDel$></div>
				</div>
				<div id="logPanel" class="Content-body">
					先建立一个静态模版文件template.htm,内容如下：<br/>&lt;html&gt;<br/>&lt;title&gt;静态输出测试！&lt;/title&gt;&nbsp;&nbsp;<br/>&lt;body&gt; <br/>$title$&lt;br/&gt;<br/>$Content$&nbsp;&nbsp;<br/>&lt;/body&gt;&nbsp;&nbsp;<br/>&lt;/html&gt;<br/><br/>生成代码文件test.asp:<br/>&lt;%&nbsp;&nbsp;<br/>If Request.Form(&#34;Content&#34;)&lt;&gt;&#34;&#34; Then<br/>Dim fso,htmlwrite&nbsp;&nbsp;<br/>Dim strTitle,strContent,strOut <br/>strTitle=Request.Form(&#34;Title&#34;)<br/>strContent=Request.Form(&#34;Content&#34;) <br/>&#39;// 创建文件系统对象&nbsp;&nbsp;<br/>Set fso=Server.Cr&#101;ateObject(&#34;Scripting.FileSystemObject&#34;)&nbsp;&nbsp;<br/>&#39;// 打开网页模板文件，读取模板内容&nbsp;&nbsp;<br/>Set htmlwrite=fso.OpenTextFile(Server.MapPath(&#34;Template.htm&#34;))&nbsp;&nbsp;<br/>strOut=htmlwrite.ReadAll&nbsp;&nbsp;<br/>htmlwrite.close&nbsp;&nbsp;<br/><br/>&#39;// 用真实内容替换模板中的标记&nbsp;&nbsp;<br/>strOut=Replace(strOut,&#34;$title$&#34;,strTitle)&nbsp;&nbsp;<br/>strOut=Replace(strOut,&#34;$Content$&#34;,strContent)&nbsp;&nbsp;<br/>&#39;// 创建要生成的静态页&nbsp;&nbsp;<br/>Set htmlwrite=fso.Cr&#101;ateTextFile(Server.MapPath(&#34;test.htm&#34;),true)&nbsp;&nbsp;<br/>&#39;// 写入网页内容&nbsp;&nbsp;<br/>htmlwrite.WriteLine strOut&nbsp;&nbsp;<br/>htmlwrite.close&nbsp;&nbsp;<br/>Response.Write &#34;生成静态页成功！&#34;&nbsp;&nbsp;<br/>&#39;// 释放文件系统对象&nbsp;&nbsp;<br/>set htmlwrite=Nothing&nbsp;&nbsp;<br/>set fso=Nothing&nbsp;&nbsp;<br/>end if<br/>%&gt; <br/>&lt;form name=&#34;form1&#34; method=&#34;post&#34; action=&#34;&#34;&gt; <br/>&lt;input name=&#34;Title&#34; id=&#34;Title&#34; size=30&gt; <br/>&lt;br&gt; <br/>&lt;br&gt; <br/>&lt;textarea name=&#34;Content&#34; cols=&#34;50&#34; rows=&#34;8&#34;&gt;&lt;/textarea&gt; <br/>&lt;br&gt; <br/>&lt;br&gt; <br/>&lt;input type=&#34;submit&#34; name=&#34;Submit&#34; value=&#34;生成HTML页&#34;&gt; <br/>&lt;/form&gt; <br/>
					<br/><br/>
				</div>
				<div class="Content-body">
					
					<img src="images/From.gif" style="margin:0px 2px -4px 0px" alt=""/><strong>文章来自:</strong> <a href="http://jinbenli.com/" target="_blank">本站原创</a><br/>
					<img src="images/icon_trackback.gif" style="margin:4px 2px -4px 0px" alt=""/><strong>引用通告:</strong> <a href="trackback.asp?tbID=26&amp;action=view" target="_blank">查看所有引用</a> | <a href="javascript:;" title="获得引用文章的链接" onclick="getTrackbackURL(26)">我要引用此文章</a><br/>
					<img src="images/tag.gif" style="margin:4px 2px -4px 0px" alt=""/><strong>Tags:</strong> <a href="default.asp?tag=asp%E4%BB%A3%E7%A0%81">asp代码</a><a href="http://technorati.com/tag/asp代码" rel="tag" style="display:none">asp代码</a> <br/>
					<img src="images/notify.gif" style="margin:4px 2px -4px 0px" alt=""/><strong>相关日志:</strong>
                    <div class="Content-body" id="related_tag" style="margin-left:25px"></div>
                    <script language="javascript" type="text/javascript">Related(26, 1, "related_tag");</script>
				</div>
				<div class="Content-bottom"><div class="ContentBLeft"></div><div class="ContentBRight"></div>评论: 0 | <a href="trackback.asp?tbID=26&amp;action=view" target="_blank">引用: 0</a> | 查看次数: <$log_ViewNums$></div>
			</div>
		</div>
