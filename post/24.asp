        <%ST(A)%>
			<div id="Content_ContentList" class="content-width"><a name="body" accesskey="B" href="#body"></a>
				<div class="pageContent">
					<div style="float:right;width:auto"><img border="0" src="images/Cprevious.gif" alt=""/><strong>上一篇:</strong> <a href="?id=23" accesskey=",">48个诡异心理学</a><br><img border="0" src="images/Cnext.gif" alt=""/><strong>下一篇:</strong> <a href="?id=25" accesskey=".">最简单的ASP生成静态文件代码</a><br></div> 
					<img src="images/icons/3.gif" style="margin:0px 2px -4px 0px" alt=""/> <strong><a href="default.asp?cateID=4" title="查看所有【资源共享】的日志">资源共享</a></strong> <a href="feed.asp?cateID=4" target="_blank" title="订阅所有【资源共享】的日志" accesskey="O"><img border="0" src="images/rss.png" alt="订阅所有【资源共享】的日志" style="margin-bottom:-1px"/></a>
				</div>
				<div class="Content">
					<div class="Content-top"><div class="ContentLeft"></div><div class="ContentRight"></div>
					<h1 class="ContentTitle"><strong>怎么实现让电脑在停电断电后，来电时能自动开机</strong></h1>
					<h2 class="ContentAuthor">作者:jinbenli 日期:2011-07-26 16:17:19</h2>
				</div>
			    <div class="Content-Info">
					<div class="InfoOther">字体大小: <a href="javascript:SetFont('12px')" accesskey="1">小</a> <a href="javascript:SetFont('14px')" accesskey="2">中</a> <a href="javascript:SetFont('16px')" accesskey="3">大</a></div>
					<div class="InfoAuthor"><img src="images/weather/hn2_sunny.gif" style="margin:0px 2px -6px 0px" alt=""/><img src="images/weather/hn2_t_sunny.gif" alt=""/> <img src="images/level3.gif" style="margin:0px 2px -1px 0px" alt=""/><$EditAndDel$></div>
				</div>
				<div id="logPanel" class="Content-body">
					首先进入CMOS的设置主界面，选择【POWER MANAGEMENT SETUP】，再选择【PWR Lost Resume State】，这一项有三个选择项。 如果选择其中的【Keep OFF】项，代表停电后再来电时，电脑不会自动启动。<br/> 如果选择【Turn On】，代表停电后再来电时，电脑会自动启动。如图所示。 如果是选择的【Last State】，那么代表停电后再来电时，电脑回复到停电前电脑的状态。断电前如果电脑是处于开机状态，那么来电后就会自动开机。断电前是处于关机状态，那么来电后电脑不会自动开机。而我的是梅捷主板，进入CMOS的设置主界面后，选择【Integrated Peripherals】(周边设备设置)，再选择【SuperIO Function】（其它集成驱动选项），然后进入【PWRON After PWR-Fail】（电源回复后的选择） 如果选择【OFF】，代表需按机箱面板上的电源开关才能开机。如果选择【Former-Sts】，代表电源回复时恢复系统断电前的状态。如果选择【ON】，代表电源回复时直接开机。
					<br/><br/>
				</div>
				<div class="Content-body">
					
					<img src="images/From.gif" style="margin:0px 2px -4px 0px" alt=""/><strong>文章来自:</strong> <a href="http://jinbenli.com/" target="_blank">本站原创</a><br/>
					<img src="images/icon_trackback.gif" style="margin:4px 2px -4px 0px" alt=""/><strong>引用通告:</strong> <a href="trackback.asp?tbID=24&amp;action=view" target="_blank">查看所有引用</a> | <a href="javascript:;" title="获得引用文章的链接" onclick="getTrackbackURL(24)">我要引用此文章</a><br/>
					<img src="images/tag.gif" style="margin:4px 2px -4px 0px" alt=""/><strong>Tags:</strong> <a href="default.asp?tag=%E7%94%B5%E8%84%91%E6%8A%80%E5%B7%A7">电脑技巧</a><a href="http://technorati.com/tag/电脑技巧" rel="tag" style="display:none">电脑技巧</a> <br/>
					<img src="images/notify.gif" style="margin:4px 2px -4px 0px" alt=""/><strong>相关日志:</strong>
                    <div class="Content-body" id="related_tag" style="margin-left:25px"></div>
                    <script language="javascript" type="text/javascript">Related(24, 1, "related_tag");</script>
				</div>
				<div class="Content-bottom"><div class="ContentBLeft"></div><div class="ContentBRight"></div>评论: 0 | <a href="trackback.asp?tbID=24&amp;action=view" target="_blank">引用: 0</a> | 查看次数: <$log_ViewNums$></div>
			</div>
		</div>
