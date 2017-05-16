<!--#include file="../../commond.asp" -->
<!--#include file="../../header.asp" -->
<!--#include file="../../plugins.asp" -->
<!--#include file="../../common/ModSet.asp" -->
<!--#include file="../../common/UBBconfig.asp" -->
<!--内容-->
<div id="highslide-container"></div>
<div id="Tbody">
<div id="mainContent">
<div id="innermainContent">
  <div id="mainContent-topimg"></div>
  <div id="Content_ContentList" class="content-width"> <%=content_html_Top%>
    <%
dim albumSet,Getplugins,albumDBPath
Set albumSet=New ModSet
Getplugins=CheckStr(Request.QueryString("plugins"))
albumSet.open(Getplugins)
albumDBPath=albumSet.getKeyValue("Database")
albumDBPath="Plugins\album\"&albumDBPath
	'打开数据库
Dim albumConn
Set albumConn=Server.CreateObject("ADODB.Connection")
albumConn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath(albumDBPath)&""
albumConn.Open
'变量定义区
Dim albumDBb,haa
%>
    <!--相册样式代码开始-->
    <style> 
.xctitle{background-color:#f5f5f5;margin:12px; font-size:12px; border:#CCC dotted 1px; padding:3px;word-break :keep-all ;text-justify :inter-word ;line-height:23px; text-align:center }
.xctitle img{vertical-align:middle; margin-right:10px}
.xctitle a{ font-size:12px; color:#006699; margin:0 10px}
.xcshowtitle{ background-color:#f5f5f5;margin:12px; font-size:12px; border:#CCC dotted 1px; padding:5px;}
.xcshowtitle a{ font-size:12px; color:#006699; margin:3px 5px}
.xcshowtitle img{vertical-align:middle; margin-right:10px}
.xcshowtxt{ font-size:14px; }
.xc{margin:10px; text-align:center }
.xclist {margin:10px 20px; border:#CCC dotted 1px; padding:10px;line-height:23px; text-align:left;text-justify :inter-word ;}
.xclist a{ margin:10px;}
.xcshow{margin:10px 20px 10px 20px;}
.xckuang{margin:4px; width:200px; height:200px;float:left; background:url(Plugins/album/images/backimg.gif) no-repeat; padding-top:8px; text-align:center }
.xcclass{  }
.xcclasskuang{margin:10px 13px; width: auto; float:left;  }
.xcclasskuang input{width: 100px; border:1px #CCCCCC solid; margin:0px; padding:0px}
.bar{clear:both}
.hui{ color:#999999}
.huitime{ color:#999999;}
.txt a { color:#FF6600}
.txt img{vertical-align:middle; margin:0px 15px; border:none}
.xctxt{margin:10px 0px 10px 0px; line-height:23px; text-align:left }
.xctxt img{vertical-align:middle; }
.divdiv{ text-align:center; width:auto;}
.divdiv img{vertical-align:middle; margin-right:10px; border:none}
.divright{ text-align:right; margin:0px 20px 0px 180px}
.divleft{text-align:right; margin:0px 10px 0px 0px}
img{ border:none}
.cleardiv{ clear:both}
</style>
    <!--相册样式代码结束-->
    <!--详细页面开始-->
    <% IF Request.QueryString("action")="Comment" Then
'====检测相册是否关闭====
if not cBool(albumSet.getKeyValue("start")) then
showmsg "错误信息","相册暂时关闭！<br/><a href=""default.asp"">单击返回首页</a>","WarningIcon",""
end if
''
dim xcid , xcclass , qx 

xcclass = CheckStr(Request.QueryString("class"))
if  stat_Admin  then qx=1 else qx=0

Dim albumdbclass,classname
set albumdbclass=Server.CreateObject("Adodb.Recordset")
sql="select * from blog_album_class where album_classID ="&xcclass
albumdbclass.Open SQL,albumConn,1,1
classname=albumdbclass("album_className")
albumdbclass.close
set albumdbclass=nothing

%>
<script type="text/javascript" src="Plugins/album/js/bingo.js"></script>
    <style type="text/css">
#divall {
		position:relative;
	text-align:center
}
#divleft {
		border:0px red solid;
	position:absolute;
	z-index:2007;
	text-align:right;
	padding-top:20px;
	padding-left:20px;
	
}
#divright {
		border:0px red solid;
	position:absolute;
	z-index:2007;
	text-align:center;
	padding-top:20px;
	padding-left:20px;
	
}
#tu{
	z-index:-2007;
	
}


</style>
    <div class="Content" >
      <div class="xc">
        <div class="xcshowtitle"> <img src="Plugins/album/images/photo.gif" ><a href="LoadMod.asp?plugins=album&class=<%=xcclass%>">&nbsp;<%=classname%>&nbsp;&nbsp;( <span id="ptj"></span>/<span id="pnum"></span>&nbsp;P )</a><span id="prep"></span>&nbsp;/&nbsp;<span id="nextp"></span>&nbsp;&nbsp;|&nbsp;<a href="LoadMod.asp?plugins=album&class=<%=xcclass%>">返回照片列表</a>&nbsp;|&nbsp;<a href="LoadMod.asp?plugins=album">返回相册分类</a> </div>
        <div class="xcshow">
          <div class="divdiv">
            <div class="xcshowtxt" id="ptitle"></div>
            <div style="text-align:right; float:right; margin-bottom:3px;"><span class="txt" id="yuantu"></span></div>
            <div class="cleardiv"></div>
            <div id="divall" ></div>
            <div id="a1"><br/>
              <br/>
              <br/>
              <img src="Plugins/album/images/load.gif" /><br/>
              <br/>
              <br/>
              <br/>
            </div>
            <div class="cleardiv"></div>
            <div style="text-align:left; float:left; margin-top:5px "> <span class="huitime">浏览<span id="pcount"></span>次</span> <br/>
            </div>
            <div style="text-align:right; float:right;margin-top:5px"> <span class="huitime">上传于：<span id="ptime"></span></span> <br />
              <span class="txt"><a onClick="javascript:ShowFLT(1)" href="javascript:void(null)"><img src="Plugins/album/images/exif.gif" >EXIF信息</a></span> </div>
            <div class="cleardiv"></div>
            <div class="xctxt" ><img src="Plugins/album/images/info.gif" ><span id="pinfo"></span></div>
            <% if memName<>empty and stat_Admin then%>
            <div id="xiugai"></div>
            <% end if %>
            <DIV class=txtbox id=LM1 style="DISPLAY: none">
              <div id="exif">
                <Table border="0" cellspacing="0" cellpadding="0" style="width:90%;border:#CCC dotted 1px">
                  <tr height="18px">
                    <td align="right">相机品牌：</td>
                    <td align="left"><span id="CameraMake"></span></td>
                    <td align="right">相机型号：</td>
                    <td align="left"><span id="CameraModel"></span></td>
                  </tr>
                  <tr height="18">
                    <td align="right">拍摄时间：</td>
                    <td align="left"><span id="DateTime"></span></td>
                    <td align="right">照片尺寸：</td>
                    <td align="left"><span id="ImageDimension"></span></td>
                  </tr>
                  <tr height="18">
                    <td align="right">编辑工具：</td>
                    <td align="left"><span id="Software"></span></td>
                    <td align="right">ISO 速度：</td>
                    <td align="left"><span id="ISOSpeed"></span></td>
                  </tr>
                  <tr height="18">
                    <td align="right">光&nbsp;&nbsp;圈：</td>
                    <td align="left"><span id="FStop"></span></td>
                    <td align="right">曝光时间：</td>
                    <td align="left"><span id="ExposureTime"></span></td>
                  </tr>
                  <tr height="18">
                    <td align="right">闪光灯：</td>
                    <td align="left"><span id="Flash"></span></td>
                    <td align="right">曝光补偿：</td>
                    <td align="left"><span id="ExposureBias"></span></td>
                  </tr>
                  <tr height="18">
                    <td align="right">焦&nbsp;&nbsp;距：</td>
                    <td align="left"><span id="FocalLength"></span></td>
                    <td align="right">测距模式：</td>
                    <td align="left"><span id="MeteringMode"></span></td>
                  </tr>
                </table>
              </div>
              <div class="Content-bottom" style="height:1px"></div>
            </DIV>
          </div>
        </div>
        <div id="test2"></div>
        <div id="MsgContent" style="width:90%; DISPLAY: none">
          <div id="MsgHead">发表评论</div>
          <div id="MsgBody">
            <form name="frm" action="Plugins/album/albumaction.asp" method="post" style="margin:0px;">
              <table width="100%" cellpadding="0" cellspacing="0">
                <tr>
                  <td align="center" width="15%"><strong>昵　称:</strong></td>
                  <td align="left"><input name="username" type="text" size="18" class="userpass" maxlength="24" <%if not memName=empty then response.write ("value="""&memName&""" readonly=""readonly""")%>/>
                    &nbsp;&nbsp;&nbsp;<font color="#FF0000">（必填）</font> </td>
                </tr>
                <%if memName=empty then%>
                <tr>
                  <td align="center" width="15%"><strong>密　码:</strong></td>
                  <td align="left"><input name="password" type="password" size="18" class="userpass" maxlength="24"/>
                    游客发言不需要密码.</td>
                </tr>
                <%end if%>
                <%if memName=empty or blog_validate=true then%>
                <tr>
                  <td align="center" width="15%"><strong>验证码:</strong></td>
                  <td align="left"><input name="validate" type="text" size="4" class="userpass" maxlength="4"/>
                    <img id="vcodeImg" src="about:blank" onerror="this.onerror=null;this.src='common/getcode.asp?s='+Math.random();" alt="验证码" title="看不清楚？点击刷新验证码！" style="margin-right:40px;cursor:pointer;width:40px;height:18px;margin-bottom:-4px;margin-top:3px;" onclick="src='common/getcode.asp?s='+Math.random()"/></td>
                </tr>
                <%end if%>
                <tr>
                  <td align="center" width="15%" valign="top"><strong>内　容:</strong></td>
                  <td align="center" width="80%">
                    <script type="text/javascript" src="common/UBBCode.js"></script>
		            <script type="text/javascript" src="common/UBBCode_help.js"></script>
		            <div id="UBBSmiliesPanel" class="UBBSmiliesPanel"></div>
		            <textarea id="editMask" class="field" style="width:99%;height:100px" onfocus="loadUBB('Message')"></textarea>
		            <div id="editorbody" style="display:none">
		              <div id="editorHead">正在加载编辑器...</div>
		              <div class="editorContent">
		                <textarea name="Message" class="editTextarea" style="height:150px;;" cols="1" rows="1" accesskey="R"></textarea>
		              </div>
		            </div></td>
                </tr>
                <tr>
                  <td colspan="2" align="center"><input name="action" type="hidden" value="post"/>
                    <input name="submit2" type="submit" class="userbutton" value="发表评论" accesskey="S"/>
　
                    <input name="button" type="reset" class="userbutton" value="重写"/></td>
                </tr>
                <tr>
                  <td colspan="2" align="right" ><%if memName=empty then%>
                    虽然发表评论不用注册，但是为了保护您的发言权，建议您<a href="register.asp">注册帐号</a>. <br/>
                    <%end if%>
                    字数限制 <b><%=blog_commLength%> 字</b> | 
                    UBB代码 <b>
                    <%if (blog_commUBB=0) then response.write ("开启") else response.write ("关闭") %>
                    </b> | 
                    [img]标签 <b>
                    <%if (blog_commIMG=0) then response.write ("开启") else response.write ("关闭") %>
                    </b> </td>
                </tr>
              </table>
              <div id="hiddenid">
                <input name="albumID" type="hidden" value="11"/>
              </div>
            </form>
          </div>
        </div>
      </div>
    </div>
<script type="text/javascript">
var qx = <%=qx%>;
var xcclass = <%=xcclass%>;
window.onload = function() {
    var iurl = parseInt(location.hash.replace("#", ""));
    var xcid = iurl;
    getweblist(xcclass, xcid)
}
function getweblist(xcclass, xcid) {
    window.location.hash = xcid;
    AJAXCALL("plugins/album/server.asp?xcclass=" + xcclass + "&xcid=" + xcid, "pxx", {
        "isCache": false,
        "statePoll": showState
    })
}
function showState(n) {
    if (parseInt(n) != 4) {
        $('divall').innerHTML = "<br/><br/><br/><img src=Plugins/album/images/load.gif /><br/><br/><br/><br/>";
    }
}


function pxx(doc) {
    var rd = unescape(doc);
    eval("var Img = " + rd);
    $('ptitle').innerHTML = Img.ptitle;
    $('pinfo').innerHTML = Img.pinfo;
    $('ptime').innerHTML = Img.ptime;
    $('pcount').innerHTML = Img.pcount;
    $('pnum').innerHTML = Img.pnum;
    $('ptj').innerHTML = Img.ptj;
    if (Img.pcomm != '') {
        if (qx == 0) {
            $('test2').innerHTML = unescape(Img.pcomm)
        } else {
            $('test2').innerHTML = unescape(Img.dcomm)
        }
    } else {
        $('test2').innerHTML = ''
    }
    if (Img.pcom == 1) {
        $('MsgContent').style.display = ""
    } else {
        $('MsgContent').style.display = "none"
    }
    var hiddenid = "<input name='albumID' type='hidden' value='" + Img.pid + "'/>";
    $('hiddenid').innerHTML = hiddenid;
    if (Img.prep != 0) {
        var prep = "<a href=javascript:void(getweblist(" + Img.pclass + "," + Img.prep + ")) title='可点击图片的左侧查看' >上一张</a>"
    } else {
        var prep = "这是第一张"
    }
    $('prep').innerHTML = prep;
    if (Img.nextp != 0) {
        var nextp = "<a href=javascript:void(getweblist(" + Img.pclass + "," + Img.nextp + ")) title='可点击图片的右侧翻页查看' >下一张</a>"
    } else {
        var nextp = "这是最后一张"
    }
    $('nextp').innerHTML = nextp;
    var cc = "<img src=" + Img.purl + " id=tu onLoad=\"javascript:if(this.width>260)this.width=335;this.style.display=\'block\';$(\'a1\').style.display=\'none\'\" onMouseMove=show_who(this,event," + Img.prep + "," + Img.nextp + ") onClick=click_who(" + Img.prep + "," + Img.nextp + ") >";
    var ccyc = "<div style=padding:80px;background-color:#CCC >隐藏图片</div>";
    if (Img.phidden == 1 || qx == 1) {
        $('divall').innerHTML = cc
    } else {
        $('divall').innerHTML = ccyc
    }
    if (Img.phidden == 1 || qx == 1) {
        var yt = "<a href=" + Img.purl + " title=" + Img.ptitle + " target=_blank ><img src=Plugins/album/images/f.gif />查看原图</a>"
    } else {
        var yt = "<img src=Plugins/album/images/f.gif />隐藏图片"
    }
    $('yuantu').innerHTML = yt;
    var noexif = "<div class=xcshow><span class=huitime>该照片没有EXIF信息</span></div>";
    if (Img.pexif == 0) {
        $('exif').innerHTML = noexif
    } else {
        $('CameraMake').innerHTML = Img.CameraMake;
        $('CameraModel').innerHTML = Img.CameraModel;
        $('DateTime').innerHTML = Img.DateTime;
        $('ImageDimension').innerHTML = Img.ImageDimension;
        $('Software').innerHTML = Img.Software;
        $('ISOSpeed').innerHTML = Img.ISOSpeed;
        $('FStop').innerHTML = Img.FStop;
        $('ExposureTime').innerHTML = Img.ExposureTime;
        $('Flash').innerHTML = Img.Flash;
        $('ExposureBias').innerHTML = Img.ExposureBias;
        $('FocalLength').innerHTML = Img.FocalLength;
        $('MeteringMode').innerHTML = Img.MeteringMode
    }
}

var evt_left, evt_top;
function show_who(o, evt, prex, nextx) {
    evt = evt?evt:window.event;
    initX = evt.offsetX?evt.offsetX:evt.layerX;
    evt_left = o.clientWidth / 2;
    evt_top = o.clientHeight / 2;
    if (initX <= o.clientWidth / 2) {
        if (prex != 0) {
            o.style.cursor = "url(Plugins/<%=albumSet.GetPath%>/images/pre.cur)";
            o.alt = "查看上一张"
        } else {
            o.style.cursor = "auto";
            o.alt = "这是第一张"
        }
    } else {
        if (nextx != 0) {
            o.style.cursor = "url(Plugins/<%=albumSet.GetPath%>/images/next.cur)";
            o.alt = "查看下一张"
        } else {
            o.style.cursor = "auto";
            o.alt = "这是最后一张"
        }
    }
    return
}

function click_who(x, y) {
    if (event.clientX <= evt_left) {
        pre(x)
    } else {
        next(y)
    }
}
function next(y) {
    if (y != 0) {
        javascript: void(getweblist(xcclass, y))
    }
}
function pre(x) {
    if (x != 0) {
        javascript: void(getweblist(xcclass, x))
    }
}
function LMYC() {
    LM1.style.display = 'none'
}
function ShowFLT() {
    if (LM1.style.display == 'none') {
        LMYC();
        LM1.style.display = ''
    } else {
        LM1.style.display = 'none'
    }
}
</script>
    <!--详细页面结束-->
    <!--列表页面开始-->
    <% ELSEIF Request.QueryString("class")<>"" or Request.QueryString("action")="newphoto" Then%>
    <div class="Content" >
      <%
'====检测相册是否关闭====
If Not cBool(albumSet.getKeyValue("start")) Then showmsg "错误信息","相册暂时关闭！<br/><a href=""default.asp"">单击返回首页</a>","WarningIcon",""
dim action, albumID, albumedit, albumDB, albumNum, PageCount, fclass
action = Request.QueryString("action")
albumID = CheckStr(Request.QueryString("id"))
set albumDB = Server.CreateObject("Adodb.Recordset")
If action = "newphoto" then
    SQL = "select top 12 * from blog_album order by album_ID desc"
Else
    fclass = CheckStr(Request.QueryString("class"))
    SQL = "select * from blog_album where album_class="&fclass&" order by album_ID desc"
    Url_Add = Url_Add&"plugins=album&class="&fclass&"&"
End If
albumDB.Open SQL,albumConn,1,1
If albumDB.eof And albumDB.bof Then
    response.write "<div style=""margin:10px 0px 10px 0px""><strong>抱歉，没有找到任何图片！</strong></div>"
Else 
    Dim albumPage, albumBank
    albumPage = Int(albumSet.getKeyValue("Page"))
    albumBank = Int(albumSet.getKeyValue("newh"))
    albumDB.PageSize = albumPage
    albumDB.AbsolutePage = CurPage
    albumNum = albumDB.RecordCount                
%>
      <div class="xctitle"><img src="Plugins/album/images/photo.gif" >
        <%
Dim dclass
IF Request.QueryString("class") <> "" then
    Set dclass = albumConn.ExeCute("select * from blog_album_class where album_classID="&CheckStr(Request.QueryString("class")))
        If dclass.eof Or dclass.bof Then showmsg "错误信息","<b>您选择的类型不存在!</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>","WarningIcon","plugins"
        Response.Write dclass("album_className")
End If
%>
        &nbsp;共&nbsp;<%=albumNum%>&nbsp;张照片&nbsp;&nbsp;|<a href="LoadMod.asp?plugins=album&action=newphoto">最新图片</a>|<a href="LoadMod.asp?plugins=album&action=newComment">照片评论</a>|<a href="LoadMod.asp?plugins=album">返回相册分类</a> </div>
      <div class="Content-bottom" style="height:1px"></div>
      <%if Not albumedit then%>
      <div class="pageContent"><%=MultiPage(albumNum,albumPage,CurPage,Url_Add,"","float:right")%></div>
      <%end if%>
      <div class="xc">
        <div class="xcclass">
          <%
Dim II,J
II = albumPage
J = 0
Do While Not albumDB.Eof And II>0
II = II-1
J = J + 1
%>
          <div class="xcclasskuang">
            <%if (albumDB("album_Hidden") and Lcase(albumDB("album_user"))=Lcase(memName)) Or stat_Admin or (not albumDB("album_Hidden")) then
fclass = albumDB("album_class")
dim imgurl
imgurl = albumDB("album_Urlm") 
If FileExist(imgurl) Then
    imgurl = albumDB("album_Urlm") 
Else
    imgurl="Plugins/album/images/noaspjpeg.jpg"
End If
%>
            <object type="application/x-shockwave-flash" data="Plugins/album/swfoto2.swf?image=<%=imgurl%>&link=LoadMod.asp?plugins=album^action=Comment^class=<%=fclass%>#<%=albumDB("album_ID")%>&roundCorner=3&windowOpen=_self&isShowLoader=1&loaderColor=0x999999" width="130" height="130" id="swfoto2">
              <param name="movie" value="Plugins/album/swfoto2.swf?image=<%=imgurl%>&link=LoadMod.asp?plugins=album^action=Comment^class=<%=fclass%>#<%=albumDB("album_ID")%>&roundCorner=3&windowOpen=_self&isShowLoader=1&loaderColor=0x999999" />
            </object>
            <br/>
            <br/>
            <a href="LoadMod.asp?plugins=album&action=Comment&class=<%=fclass%>#<%=albumDB("album_ID")%>"><%=albumDB("album_Title")%></a>
            <%else%>
            <div style=" height:80px; width:130px; padding-top:50px; background-color:#ccc ">隐藏图片</div>
            <%end if%>
          </div>
          <%
If (J mod albumBank)=0 Then 
Response.Write "<div class=bar></div>" 
End If
albumDB.MoveNext
Loop
response.write "</div></div>"
if Not albumedit then response.write "<div class=bar></div><div class=""pageContent"">"&MultiPage(albumNum,albumPage,CurPage,Url_Add,"","float:right")&"</div>" end if
end if
albumDB.close
set albumDB=nothing
%>
          <!--相册列表-->
          <div class="Content-bottom" style="height:1px"></div>
          <div class="xclist" >
            <%
Dim albumclassDB,album_class
set albumclassDB=Server.CreateObject("Adodb.Recordset")
SQLQueryNums=SQLQueryNums+1
Set album_class=albumConn.execute("select * from blog_album_class order by album_classID desc")
do until album_class.eof 
%>
            <a href="LoadMod.asp?plugins=album&class=<%=album_class("album_classID")%>"><%=album_class("album_className")%></a>
            <% 
album_class.movenext
loop
%>
          </div>
          <!--相册列表结束-->
        </div>
        <!--列表页面结束-->
        <!--最新评论开始-->
        <% ElseIF Request.QueryString("action")="newComment" Then%>
        <div class="Content" >
          <div class="xctitle"><img src="Plugins/album/images/photo.gif" > <a href="LoadMod.asp?plugins=album&action=newphoto">最新图片</a>|<a href="LoadMod.asp?plugins=album">返回相册分类</a></div>
          <div class="xc">
            <%
'====检测相册是否关闭====
if Not cBool(albumSet.getKeyValue("start")) Then showmsg "错误信息","相册暂时关闭！<br/><a href=""default.asp"">单击返回首页</a>","WarningIcon",""
Dim Commentnum
Commentnum = int(albumSet.getKeyValue("Commentnum"))
dim album_Commentsnew, pcomclass
Set album_Commentsnew = albumConn.execute("select top "&Commentnum&" * from blog_album_Comment order by album_CommentID desc")
do until album_Commentsnew.eof
pcomclass = albumConn.execute("select album_class from blog_album where album_ID="&album_Commentsnew("album_ID"))(0)
%>
            <br>
            <div class="comment" >
              <div class="commenttop"><img border="0" src="images/<%if memName=album_Commentsnew("album_CommentUSER") then response.write ("icon_quote_author.gif") else response.write ("icon_quote.gif") end if%>" alt="" style="margin:0px 4px -3px 0px"/><a href="member.asp?action=view&memName=<%=Server.URLEncode(album_Commentsnew("album_CommentUSER"))%>"><strong><%=album_Commentsnew("album_CommentUSER")%></strong></a> <span class="commentinfo">[<%=DateToStr(album_Commentsnew("album_CommentTIME"),"Y-m-d H:I A")%>
                <%if stat_Admin then response.write (" | "&album_Commentsnew("album_CommentIP")) end if%>
                <%if stat_Admin=true then response.write (" | <a href=""plugins/album/albumaction.asp?action=delcomms&amp;commID="&album_Commentsnew("album_CommentID")&""" onclick=""if (!window.confirm('是否删除该评论?')) {return false}""><img src=""images/del1.gif"" alt=""删除该评论"" border=""0""/></a>") end if%>
                ]</span></div>
              <div class="commentcontent" id="commcontent_<%=album_Commentsnew("album_CommentID")%>"> <%=UBBCode(HtmlEncode(album_Commentsnew("album_CommentMessager")),0,0,0,1,1)%></a><br>
                <br>
                <a href="LoadMod.asp?plugins=album&action=Comment&class=<%=pcomclass%>#<%=album_Commentsnew("album_ID")%>">查看该图片</a> </div>
            </div>
            <div class="bar"></div>
            <%
album_Commentsnew.movenext
loop
album_Commentsnew.close
Set album_Commentsnew=nothing
%>
          </div>
        </div>
        <!--最新评论结束-->
        <%else%>
        <!--相册整体框架开始-->
        <div class="Content" >
          <%
'====检测相册是否关闭====
If Not cBool(albumSet.getKeyValue("start")) Then showmsg "错误信息","相册暂时关闭！<br/><a href=""default.asp"">单击返回首页</a>","WarningIcon",""
Dim classtj, cclass, tjnum
classtj = Int(albumConn.Execute("select count(*) from blog_album_class")(0))
Set cclass = albumConn.execute("select * from blog_album_class order by album_classID desc")
%>
          <div class="xctitle"><img src="Plugins/album/images/photo.gif" > 共&nbsp;<%=classtj%>&nbsp;个照片分类&nbsp;&nbsp;&nbsp;|<a href="LoadMod.asp?plugins=album&action=newphoto">最新图片</a>|<a href="LoadMod.asp?plugins=album&action=newComment">照片评论</a> </div>
          <div  class="xc">
            <%
haa = 0
Response.Write "<div class=bar></div>"
Do While Not cclass.Eof
haa = haa+1
tjnum= Int(albumConn.Execute("select count(*) from blog_album where album_class="&cclass("album_classID"))(0))
%>
            <div class="xckuang">
              <object type="application/x-shockwave-flash" data="Plugins/album/swfoto2.swf?image=<%=cclass("album_classPic")%>&link=LoadMod.asp?plugins=album^class=<%=cclass("album_classID")%>&roundCorner=3&windowOpen=_self&isShowLoader=1&loaderColor=0x999999" width="130" height="130" id="swfoto">
                <param name="movie" value="Plugins/album/swfoto2.swf?image=<%=cclass("album_classPic")%>&link=LoadMod.asp?plugins=album^class=<%=cclass("album_classID")%>&roundCorner=3&windowOpen=_self&isShowLoader=1&loaderColor=0x999999" />
              </object>
              <br/>
              <br/>
              <a href="LoadMod.asp?plugins=album&class=<%=cclass("album_classID")%>"><font color="#006699" style="font-size:14px"><%=cclass("album_className")%></font></a><br />
              <div class="hui"><%=tjnum%>p</div>
            </div>
            <%
If (haa mod Int(albumSet.getKeyValue("Bank")))=0 Then 
Response.Write "<div class=bar></div>" 
End If
cclass.MoveNext
Loop
response.write "</div>"
%>
          </div>
          <%end if%>
          <!--相册整体框架结束-->
        </div>
        <%=content_html_Bottom%>
        <div id="mainContent-bottomimg"></div>
      </div>
    </div>
    <div id="sidebar">
      <div id="innersidebar">
        <div id="sidebar-topimg">
          <!--工具条顶部图象-->
        </div>
        <%=side_html%>
        <div id="sidebar-bottomimg"></div>
      </div>
    </div>
  </div>
  <div style="font: 0px/0px sans-serif;clear: both;display: block"></div>
  <!--#include file="../../footer.asp" -->
