<!--#include file="conn.asp" -->
<%

dim at
IF session("upadmin") ="guanli" Then
at=1
else 
at=0
end if

response.CharSet = "UTF-8"
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.cachecontrol = "no-cache"
Server.ScriptTimeout = 9999
function HTMLEncode1(fString)
if fString<>"" then
    fString = Replace(fString, "", CHR(13))
	fString = Replace(fString, "&quot;",CHR(34) )
	fString = Replace(fString, "<br>" , CHR(10))
	fString = Replace(fString, "&nbsp;",CHR(32) ) 
    fString = Replace(fString, "&nbsp;",CHR(9) )
end if
	HTMLEncode1= fString
end function
function HTMLEncode2(fString)
if fString<>"" then
	fString = Replace(fString, ">", "&gt;")
	fString = Replace(fString, "<", "&lt;")
	fString = Replace(fString, CHR(34), "&quot;")
	fString = Replace(fString, CHR(13), "")
	fString = Replace(fString, CHR(10), "<br>")
	fString = Replace(fString, CHR(32), "&nbsp;")
end if
	HTMLEncode2 = fString
end function
Select case Request.QueryString("act")
case "gtname"
	Call endphoto()
case "sphoto"
	Call Showphoto()
case "ephoto"
    Call ediphoto()
case "putsb"
	Call putphoto()
case "comsb"
	Call comphoto()
case "edohuan"
	Call EditOthe()
case "epicoh"
	Call edophoto()
case else
		Call Errorc()
end Select
Sub Errorc()
	response.write("你还没有登陆或超时")
End Sub

Sub endphoto()

if at=0 then
		Errorc()
		exit Sub
	end if

	album_class=clng(Request.QueryString("id"))
	Set rs = Server.CreateObject("ADODB.Recordset")
	sql="select album_className from blog_album_class where album_classID="&album_class&" order by album_classID desc"
	rs.open sql,conn,1,1 
		if rs.eof and rs.bof then
			Response.Write""
		else
			Response.Write rs("album_className")
		end if
	rs.Close
	set rs=nothing
End Sub

Sub Showphoto()
	
	if at=0 then
		Errorc()
		exit Sub
	end if
	dim typeid,rs,dcount,pagest,jj,hots,title,coms,titlec
typeid=clng(Request.QueryString("typeid"))
	if IsNumeric(Request.QueryString("typeid"))=False then
		Response.Write("[{pagehtm:""<div class='picxno'>错误的分类id<div></div></div>"",pages:0,dcount:0,page:0}]")
		
	end if
	
	Dim ipagecount
	Dim ipagecurrent
	Dim irecordsshown  
	if request("page")="" then
		ipagecurrent=1
	else
		ipagecurrent=cint(request("page"))
	end if
	sql = "select * from blog_album where album_class="&typeid&" order by album_ID desc" 
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.pagesize = 20
	rs.cachesize = 20
	rs.open sql,conn,1,1
	dcount=rs.recordcount
	ipagecount = rs.pagecount
	if ipagecurrent > ipagecount Then ipagecurrent = ipagecount
	if ipagecurrent < 1 Then ipagecurrent = 1
	if ipagecount=0 then
		Response.Write("[{pagehtm:""<div class='picxno'>没 有 任 何 图 片，请 先 添 加 图 片<div></div></div>"",pages:0,dcount:0,page:0}]")
	else
		pagest=""
		rs.absolutepage = ipagecurrent
		jj = 0
		do while jj<20 and NOT rs.EOF
		jj = jj +1
			if  rs("album_Hidden") then
				hots="class='red'"
				title="点击不隐藏本图片"
			else
				hots=""
				title="点击隐藏本图片"
			end if
			if  rs("album_Comment") then
				coms="class='red'"
				titlec="点击取消禁止评论"
			else
				coms=""
				titlec="点击禁止评论"
			end if
			pagest=pagest&"<div class='picx'>"
			pagest=pagest&"<div class='atpic'><img src='/"&rs("album_Urlm")&"' /></div>"
			pagest=pagest&"<div class='edpictxt'>"
			pagest=pagest&"<div class='edtsh' id='sh"&jj&"'><div class='edpictxt1'onclick='showpedt("&jj&")'><a href='javascript:void(0)' title='点击修改名称'><input type='text' id='edzen"&jj&"' class='editinput2'value='"&rs("album_Title")&"' onmouseover=this.style.backgroundColor='#BFDFFF' onmouseout=this.style.backgroundColor='#FFFFFF'></a></div>"
			pagest=pagest&"<div class='edpictxtx'>"
			pagest=pagest&"<a href='javascript:void(0)' onclick='editphoto("&rs("album_ID")&","&rs("album_class")&")'><span>修改</span></a>|"
			pagest=pagest&"<a href='javascript:void(0)' onclick=Sptpic("&jj&","&rs("album_ID")&") title="&title&" onfocus='this.blur()'><span id='tj"&jj&"' "&hots&">隐藏</span></a>|"
			pagest=pagest&"<a href='javascript:void(0)' onclick=comtpic("&jj&","&rs("album_ID")&") title="&titlec&" onfocus='this.blur()'><span id='com"&jj&"' "&coms&">评论</span></a>|"
			pagest=pagest&"<a href='javascript:void(0)' onclick='DePhoto("&rs("album_ID")&")'><span class='red'>删除</span></a></div></div>"
			pagest=pagest&"<div class='edtxh' id='eh"&jj&"'><div class='edpictxt1'><input type='text'id='edzin"&jj&"' class='editinput1'value='"&rs("album_Title")&"'></div>"
			pagest=pagest&"<div class='edpictxtx'>[<a href='javascript:void(0)' onclick='editpic("&jj&","&rs("album_ID")&")'>确定</a>|<a href='javascript:void(0)' onclick='showmout("&jj&")'>取消</a>]</div>"
			pagest=pagest&"</div></div></div>"
			rs.movenext			
		loop
		Response.Write("[{pagehtm:"""&pagest&""",")
		Response.Write("pages:"&ipagecount&",")
		Response.Write("dcount:"&dcount&",")
		Response.Write("page:"&ipagecurrent&"")
		Response.Write("}]")
	end if
	rs.Close
	set rs=nothing
	CloseConn()
End Sub

Sub ediphoto()
	
	if at=0 then
		Errorc()
		exit Sub
	end if
	photoid=clng(Request.QueryString("photoid"))
	if photoid="" then
		Response.Write("[{errzt:0,errtx:""错误的id""}]")
		exit Sub
	end if
	edtname=trim(unescape(Request.QueryString("etname")))
	if edtname="" then
		Response.Write("[{errzt:0,errtx:""名称不能为空""}]")
		exit Sub
	end if
	sql="update blog_album set album_Title='"&edtname&"' where album_ID="&photoid
    conn.execute sql
	Response.Write("[{errzt:1,errtx:""0"",edname:"""&edtname&"""}]")
	call closeconn()
End Sub


Sub putphoto()
	
	if at=0 then
		Errorc()
		exit Sub
	end if
	photoid=clng(Request.QueryString("photoid"))
	if photoid="" then
		Response.Write("错误的分类id")
		exit Sub
	end if
	hots=Request("hots")
	sql="update blog_album set album_Hidden="&hots&" where album_ID="&photoid
    conn.execute sql
	Response.Write("设置成功")
	call closeconn()
End Sub

Sub comphoto()
	
	if at=0 then
		Errorc()
		exit Sub
	end if
	photoid=clng(Request.QueryString("photoid"))
	if photoid="" then
		Response.Write("错误的分类id")
		exit Sub
	end if
	coms=Request("coms")
	sql="update blog_album set album_Comment="&coms&" where album_ID="&photoid
    conn.execute sql
	Response.Write("设置成功")
	call closeconn()
End Sub


Function edophoto()
	
	if at=0 then
		Errorc()
		exit Function
	end if
	mname=request.Form("minname")
	if mname="" then
		Response.Write("[{Dezt:1,Detx:""名称不能为空""}]")
		exit Function
	end if
	classid=trim(request.Form("sel1"))
	photoid=trim(request("phoid"))
	picname=unescape(mname)
	dalesm=unescape(request.Form("suomin"))
	
	smotu=unescape(request.Form("smotu"))
	bigtu=unescape(request.Form("bigtu"))

	set rs=server.CreateObject("adodb.recordset")
	rs.open "select * from blog_album where album_ID="&photoid,conn,1,3
	rs("album_Title")=picname
	rs("album_class")=classid
		
	rs("album_Urlm")=smotu
	rs("album_Url")=bigtu
	
	rs("album_Messager")=HTMLEncode2(dalesm)
	rs.update
	rs.close
	set rs=nothing
	Response.Write("[{Dezt:1,Detx:""修改成功""}]")
	call closeconn()
End Function

Sub EditOthe()
	if at=0 then
		Response.Write("1")
		exit Sub
	end if
	photoid=clng(Request.QueryString("xgpid"))
	classid=clng(Request.QueryString("xgcid"))
	
	xiuid=photoid
	
	Response.write "<form id='myform' name='myform' method='post'  onsubmit=""Eddcheck('myform');return false;"">"
	Response.write "<div class='tjiao'>"
	Response.write "<ul><li><span class='tjiaospan'>选择大类：</span><label>"
	Response.write "<select name='sel1' >"
				set rs=server.createobject("adodb.recordset")
			rs.Open "select * from blog_album_class ORDER BY album_classID desc",conn,1,1
			do while not rs.eof
			if classid=rs("album_classID") then
	Response.write "<option value='"&rs("album_classID")&"' selected='selected'>"&rs("album_className")&"</option>"
			else
	Response.write "<option value='"&rs("album_classID")&"'>"&rs("album_className")&"</option>"
			end if
			rs.movenext
			loop
			rs.close
			set rs=nothing
	Response.write "</select><span id='spabtex1'></span></label></li>"
	

	
		set temprs=conn.execute("select album_Title,album_Messager,album_Urlm,album_Url from blog_album where album_ID="&photoid)
		tempname=temprs(0)
		tempjj=temprs(1)
		tempsmo=temprs(2)
		tempbig=temprs(3)
	
	set temprs=nothing
	Response.write "<li><span class='tjiaospan'>填写名称：</span><label><input id='minname' type='text' value='"&tempname&"' name='minname' /><input id='phoid' type='hidden' value='"&xiuid&"' name='phoid' /> *<span id='addspn'></span></label></li>"
	Response.write "<li><span class='tjiaospan'>小图路径：</span><label><input id='smotu' type='text' value='"&tempsmo&"' name='smotu' /><span id='smospan'></span></label></li>"
	Response.write "<li><span class='tjiaospan'>大图路径：</span><label><input id='bigtu' type='text' value='"&tempbig&"' name='bigtu' />"&tempinput&"<span id='bigspan'></span></label></li>"
	Response.write "<li><span class='tjiaospan'>填写说明：</span><label><textarea  name='suomin' id='suomin'>"&HTMLEncode1(tempjj)&"</textarea></label></li>"
	Response.write "<li><span class='anniu'>&nbsp;</span><label>"
	Response.write "<input type='submit' id='buttonm1' name='buttonm1' class='an1' value='修改' />   <input name='Reset1' type='reset' class='an1' value='重写' />"
	Response.write "<a href='javascript:void(0)' onclick='addshowout();' > <span class='red'>返回</span></a>"
	Response.write "</label></li></ul></div></form>"
call closeconn()
End Sub



%>

