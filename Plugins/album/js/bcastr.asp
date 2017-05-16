<!--#include file="conn.asp" -->
<%
Select case Request.QueryString("act")
	case ""
		Call Errorc()
	case "pai"
		Call Checkpai()
	case "huan"
		Call Checkhuan()
	case "llan"
		Call Checkllan()
	case "hot"
		Call Checkhot()
	case "seac"
		Call Checkseac()
	case "luclass"
		Call Checkluclass()
	case "lutype"
		Call Checklutype()
	case "fen"
		Call Checkfen()
	case "luminor"
		Call Checkluminor()
	case else
		Call Errorc()
end Select
Sub Errorc()
	response.write("")
End Sub
''---------------------------------------------------------------------------'class排行''
Sub Checkpai()
	response.CharSet = "GB2312"
	dim top
	top=0
	classid=clng(Request.QueryString("id"))
	sql = "select name,ck,id,minipic from desktop where classid="&classid&" order by ck desc"  
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.OPEN sql,Conn,0,1
		if rs.eof and rs.bof then
			Response.Write "<li class=""nopic"">还没有排行</li>"
		else
			do while top<10 and not rs.eof
			Response.Write "<li><a href='showpic.asp?id="&rs("id")&"' target='_blank' title='点击次数："&rs("ck")&"'>"&LeftStr(rs("name"),20)&"<span class='tim'>"&rs("ck")&"</span></a></li>"
			rs.movenext
			top=top+1
			loop
		end if
	rs.close
	Set rs=nothing
	call closeconn()
End Sub
''-----------------------------------------------------------------------------------''
Sub Checkhuan()
	response.CharSet = "GB2312"
	typeid=clng(Request.QueryString("typeid"))
	Dim ipagecount
	Dim ipagecurrent
	Dim strorderBy
	Dim irecordsshown
	if request.querystring("page")="" then
		ipagecurrent=1
	else
		ipagecurrent=clng(request.querystring("page"))
	end if

	sql = "select minipic,pic,ck,name,id from desktop where typeid="&typeid&" order BY id ASC" 
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.pagesize = 10
	rs.cachesize = 10
	rs.open sql,conn,1,1
	ipagecount = rs.pagecount
	If ipagecurrent > ipagecount Then ipagecurrent = ipagecount
	If ipagecurrent < 1 Then ipagecurrent = 1
	if ipagecount=0 then
		Response.Write("[{totalPages:0}]")
	else
		rs.absolutepage = ipagecurrent
		irecordsshown = 0
		i=0
		listt="<ul>"
		picurl=""
		picid=""
		do while irecordsshown<10 and NOT rs.EOF
			i=i+1
			picurl=picurl&""&rs("pic")&","
			picid=picid&""&rs("id")&","
			listt=listt&"<li id='tag"&i&"'><a href='javascript:check("&i&")'><img alt='"&rs("name")&"' src='"&rs("minipic")&"'></a></li>"
			irecordsshown = irecordsshown +1
			rs.movenext		
		loop
		listt=listt&"</ul>"
		Response.Write("[{")
		Response.Write("list:"""&listt&""",")
		Response.Write("picurl:"""&picurl&""",")
		Response.Write("picid:"""&picid&""",")
		Response.Write("pictool:"&i&",")
		Response.Write("page:"&ipagecurrent&",")
		Response.Write("totalPages:"&ipagecount&"") 
		Response.Write("}]")
	end if
	rs.Close
	set rs=nothing
	CloseConn()
End Sub
Sub Checkllan
	response.CharSet = "GB2312"
	tid=clng(Request.QueryString("typeid"))
	imgid=clng(Request.QueryString("id"))
	Set rss = Server.CreateObject("ADODB.Recordset")
	sql="select top 1 id,name,pic,jj from desktop where typeid="&tid&"and id<"&imgid&" order by id desc"
	rss.open sql,conn,1,1 
	if rss.eof and rss.bof then
		Pid=0
		Ppic=""
		Pname=""
	else
		Pid=rss("id")
		Ppic=rss("pic")
		Pname=rss("name")
		PimgExplain=rss("jj")
	end if
	rss.Close
	set rss=nothing
	Set rsx = Server.CreateObject("ADODB.Recordset")
	sql="select top 1 id,name,pic,jj from desktop where typeid="&tid&"and id>"&imgid&" order by id asc"
	rsx.open sql,conn,1,1
	if rsx.eof and rsx.bof then
		Nid=0
		Npic=""
		Nname=""
	else
		Nid=rsx("id")
		Npic=rsx("pic")
		Nname=rsx("name")
		NimgExplain=rsx("jj")
	end if
	rsx.Close
	set rsx=nothing
	Response.Write("[{name:"""&Pname&""", id:"&Pid&",pic:"""&Ppic&""",PimgE:"""&PimgExplain&"""},{name:"""&Nname&""",id:"&Nid&",pic:"""&Npic&""",NimgE:"""&NimgExplain&"""}]")
	sql="update desktop set ck=ck+1 where id="&imgid
	conn.execute sql
call closeconn()
End Sub
Sub Checkseac()
	response.CharSet = "GB2312"
	key=unescape(Request("ckey"))
	key=HTMLEncode(key)
	leilx=clng(Request("lei"))
	Dim ipagecount
	Dim ipagecurrent
	Dim strorderBy
	Dim irecordsshown  
	if request("page")="" then
		ipagecurrent=1
	else
		ipagecurrent=cint(request("page"))
	end if
	if leilx=0 then
		sql = "select minipic,name,id,typejs from type where name like '%"&trim(key)&"%' order by id desc" 
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.pagesize = 20
		rs.cachesize = 20
		rs.open sql,conn,1,1
		ipagecount = rs.pagecount
		if ipagecurrent > ipagecount Then ipagecurrent = ipagecount
		if ipagecurrent < 1 Then ipagecurrent = 1
		if ipagecount=0 then
			Response.Write("[{pagehtm:""<div class='picxno'>没 找 到 任 何 东 西 换 其 他 关 键 字 看 看<div></div></div>"",pages:0,dcount:0,page:0}]")
		else
			pagest=""
			rs.absolutepage = ipagecurrent
			irecordsshown = 0
			do while irecordsshown<20 and NOT rs.EOF
		
			pagest=pagest&"<div class='picx'>"
			pagest=pagest&"<div class='picmt'><a target='_blank' href='showtype.asp?typeid="&rs("id")&"'><img alt='' src='"&rs("minipic")&"' /></a></div>"
			pagest=pagest&"<div class='pictxt'><a target='_blank' href='showtype.asp?typeid="&rs("id")&"'>"&replace(rs("name"),trim(key), "<span class='red'>"&key&"</span>")&"<span class='txi'>"&rs("typejs")&"张</span></a></div></div>"
			irecordsshown = irecordsshown +1
			rs.movenext			
			loop
			Response.Write("[{pagehtm:"""&pagest&""",")
			Response.Write("pages:"&ipagecount&",")
			Response.Write("dcount:"&rs.recordcount&",")
			Response.Write("page:"&ipagecurrent&"")
			Response.Write("}]")
		end if
		rs.Close
		set rs=nothing
	else
		sql = "select id,name,minipic from minor where name like '%"&trim(key)&"%' order by id desc" 
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.pagesize = 20
		rs.cachesize = 20
		rs.open sql,conn,1,1
		ipagecount = rs.pagecount
		if ipagecurrent > ipagecount Then ipagecurrent = ipagecount
		if ipagecurrent < 1 Then ipagecurrent = 1
		if ipagecount=0 then
			Response.Write("[{pagehtm:""<div class='picxno'>没 找 到 任 何 东 西 换 其 他 关 键 字 看 看<div></div></div>"",pages:0,dcount:0,page:0}]")
		else
			pagest=""
			rs.absolutepage = ipagecurrent
			irecordsshown = 0
			do while irecordsshown<20 and NOT rs.EOF
		
			pagest=pagest&"<div class='picx'>"
			pagest=pagest&"<div class='picmt'><a target='_blank' href='showminor.asp?minorid="&rs("id")&"'><img alt='' src='"&rs("minipic")&"' /></a></div>"
			pagest=pagest&"<div class='pictxt'><a target='_blank' href='showminor.asp?minorid="&rs("id")&"'>"&replace(rs("name"),trim(key), "<span class='red'>"&key&"</span>")&"<span class='txi'>张</span></a></div></div>"
			irecordsshown = irecordsshown +1
			rs.movenext			
			loop
			Response.Write("[{pagehtm:"""&pagest&""",")
			Response.Write("pages:"&ipagecount&",")
			Response.Write("dcount:"&rs.recordcount&",")
			Response.Write("page:"&ipagecurrent&"")
			Response.Write("}]")
		end if
		rs.Close
		set rs=nothing
	end if
	CloseConn()
End Sub
'---------------------------------------------------------------------------------'大分类
Sub Checkluclass()
	response.CharSet = "GB2312"
	classid=Request("luid")
	Dim ipagecount
	Dim ipagecurrent
	Dim strorderBy
	Dim irecordsshown  
	if request("page")="" then
		ipagecurrent=1
	else
		ipagecurrent=cint(request("page"))
	end if
	sql = "select id,name,minipic from minor where classid="&classid&" order by id desc" 
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.pagesize = 20
	rs.cachesize = 20
	rs.open sql,conn,1,1
	dcount=rs.recordcount
	ipagecount = rs.pagecount
	if ipagecurrent > ipagecount Then ipagecurrent = ipagecount
	if ipagecurrent < 1 Then ipagecurrent = 1
	if ipagecount=0 then
		Response.Write("[{pagehtm:""<div class='mainpiczt'>没 有 任 何 图 片</div>"",pages:0,dcount:0,page:0}]")
	else
		pagest=""
		rs.absolutepage = ipagecurrent
		irecordsshown = 0
		do while irecordsshown<20 and NOT rs.EOF
			pagest=pagest&"<div class='picx'>"
			pagest=pagest&"<div class='picmt'><a href='showminor.asp?minorid="&rs("id")&"&fpage="&ipagecurrent&"'><img alt='' src='"&rs("minipic")&"' /></a></div>"
			pagest=pagest&"<div class='pictxt'style='text-align:center'><a href='showminor.asp?minorid="&rs("id")&"&fpage="&ipagecurrent&"'>"&rs("name")&"</span></a></div></div>"
			irecordsshown = irecordsshown +1
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
'---------------------------------------------------------------------------------'小分类
Sub Checkluminor()
	response.CharSet = "GB2312"
	minorid=Request("luid")
	Dim ipagecount
	Dim ipagecurrent
	Dim strorderBy
	Dim irecordsshown  
	if request("page")="" then
		ipagecurrent=1
	else
		ipagecurrent=cint(request("page"))
	end if
	sql = "select id,name,minipic,typejs from type where minorid="&minorid&" order by id desc" 
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.pagesize = 20
	rs.cachesize = 20
	rs.open sql,conn,1,1
	dcount=rs.recordcount
	ipagecount = rs.pagecount
	if ipagecurrent > ipagecount Then ipagecurrent = ipagecount
	if ipagecurrent < 1 Then ipagecurrent = 1
	if ipagecount=0 then
		Response.Write("[{pagehtm:""<div class='mainpiczt'>没 有 任 何 图 片</div>"",pages:0,dcount:0,page:0}]")
	else
		pagest=""
		rs.absolutepage = ipagecurrent
		irecordsshown = 0
		do while irecordsshown<20 and NOT rs.EOF
			pagest=pagest&"<div class='picx'>"
			pagest=pagest&"<div class='picmt'><a href='showtype.asp?typeid="&rs("id")&"&fpage="&ipagecurrent&"'><img alt='' src='"&rs("minipic")&"' /></a></div>"
			pagest=pagest&"<div class='pictxt'><a href='showtype.asp?typeid="&rs("id")&"&fpage="&ipagecurrent&"'>"&LeftStr(rs("name"),8)&"<span class='txi'>:"&rs("typejs")&"张</span></a></div></div>"
			irecordsshown = irecordsshown +1
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
'-------------------------------------------------------------------------------'全部分类
Sub Checkfen()
	response.CharSet = "GB2312"
	classid=clng(Request.QueryString("classid"))
	sql = "select id,name from minor where classid="&classid&" order by id desc"  
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sql,Conn,1,1
		if rs.eof and rs.bof then
			Response.Write "<div class='wholezt'>还没有分类</div>"
		else
			Response.Write "<div class='wholeleft'>共：<span style='color:red'>"&rs.recordcount&"</span> 辑</div>"
			Response.Write "<div class='wholeall'>"
			do while  not rs.eof
				Response.Write "<div class='wholefen'><a href='showminor.asp?minorid="&rs("id")&"' target='_blank'>"&LeftStr(rs("name"),9)&"<span class='tum'></span></a></div>"
			rs.movenext
			loop
			Response.Write "</div><div class='cls'></div></div>"
		end if
	rs.close
	Set rs=nothing
	call closeconn()
End Sub
''-----------------------------------------------------------------------------------'type分类
Sub Checklutype()
	response.CharSet = "GB2312"
	typeid=Request("luid")
	Dim ipagecount
	Dim ipagecurrent
	Dim strorderBy
	Dim irecordsshown  
	if request("page")="" then
		ipagecurrent=1
	else
		ipagecurrent=cint(request("page"))
	end if
	sql = "select minipic,ck,name,id from desktop where typeid="&typeid&" order by id desc" 
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.pagesize = 20
	rs.cachesize = 20
	rs.open sql,conn,1,1
	ipagecount = rs.pagecount
	if ipagecurrent > ipagecount Then ipagecurrent = ipagecount
	if ipagecurrent < 1 Then ipagecurrent = 1
	if ipagecount=0 then
		Response.Write("[{pagehtm:""<div class='mainpiczt'>没 有 任 何 图 片</div>"",pages:0,dcount:0,page:0}]")
	else
		pagest=""
		rs.absolutepage = ipagecurrent
		irecordsshown = 0
		do while irecordsshown<20 and NOT rs.EOF
		
		pagest=pagest&"<div class='picx'>"
		pagest=pagest&"<div class='picmt'><a target='_blank' href='showpic.asp?id="&rs("id")&"'><img alt='' src='"&rs("minipic")&"' /></a></div>"
		pagest=pagest&"<div class='pictxt'><a target='_blank' href='showpic.asp?id="&rs("id")&"'>"&LeftStr(rs("name"),20)&"<span class='txi'>点击:"&rs("ck")&"</span></a></div></div>"
		irecordsshown = irecordsshown +1
		rs.movenext			
		loop
		Response.Write("[{pagehtm:"""&pagest&""",")
		Response.Write("pages:"&ipagecount&",")
		Response.Write("dcount:"&rs.recordcount&",")
		Response.Write("page:"&ipagecurrent&"")
		Response.Write("}]")
	end if
	rs.Close
	set rs=nothing
	CloseConn()
End Sub
Sub Checkhot()
	id=clng(Request.QueryString("id"))
	sql="update desktop set ck=ck+1 where id="&id
	conn.execute sql	
End Sub
function LeftStr(Str,Count)
	if len(Str)>Count then
		LeftStr = left(Str,count) & "..."
	else
		LeftStr = Str
	end if
end function
Function HTMLEncode(Str)
 If Isnull(Str) Then
	 HTMLEncode = ""
	 Exit Function 
 End If
 Str = Replace(Str,Chr(0),"", 1, -1, 1)
 Str = Replace(Str, """", "&quot;", 1, -1, 1)
 Str = Replace(Str,"<","&lt;", 1, -1, 1)
 Str = Replace(Str,">","&gt;", 1, -1, 1) 
 Str = Replace(Str, "script", "&#115;cript", 1, -1, 0)
 Str = Replace(Str, "SCRIPT", "&#083;CRIPT", 1, -1, 0)
 Str = Replace(Str, "Script", "&#083;cript", 1, -1, 0)
 Str = Replace(Str, "script", "&#083;cript", 1, -1, 1)
 Str = Replace(Str, "object", "&#111;bject", 1, -1, 0)
 Str = Replace(Str, "OBJECT", "&#079;BJECT", 1, -1, 0)
 Str = Replace(Str, "Object", "&#079;bject", 1, -1, 0)
 Str = Replace(Str, "object", "&#079;bject", 1, -1, 1)
 Str = Replace(Str, "applet", "&#097;pplet", 1, -1, 0)
 Str = Replace(Str, "APPLET", "&#065;PPLET", 1, -1, 0)
 Str = Replace(Str, "Applet", "&#065;pplet", 1, -1, 0)
 Str = Replace(Str, "applet", "&#065;pplet", 1, -1, 1)
 Str = Replace(Str, "[", "&#091;")
 Str = Replace(Str, "]", "&#093;")
 Str = Replace(Str, """", "", 1, -1, 1)
 Str = Replace(Str, "=", "&#061;", 1, -1, 1)
 Str = Replace(Str, "'", "''", 1, -1, 1)
 Str = Replace(Str, "select", "sel&#101;ct", 1, -1, 1)
 Str = Replace(Str, "execute", "&#101xecute", 1, -1, 1)
 Str = Replace(Str, "exec", "&#101xec", 1, -1, 1)
 Str = Replace(Str, "join", "jo&#105;n", 1, -1, 1)
 Str = Replace(Str, "union", "un&#105;on", 1, -1, 1)
 Str = Replace(Str, "where", "wh&#101;re", 1, -1, 1)
 Str = Replace(Str, "insert", "ins&#101;rt", 1, -1, 1)
 Str = Replace(Str, "delete", "del&#101;te", 1, -1, 1)
 Str = Replace(Str, "update", "up&#100;ate", 1, -1, 1)
 Str = Replace(Str, "like", "lik&#101;", 1, -1, 1)
 Str = Replace(Str, "drop", "dro&#112;", 1, -1, 1)
 Str = Replace(Str, "create", "cr&#101;ate", 1, -1, 1)
 Str = Replace(Str, "rename", "ren&#097;me", 1, -1, 1)
 Str = Replace(Str, "count", "co&#117;nt", 1, -1, 1)
 Str = Replace(Str, "chr", "c&#104;r", 1, -1, 1)
 Str = Replace(Str, "mid", "m&#105;d", 1, -1, 1)
 Str = Replace(Str, "truncate", "trunc&#097;te", 1, -1, 1)
 Str = Replace(Str, "nchar", "nch&#097;r", 1, -1, 1)
 Str = Replace(Str, "char", "ch&#097;r", 1, -1, 1)
 Str = Replace(Str, "alter", "alt&#101;r", 1, -1, 1)
 Str = Replace(Str, "cast", "ca&#115;t", 1, -1, 1)
 Str = Replace(Str, "exists", "e&#120;ists", 1, -1, 1)
 Str = Replace(Str,Chr(13),"<br>", 1, -1, 1)
 HTMLEncode = Replace(Str,"'","''", 1, -1, 1)
End Function

%>