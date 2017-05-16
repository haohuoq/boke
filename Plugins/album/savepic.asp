<!--#include file="conn.asp" -->
<!--#include file="sysfile/Function.asp" -->
<%
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.cachecontrol = "no-cache"

Server.ScriptTimeout = 9999

dim at
IF session("upadmin") ="guanli" Then
at=1
else 
at=0
end if
blogpath="Plugins/album/"

Select case Request("act")
case ""
  Call Errorc()
case "uploadpic"
  Call upmodule()
case "dphoto"
  Call Delphoto()
case else
  Call Errorc()
end Select

Sub Errorc()
	response.write("出错了")
End Sub

Sub upmodule()
	if at=0 then
		Upload_0()
		exit Sub
	end if
		Select Case (up_module)
			Case 0
			Upload_0()
			Case 1
			Upload_1()
			
		End Select
End Sub
Sub Upload_0()
response.CharSet = "UTF-8"
	tempmsg="<div class='tipxulog'><span class='red'>你还没有登陆或超时</span></div>"
	Response.write"<script>parent.document.getElementById(""upid"").innerHTML="""&tempmsg&"""</script>"

End Sub
''--------------------------------------------------------------------------------------------------------
''无组件上传
''--------------------------------------------------------------------------------------------------------
Sub Upload_1()
response.CharSet = "UTF-8"
	if at=0 then
		Upload_0()
		exit Sub
	end if
	
Dim upl,FSOIsOK
		FSOIsOK=1
Set upl=Server.CreateObject("Scripting.FileSystemObject")
		If Err<>0 Then
			Err.Clear
			FSOIsOK=0
		End If
		Dim D_Name,F_Name
		If FSOIsOK=1 Then
			D_Name="photo/month_"&DateToStr(Now(),"ym")
			If upl.FolderExists(Server.MapPath("photo"))=False Then
				upl.CreateFolder Server.MapPath("photo")
			End If
			If upl.FolderExists(Server.MapPath(D_Name))=False Then
				upl.CreateFolder Server.MapPath(D_Name)
			End If
			If upl.FolderExists(Server.MapPath(D_Name&"/small"))=False Then
				upl.CreateFolder Server.MapPath(D_Name&"/small")
			End If
		Else
			D_Name="All_Files"
		End If
		Set upl=Nothing
    
	
	set upload=new upload_file
	classid=upload.form("sel1")
	pcomment=0
	phidden=0
	rutime=date()
	i=1
	
		for each formName in upload.File
			picname= upload.form("tuname"&i&"")
			picsuom= upload.form("suomin"&i&"")
			if picsuom="" then picsuom=picname end if
			set file=upload.File(formName)
			fileext=lcase(file.FileExt)
			ftype=file.FileType
			if inStr( ftype, "image" ) = 0 then
 	    		tempmsg="<span class='red'>文件不是图片文件没有上传</span>"
           		pic=""
           		minipic=""
 	    	else
 				savename = FormatName(fileext)
				pic = D_Name&"/"&savename
				xiaopath=D_Name&"/small"	
				file.SaveToFile Server.mappath(pic)  
				minipic=xiaotu(pic,xiaopath,savename,picwidth,picheight,0,suiyin)
				tempmsg="<span>图片上传成功</span>"
				sql="insert into blog_album(album_Title,album_Messager,album_Hidden,album_Comment,album_class,album_Urlm,album_Url,album_Time) values('"&picname&"','"&picsuom&"','"&pcomment&"','"&phidden&"','"&classid&"','"&blogpath&""&minipic&"','"&blogpath&""&pic&"','"&rutime&"')"
				conn.execute sql
					    	
			end if
				pagest=pagest&"<li><div class='timxu2'>第"&i&"张处理完成</div><div class='tiuxu'>"
				pagest=pagest&"<div>缩略图:<span id=''><input type='text' size='42'  class='anb' value='"&minipic&"'/></span></div>"
				pagest=pagest&"<div>原始图:<span id=''><input type='text' size='42'  class='anb'  value='"&pic&"'/></span></div>"
				pagest=pagest&"</div><div class='tizxu2'>"&tempmsg&"</div></li>"
		 	i=i+1
		 	set file=nothing
		Next
	
	Set Upload=Nothing
Response.write"<script>parent.document.getElementById(""upid"").innerHTML="""&pagest&"""</script>"
End Sub

Sub Delphoto()
	response.CharSet = "UTF-8"
	if at=0 then
		 Errorc()
		exit Sub
	end if
	id=clng(Request.QueryString("id"))
	if id="" then
		Response.Write("错误的分类id")
		exit Sub
	end if
		set temprs=conn.execute("select album_Urlm,album_Url from blog_album where album_ID="&id)
	Fsodelwj(temprs(0))
	Fsodelwj(temprs(1))
	set temprs=nothing

	conn.execute ("delete from blog_album where album_ID="&id)
	Response.Write("删除成功")
	call closeconn()
End Sub

function Anewpics()
	response.CharSet = "utf-8"
	if at=0 then
		Response.Write("[{Dezt:0,Detx:""你还没有登陆""}]")
		exit function
	end if
	smotu=unescape(request("xpc"))
	bigtu=unescape(request("dpc"))
	savename="0"
	smpath=left(smotu,InStrRev(smotu,"/")-1)
	if createdir(server.MapPath(smpath)) = false then
		Response.Write("[{Dezt:0,Detx:""创建目录失败，请检查目录权限""}]")
		exit function
	end if	
	if ReportFileStatus(Server.MapPath(bigtu))=True then
		Select Case (img_module)
			Case 1
				minipic=xiaotu(bigtu,smotu,savename,picwidth,picheight,1,0)
			Case 2
				minipic=IDxiaotu(bigtu,smotu,savename,picwidth,picheight,1,0)
		End Select
		Response.Write("[{Dezt:1,Detx:""重做缩略图完成""}]")
	else
		Response.Write("[{Dezt:1,Detx:""大图并不存在""}]")
	end if
End function
%>