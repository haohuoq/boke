<!--#include file="../../const.asp" -->
<!--#include file="../../p_conn.asp" -->
<!--#include file="../../common/function.asp" -->
<!--#include file="../../common/library.asp" -->
<!--#include file="../../common/cache.asp" -->
<!--#include file="../../common/checkUser.asp" -->
<!--#include file="../../common/ModSet.asp" -->
<%
'=====================================
'  相册插件信息处理页面
'    更新时间: 2007-01-15
'=====================================
%>

<%
dim albumSet,OpenState,Getplugins,albumDBPath
Set albumSet=New ModSet
albumSet.open("album")
albumDBPath=albumSet.getKeyValue("Database")
	'打开数据库
Dim albumConn
Set albumConn=Server.CreateObject("ADODB.Connection")
albumConn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath(albumDBPath)&""
albumConn.Open


albumSet.open("album")
if not albumSet.PasreError<>-18903 then
   showmsg "错误信息","相册插件没有安装<br/><a href=""javascript:history.go(-1)"">单击返回</a>","MessageIcon","plugins"
end if
OpenState=albumSet.getKeyValue("start")
if not cBool(OpenState) then
   showmsg "错误信息","相册暂时关闭！<br/><a href=""../../default.asp"">单击返回首页</a>","WarningIcon","plugins"
end if

if request.form("action")="add" then
   addalbum '发表图片
 elseif Request.QueryString("action")="del" then 
   delalbum  '删除图片
 elseif Request.form("action")="edit" then 
   editalbum '编辑图片
 elseif Request.form("action")="class" then 
   aclass '编辑类型
 elseif request.form("action")="post" then
   PostBComm '发表评论
 elseif Request.QueryString("action")="delcomms" then 
   delcomm  '删除评论
 else
   showmsg "错误信息","非法操作！<br/><a href=""javascript:history.go(-1)"">单击返回</a>","ErrorIcon","plugins"
end if

'============================= 发表图片 ========================================
function addalbum
 dim album_user,post_Message,Title
 dim hiddenpic,album_Url,album_Urlm,aclass,Commentpic
'dim validate
  album_user=CheckStr(request.form("user"))
  post_Message=CheckStr(request.form("album_Messager"))
  Title=CheckStr(request.form("album_Title"))
  hiddenpic=request.form("hidden")
  Commentpic=request.form("Comment")
  album_Url=CheckStr(request.form("album_Url"))
  album_Urlm=CheckStr(request.form("album_Urlm"))
  aclass=CheckStr(request.form("aclass"))
if album_Urlm="" then
 album_Urlm=album_Url
end if

'  validate=trim(request.form("validate"))

  if Not stat_Admin then
     showmsg "相册错误信息","<b>你没权发布信息</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>","ErrorIcon","plugins"
	 exit function
  end if

'  IF cstr(lcase(Session("GetCode")))<>cstr(lcase(validate)) then
'      showmsg "相册错误信息","<b>验证码有误，请返回重新输入</b><br/><a href=""LoadMod.asp?plugins=GuestBookForPJBlog"">请返回重新输入</a>","ErrorIcon","plugins"
'      exit function
'  end if
  if len(Title)<1 then
     showmsg "相册错误信息","<b>不允许发表空图片标题</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>","WarningIcon","plugins"
	 exit function
  end if
  
  if len(post_Message)<1 then
     showmsg "相册错误信息","<b>不允许发表空图片信息</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>","WarningIcon","plugins"
	 exit function
  end if

  if len(album_Url)<1 then
     showmsg "相册错误信息","<b>图片地址不能为空</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>","WarningIcon","plugins"
	 exit function
  end if
if hiddenpic=1 then hiddenpic=true else hiddenpic=false
if Commentpic=1 then Commentpic=true else Commentpic=false
 '插入数据
albumConn.ExeCute("INSERT INTO blog_album(album_user,album_Title,album_Messager,album_Url,album_Urlm,album_Hidden,album_class,album_Comment) VALUES('"&album_user&"','"&Title&"','"&post_Message&"','"&album_Url&"','"&album_Urlm&"',"&hiddenpic&",'"&aclass&"',"&Commentpic&")")
SQLQueryNums=SQLQueryNums+1
Session(CookieName&"_Lastalbum")="Addphoto"
//getInfo(2)
showmsg "相册信息","<b>相片发布成功</b><br/><a  href=""LoadMod.asp?plugins=album"">单击返回相册分类</a><br/><a href=""LoadMod.asp?plugins=album&class="&aclass&""">单击返回图片所在相册</a>","MessageIcon","plugins" 
end function

'==================================== 删除图片 ===============================================
function delalbum
dim picUrl,picUrlm,album_ID,album_DB,SQL
album_ID=CheckStr(request.QueryString("id"))
set album_DB=Server.CreateObject("Adodb.Recordset")  
SQL="select * from blog_album where album_ID="&album_ID
album_DB.Open SQL,albumConn,1,1
  if album_DB.eof and album_DB.bof then
	  showmsg "错误信息","<b>此图片不存在!</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>","ErrorIcon","plugins"
  end if
picUrl="../../"&album_DB("album_Url")
picUrlm="../../"&album_DB("album_Urlm")
  if memName<>empty and stat_Admin then
if Instr(picUrl,"http://") = 0 then
     delpicfiles (Server.MapPath(picUrl))
     delpicfiles (Server.MapPath(picUrlm))
end if
album_DB.close
set album_DB=nothing
     albumConn.ExeCute("DELETE * FROM blog_album WHERE album_ID="&album_ID)      
     albumConn.ExeCute("DELETE * FROM blog_album_Comment WHERE album_ID="&album_ID)
	 Session(CookieName&"_Lastalbum")="Delphoto"
	// getInfo(2)
     showmsg "图片删除成功","<b>图片已经被删除!</b><br/><a href=""LoadMod.asp?plugins=album"">单击返回相册</a>","MessageIcon","plugins"
  else
     showmsg "错误信息","<b>你没有权限删除该图片</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>","ErrorIcon","plugins"
  end if
end function 
'==================================== 修改图片 ===============================================
function editalbum
  dim album_ID,post_Message,album_Url,album_Urlm,Title,hiddenpic,album_user,eaclass,Comments

  album_ID = CheckStr(Request.form("ID"))
  album_user=CheckStr(request.form("user"))
  post_Message=CheckStr(request.form("album_Messager"))
  album_Url=CheckStr(request.form("album_Url"))
  album_Urlm=CheckStr(request.form("album_Urlm"))
  Title=CheckStr(request.form("album_Title"))
  hiddenpic=request.form("hidden")
  eaclass=CheckStr(request.form("aclass"))
  Comments=request.form("Comment")
if album_Urlm="" then
 album_Urlm=album_Url
end if

  
  if hiddenpic=1 then hiddenpic=true else hiddenpic=false
  if Comments=1 then Comments=true else Comments=false
  if album_Urlm="" then
 album_Urlm=album_Url
end if

	   If len (album_Urlm)<1 then 
	            showmsg "错误信息","缩略图地址未填写<br/><a href=""javascript:history.go(-1)"">单击返回</a>","WarningIcon","plugins" 
	   end if
	   If len (album_Url)<1 then 
	            showmsg "错误信息","图片地址未填写<br/><a href=""javascript:history.go(-1)"">单击返回</a>","WarningIcon","plugins" 
	   end if
       if not stat_Admin then
                showmsg "错误信息","你没有权限修改图片信息<br/><a href=""javascript:history.go(-1)"">单击返回</a>","WarningIcon","plugins"
	   end if
	   If album_ID=Empty then 
	            showmsg "错误信息","非法操作<br/><a href=""javascript:history.go(-1)"">单击返回</a>","WarningIcon","plugins" 
	   end if
	   If IsInteger(album_ID)=False then 
	            showmsg "错误信息","非法操作<br/><a href=""javascript:history.go(-1)"">单击返回</a>","WarningIcon","plugins" 
	   end if

   albumConn.ExeCute("update blog_album set album_user='"&album_user&"',album_Url='"&album_Url&"',album_Urlm='"&album_Urlm&"',album_Title='"&Title&"',album_Messager='"&post_Message&"',album_Hidden="&hiddenpic&",album_Comment="&Comments&",album_class='"&eaclass&"',album_Time=#"&DateToStr(now(),"Y-m-d H:I:S")&"# WHERE album_ID="&album_ID)
   showmsg "修改图片信息","修改成功!<br/><a href=""LoadMod.asp?plugins=album"">单击返回相册分类</a><br/><a href=""LoadMod.asp?plugins=album&action=albumedit&id="&album_ID&""">单击返回你所编辑的图片</a>","MessageIcon","plugins" 
   Session(CookieName&"_Lastalbum")="Addphoto"
   //getInfo(2)
end function 
'=====================================编辑类型===================================================
function aclass

if not (memName<>empty and stat_Admin) then
    showmsg "错误信息","你没有权限修改图片信息<br/><a href=""javascript:history.go(-1)"">单击返回</a>","WarningIcon","plugins"
	exit function
else
Dim classIDs,classNames,i,classPics
	   if Request.form("a_class")="DelSelect" then
		    classIDs = split(Request.form("selectclassID"),", ")
		    for i=0 to ubound(classIDs)
				albumConn.execute("DELETE * from blog_album_class where album_classID="&classIDs(i))
			next
		    session(CookieName&"_ShowMsg")=true
		    session(CookieName&"_MsgText")=session(CookieName&"_MsgText")&(ubound(classIDs)+1)&" 个分类被删除!"
		    Response.Redirect("config.asp?menu=Class")
    	   else
		     classIDs=split(Request.form("classID"),", ")
		     classNames=split(Request.form("className"),", ")
		     classPics=split(Request.form("classPic"),", ")
			 for i=0 to ubound(classIDs)
		     if int(classIDs(i))<>-1 then
		        albumConn.execute("update blog_album_class set album_classPic='"&CheckStr(classPics(i))&"',album_className='"&CheckStr(classNames(i))&"' where album_classID="&classIDs(i))
		       else
		         if len(trim(CheckStr(classNames(i))))>0 then
		          albumConn.execute("insert into blog_album_class (album_className,album_classPic) values ('"&CheckStr(classNames(i))&"','"&CheckStr(classPics(i))&"')")
		          session(CookieName&"_MsgText")="新分类添加成功! "
		         end if
		     end if
		    next
		    session(CookieName&"_ShowMsg")=true
		    session(CookieName&"_MsgText")=session(CookieName&"_MsgText")&"分类保存成功!"
		    Response.Redirect("config.asp?menu=Class")
		end if
end if
end function
'=====================================删除图片文件===================================================
function delpicfiles(path)
dim picfile
Set picfile=Server.CreateObject("scripting.FileSystemObject")
if picfile.fileExists(path) then picfile.DeleteFile path
Set picfile=nothing
end function
'============================ 删除评论函数 =================================================
function delcomm
 dim post_commID,blog_Comm,blog_CommAuthor,comm
  post_commID=CheckStr(request.QueryString("commID"))
  set blog_Comm=albumConn.ExeCute("select top 1 album_CommentID,album_ID,album_CommentUSER from blog_album_Comment where album_CommentID="&post_commID)
  if blog_Comm.eof or blog_Comm.bof then
  showmsg "错误信息","<b>不存在此评论,或该评论已经被删除!</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>","WarningIcon","plugins"
  exit function
  end if
  blog_CommAuthor=blog_Comm("album_CommentUSER")
  comm=blog_Comm("album_ID")
  if stat_Admin=true or (stat_CommentDel=true and memName=blog_CommAuthor) then
     albumConn.ExeCute("DELETE * FROM blog_album_Comment WHERE album_CommentID="&post_commID)
     albumConn.ExeCute("update blog_album set commNUM=commNUM-1 where album_ID="&comm)
  showmsg "评论删除成功","<b>评论已经被删除成功!</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>","MessageIcon","plugins"
  else
  showmsg "错误信息","<b>你没有权限删除评论</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>","WarningIcon","plugins"
  end if

end function
'====================== 评论发表函数 ===========================================================
function PostBComm
 dim username,post_logID,post_Message,validate,password,FlowControl,LastMSG
  username=trim(CheckStr(request.form("username")))
  password=trim(CheckStr(request.form("password")))
  post_logID=CheckStr(request.form("albumID"))
  validate=trim(request.form("validate"))
  post_Message=CheckStr(request.form("Message"))
  FlowControl=false

  set LastMSG=albumConn.execute("select top 1 album_CommentMessager from blog_album_Comment where album_ID="&post_logID&" order by album_CommentID desc")
  if LastMSG.eof then
     FlowControl=false
   else
    if LastMSG("album_CommentMessager")=post_Message then FlowControl=true
  end if

  if FlowControl then 
	  showmsg "评论发表错误信息","<b>禁止恶意灌水！</b><br/><a href=""javascript:history.go(-1);"">返回</a>","WarningIcon","plugins"
      exit function 
  end if 

  if not stat_CommentAdd then
	 showmsg "评论发表错误信息","<b>你没有权限发表评论</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>","ErrorIcon","plugins"
	 exit function
  end if

  if albumConn.ExeCute("select album_Comment from blog_album where album_ID="&post_logID)(0) then 
	 showmsg "评论发表错误信息","<b>该日志不允许发表任何评论</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>","WarningIcon","plugins"
	 exit function
  end if 

  IF (memName=empty or blog_validate=true) and cstr(lcase(Session("GetCode")))<>cstr(lcase(validate)) then
	  showmsg "评论发表错误信息","<b>验证码有误，请返回重新输入</b><br/><a href=""javascript:history.go(-1);"">请返回重新输入</a>","ErrorIcon","plugins"
      exit function
  end if
  
  if filterSpam(post_Message,"../../spam.xml") and stat_Admin=false then
	  showmsg "评论发表错误信息","<b>评论中包含被屏蔽的字符</b><br/><a href=""javascript:history.go(-1);"">返回</a>","WarningIcon","plugins"
      exit function 
  end if

  if DateDiff("s",Request.Cookies(CookieName)("memLastPost"),Now())<blog_commTimerout then 
	  showmsg "评论发表错误信息","<b>发言太快,请 "&blog_commTimerout&" 秒后再发表评论</b><br/><a href=""javascript:history.go(-1);"">返回</a>","WarningIcon","plugins"
	  exit function  
  end if
  if len(username)<1 then
	  showmsg "评论发表错误信息","<b>请输入你的昵称.</b><br/><a href=""javascript:history.go(-1);"">请返回重新输入</a>","ErrorIcon","plugins"
      exit function  
  end if

  if IsValidUserName(username)=false then
	 showmsg "评论发表错误信息","<b>非法用户名！<br/>请尝试使用其他用户名！</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>","ErrorIcon","plugins"
	 exit function
 end if
  
  dim checkMem
  if memName=empty then
    if len(password)>0 then
        Dim loginUser
        loginUser=login(Request.Form("username"),Request.Form("password"))
         if not request.Cookies(CookieName)("memName")=username then
				 showmsg "评论发表错误信息","<b>登录失败，请检查用户名和密码</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>","WarningIcon","plugins"
            	 exit function
         end if
    else
       set checkMem=Conn.ExeCute("select top 1 mem_id from blog_Member where mem_Name='"&username&"'")
       if not checkMem.eof then
		 showmsg "评论发表错误信息","<b>该用户已经存在，无法发表评论</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>","WarningIcon","plugins"
    	 exit function
       end if
    end if
  end if

  if not stat_CommentAdd then
	 showmsg "评论发表错误信息","<b>你没有权限发表评论</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>","ErrorIcon","plugins"
	 exit function
  end if 

  if len(post_Message)<1 then
	 showmsg "评论发表错误信息","<b>不允许发表空评论</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>","ErrorIcon","plugins"  
	 exit function
  end if

  if len(post_Message)>blog_commLength then
	 showmsg "评论发表错误信息","评论超过最大字数限制<br/><a href=""javascript:history.go(-1);"">单击返回</a>","ErrorIcon","plugins"
	 exit function
  end if

'插入数据
 albumConn.ExeCute("INSERT INTO blog_album_Comment(album_ID,album_CommentMessager,album_CommentUSER,album_CommentIP) VALUES ("&post_logID&",'"&post_Message&"','"&username&"','"&getIP()&"')")
 albumConn.ExeCute("update blog_album set commNUM=commNUM+1 WHERE album_ID="&post_logID)
 SQLQueryNums=SQLQueryNums+1
 Response.Cookies(CookieName)("memLastpost")=Now()
 showmsg "评论发表成功","<b>你成功地对该相片发表了评论</b><br/><a href=""LoadMod.asp?plugins=album"">单击返回相册</a>","MessageIcon","plugins"
end function
%>
