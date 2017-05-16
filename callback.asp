<!--#include file="commond.asp" -->
<!--#include file="header.asp" -->
<!--#include file="common/sha1.asp" -->
<!--#include file="Plugins/QQconnect/oauth/class.asp"-->
<!--#include file="Plugins/QQconnect/oauth/config.asp"-->
<!--#include file="Plugins/QQconnect/oauth/openid.asp"-->
<!--内容-->
<%
'==================================
'  QQ帐号登录 for pjblog3
'    更新时间: 2011年6月7日
'    作者: 双木杉子 http://www.54bq.com
'==================================
If blog_Disregister Then showmsg "错误信息", "站点不允许注册新用户<br/><a href=""default.asp"">单击返回</a>", "ErrorIcon", ""
%>

 <div id="Tbody">
<%
Dim Reg:Reg=Array("","","")
Set WB=new WBoauth_54bq '新建类实例
Dim GET_Array,SIGN_Array,Gmethod,i,WB,API_URL,wb_page,wb_json,name,rett,oauth_vericode,openid,format,headpic,qqSql,mem_id
Dim firstD,secondD,thirdD
firstD=getopenid()
If firstD(0)=True Then
	openid=firstD(1)
	oauth_token=firstD(2)
	oauth_token_secret=firstD(3)
	mem_id=checkopenid(openid)
	If mem_id(0)=False Then
		secondD=getnickname(openid,oauth_token,oauth_token_secret)
		If secondD(0)=True Then
			Reg = register(secondD(1),secondD(2),openid)
		Else
	        Reg(0) = "错误信息"
	        Reg(1) = "<b>资料读取失败!</b><br/><a href=""javascript:void(0)"" onclick""location.reload()"">再试一次</a>或者<a href=""/"">返回首页</a>"
	        Reg(2) = "WarningIcon"
		End If	
	Else
		mem_idlogin(mem_id(1))
	End If
Else
        Reg(0) = "错误信息"
        Reg(1) = "<b>登录失败!</b><br/><a href=""/"">单击返回首页</a>再试一次"
        Reg(2) = "WarningIcon"
End If
ConnQQ.close
Set ConnQQ=Nothing
'读取QQ帐号对应openid
Function getopenid()
	Dim Arr:Arr=Array("","","","")
	'设置传递的参数需要手动排列成签名顺序
	GET_Array=Array("oauth_consumer_key","oauth_nonce","oauth_signature_method","oauth_timestamp","oauth_token","oauth_vericode","oauth_version")
	SIGN_Array=GET_Array
	'传参方式
	Gmethod="GET"
	'接口地址
	API_URL=access_token_url
	'读取Cookies
	oauth_token=WB.getcookie("oauth_token_first")
	oauth_token_secret=WB.getcookie("oauth_token_secret_first")
	'返回oauth_vericode
	oauth_vericode=Request.QueryString("oauth_vericode")
	'判断是否正常接入
	If oauth_token_secret="" Or oauth_token="" Then
        Reg(0) = "浏览器不支持Cookies"
        Reg(1) = "<b>请您开启浏览器Cookies，如果您正在使用手机浏览请进入一下地址</b><br/><a href=""/"">单击返回首页</a>再试一次"
        Reg(2) = "WarningIcon"
		getopenid=Arr
		Exit Function
	End If
	wb_page=WB.main()
	If left(wb_page,16)<>"oauth_signature=" then
        Reg(0) = "签名错误"
        Reg(1) = "<b>与QQ通信时出现技术错误，或许是插件后台设置错误</b><br/><a href='http://www.54bq.com' target='_blank'>寻求技术支持</a>"
        Reg(2) = "WarningIcon"
		getopenid=Arr
		Exit Function
	End If
	http_restore_query2(wb_page)
	If oauth_token<>"" And openid<>"" Then
		Arr(0)=True
		Arr(1)=openid
		Arr(2)=oauth_token
		Arr(3)=oauth_token_secret
		getopenid=Arr
	Else
		Arr(0)=False
		Arr(1)=""
		Arr(2)=""
		Arr(3)=""
		getopenid=Arr
	End If
End Function
'获取openid对应昵称
Function getnickname(openid,oauth_token,oauth_token_secret)
	Dim Arr:Arr=Array("","","","")
	'数据类型xml
	format="xml"
	'设置传递的参数需要手动排列成签名顺序
	GET_Array=Array("format","oauth_consumer_key","oauth_nonce","oauth_signature_method","oauth_timestamp","oauth_token","oauth_version","openid")
	SIGN_Array=GET_Array
	'传参方式
	Gmethod="GET"
	'接口地址
	API_URL=user_info_url
	If oauth_token_secret="" Or oauth_token="" Then
        Reg(0) = "注册失败，请重新登录再试"
        Reg(1) = "<b>注册不成功，可能是登录时间太长所致</b><br/><a href='default.asp' target='_blank'>返回再试一次</a>"
        Reg(2) = "WarningIcon"
		Exit Function
	End If
	wb_page=WB.main()
	if left(wb_page,1)<>"<" then
        Reg(0) = "读取资料失败"
        Reg(1) = "<b>与QQ链接成功但是授权失败，可能是网络原因</b><br/><a href=""javascript:void(0)"" onclick""location.reload()"">刷新再试一次</a>"
        Reg(2) = "WarningIcon"
		Exit Function
	End If
	Dim xmlDoc,nickname,nickname0,figureurl_2
	Set xmlDoc=Server.CreateObject("Microsoft.XMLDOM")
	xmlDoc.loadXML(wb_page)
	Set nickname0=xmlDoc.getElementsByTagName("nickname")
	Set figureurl_2=xmlDoc.getElementsByTagName("figureurl_2")
	nickname=nickname0.Item(0).Text
	headpic=figureurl_2.Item(0).Text
	If len(nickname)>0 Then
		Arr(0)=True
		Arr(1)=nickname
		Arr(2)=headpic
		getnickname=Arr
	Else
		Arr(0)=False
		Arr(1)=""
		Arr(2)=""
		Arr(3)=""
		getnickname=Arr
	End If
End Function
'检测openid是否已绑定
Function checkopenid(openid)
	Dim Arr:Arr=Array(False,"")
    Dim checktmp,RowsAffected
	Set checktmp = connQQ.Execute("select top 1 mem_id from qqopenid where openid='"&openid&"'")
	SQLQueryNums = SQLQueryNums + 1
	If Not checktmp.EOF Then
		ConnQQ.Execute("update qqopenid set oauth_token='"&oauth_token&"',oauth_token_secret='"&oauth_token_secret&"' where openid='"&openid&"'"),RowsAffected,&H0001 
        Arr(0)=True
		Arr(1)=checktmp(0)
		checkopenid=Arr
	Else 
		Arr(0)=False
		Arr(1)=""
		checkopenid=Arr
    End If
End Function
'用于二次登录
Function mem_idlogin(mem_id)
	Dim getmem_Name
	Set getmem_Name = conn.Execute("select top 1 mem_Name from blog_Member where mem_ID="&mem_ID&"")
    SQLQueryNums = SQLQueryNums + 1
	If getmem_Name.EOF Then
        Reg(0) = "意外错误"
        Reg(1) = "<b>该QQ注册过了，但是登录失败!</b><br/><a href=""/"">单击返回首页</a>再试一次"
        Reg(2) = "WarningIcon"		
        Exit Function
	Else
		Dim strSalt, AddUser,mem_Name, hashkey,password,password0,Referer_Url,Updated,RowsAffected
		password0=randomStr(6)
		hashkey = SHA1(randomStr(6)&Now())
		strSalt = randomStr(6)
		password = SHA1(password0&strSalt)
		mem_Name=getmem_Name(0)
        Conn.Execute("update blog_member set mem_Password='"&password&"',mem_salt='"&strSalt&"',mem_hashkey='"+hashkey+"' where mem_id="&mem_id&""),RowsAffected,&H0001 
	    SQLQueryNums = SQLQueryNums + 1
		If RowsAffected=0 Then
	        Reg(0) = "意外错误"
	        Reg(1) = "<b>该QQ注册过了，但是登录失败!</b><br/><a href=""/"">单击返回首页</a>再试一次"
	        Reg(2) = "WarningIcon"
		Exit Function
		End If
		Reg(0) = "欢迎回来"
		Referer_Url = Session(CookieName & "_Register_Referer_Url")
    If len(Referer_Url) < 8 Then
    	Reg(1) = "<b>"&mem_Name&"</b>欢迎回来，请记住新密码<b>"&password0&"</b> <br/><a href=""default.asp"">点击这里返回首页</a>"
    Else
    	Reg(1) = "<b>"&mem_Name&"欢迎回来，请记住新密码"&password0&"</b> <br/><a href="""&Referer_Url&""">返回登录前页面</a>&nbsp;|&nbsp;<a href=""default.asp"">返回首页</a><br/>"
    End If
    Reg(2) = "MessageIcon"
    Response.Cookies(CookieName)("memName") = mem_Name
    Response.Cookies(CookieName)("memHashKey") = hashkey
    Response.Cookies(CookieName).Expires = Date+365
    Session(CookieName&"_LastDo") = "RegisterUser"
	End If
End Function
%>
<br/><br/>
   <div style="text-align:center;">
    <div id="MsgContent" style="width:300px">
      <div id="MsgHead"><%=reg(0)%></div>
      <div id="MsgBody">
	   <div class="<%=reg(2)%>"></div>
       <div class="MessageText"><%=reg(1)%></div>
	  </div>
	</div>
  </div><br/><br/>
<%
Function register(username,headpics,openid)
    Dim ReInfo
	Dim qqname,password0,password,email, homepage, HideEmail, checkUser,Gender,qqopenid
    qqopenid=Trim(CheckStr(openid))
	ReInfo = Array("错误信息", "", "MessageIcon")
	qqname=username
    username = Trim(CheckStr(username))
    password0 = Get_nonce2(6)
    Gender = 1
    email = Trim(CheckStr(email))
    homepage = ""
        HideEmail = True

    If Len(username) = 0 Then
        ReInfo(0) = "错误信息"
        ReInfo(1) = "<b>请输入用户名(昵称)!</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
        ReInfo(2) = "WarningIcon"
        register = ReInfo
        Exit Function
    End If
    If Len(username)<2 Or Len(username)>24 Then
        ReInfo(0) = "错误信息"
        ReInfo(1) = "<b>您的QQ空间(昵称)小于2或<br/>大于24个字符，请修改后再试！</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
        ReInfo(2) = "ErrorIcon"
        register = ReInfo
        Exit Function
    End If
    If IsValidUserName(username) = False Then
        ReInfo(0) = "错误信息"
        ReInfo(1) = "<b>您的QQ空间（昵称）含有非法字符！<br/>请尝试修改后再试！</b><br/><a href=""javascript:history.go(-1);"">单击返回</a>"
        ReInfo(2) = "ErrorIcon"
        register = ReInfo
        Exit Function
    End If
    Set checkUser = connQQ.Execute("select top 1 mem_id from qqopenid where openid='"&qqopenid&"'")
    SQLQueryNums = SQLQueryNums + 1
	If Not checkUser.EOF Then
        ReInfo(0) = "错误信息"
        ReInfo(1) = "<b>该QQ号已经与本站关联！<br/>请稍等，正在为您登录……</b><br/>"&checkUser(0)
        ReInfo(2) = "MessageIcon"
        register = ReInfo
        Exit Function
    End If
	'当用户名已经存在时,用户名添加随机字符串
	Set checkUser = conn.Execute("select top 1 mem_id from blog_Member where mem_Name='"&username&"'")
    If Not checkUser.EOF Then
		username=username&Get_nonce2(5)
		'当添加随机字符串的用户名仍然存在时，提示用户刷新（几率在六千万分之一）当用户数量超过千万时请联系作者购买高级版插件  
    End If
	   Set checkUser = conn.Execute("select top 1 mem_id from blog_Member where mem_Name='"&username&"'")
	   If Not checkUser.EOF Then
		  ReInfo(0) = "错误信息"
		 ReInfo(1) = "<b>随机用户名失败！<br/></b><br/><a  href=""javascript:void(0)"" onclick""location.reload()"">请刷新一次尝试再次导入</a>"
		 ReInfo(2) = "ErrorIcon"
		 register = ReInfo
		 Exit Function
	   End If
    Dim strSalt, AddUser, hashkey
    hashkey = SHA1(randomStr(6)&Now())
    strSalt = randomStr(6)
    password = SHA1(password0&strSalt)
    AddUser = Array(Array("mem_Name", username), Array("mem_Password", password), Array("mem_Sex", Gender), Array("mem_salt", strSalt), Array("mem_Email", email), Array("mem_HideEmail", Int(HideEmail)), Array("mem_HomePage", homepage), Array("mem_LastIP", getIP), Array("mem_lastVisit", Now()), Array("mem_hashKey", hashkey))
    DBQuest "blog_member", AddUser, "insert"
    'Conn.Execute("INSERT INTO blog_member(mem_Name,mem_Password,mem_Sex,mem_salt,mem_Email,mem_HideEmail,mem_HomePage,mem_LastIP) Values ('"&username&"','"&password&"',"&Gender&",'"&strSalt&"','"&email&"',"&HideEmail&",'"&homepage&"','"&getIP&"')")
	Dim rsqq,User_ID
	Set rsqq = conn.Execute("select top 1 mem_id from blog_Member where mem_Name='"&username&"'")
    SQLQueryNums = SQLQueryNums + 1
    If Not checkUser.EOF Then
        ReInfo(0) = "错误信息"
        ReInfo(1) = "<b>意外错误，帐号注册失败！<br/>请再试一次……</b><br/>"
        ReInfo(2) = "MessageIcon"
        register = ReInfo
        Exit Function
	Else
	   User_ID=rsqq(0)
		'添加条目
		qqSql="INSERT INTO qqopenid(mem_ID,mem_Name,openid,oauth_token,oauth_token_secret,nickname,headpic) Values ("
		qqSql=qqSql&User_ID&",'"&username&"','"&openid&"','"&oauth_token&"','"&oauth_token_secret
		qqSql=qqSql&"','"&qqname&"','"&Trim(headpics)&"')"
		ConnQQ.Execute(qqSql)
	    SQLQueryNums = SQLQueryNums + 1
	End If
    Conn.Execute("UPDATE blog_Info SET blog_MemNums=blog_MemNums+1")
    getInfo(2)
    SQLQueryNums = SQLQueryNums + 2
    ReInfo(0) = "注册并登录成功"
    	ReInfo(1) = "<b>"&username&"</b>请牢记密码:<b>"+password0+"</b> <br/><a href=""default.asp"">点击这里返回首页</a>"
	ReInfo(2) = "MessageIcon"
    register = ReInfo
    Response.Cookies(CookieName)("memName") = username
    Response.Cookies(CookieName)("memHashKey") = hashkey
    Response.Cookies(CookieName).Expires = Date+365
    Session(CookieName&"_LastDo") = "RegisterUser"
End Function
%>
 </div>
<!--#include file="footer.asp" -->