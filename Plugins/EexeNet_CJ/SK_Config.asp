﻿<%
'==================================================================================================
' 软件名称：SK信息采集管理系统
' 当前版本：3.1 SP1 Build070314
' 更新日期：2007-3-14
' 程序版权：SK网络
' 程序开发：SK网络开发组（总策划：沈志昌）
' 演示站点：http://www.skxiu.com/cj
' 官方网站：http://www.skxiu.com  QQ：85103270 电话：0596-2821043
' 郑重声明:
'    ①、免费版本请在程序首页保留版权信息，并做上本站LOGO友情连接,商业版本无此要求.
'    ②、任何个人或组织不得删除、修改、拷贝本软件及其他副本上一切关于版权的信息.
'    ③、SK网络保留此软件的法律追究权利.
'===================================================================================================
%>
<!--#include file="inc/setup.asp"-->
<!--#include file="SK_Session.asp"-->
<link rel="stylesheet" type="text/css" href="css/Admin_Style.css">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<script src="inc/common.js" language="javascript"></script>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
Action=Trim(Request("Action"))
if action="save" then
	call SaveEdit()
else
	call Main()'显示页面
end if

'关闭数据库链接
Call CloseConnItem()

'保存
Sub SaveEdit()
		ConnItem.Execute("Update Sk_cj Set Dir='"& Trim(Request.Form("ArticleDir")) &"' Where Id=1")
		ConnItem.Execute("Update Sk_cj Set Dir='"& Trim(Request.Form("photoDir")) &"' Where Id=2")
		ConnItem.Execute("Update Sk_cj Set Dir='"& Trim(Request.Form("DownDir")) &"' Where Id=3")
		SqlItem="select top 1 * from SK_Config"
     	Set Rs=server.CreateObject("adodb.recordset")
     	Rs.Open SqlItem,ConnItem,3,3
		rs("WebName")=Request.Form("WebName")
		rs("WebUrl")=Request.Form("WebUrl")
		rs("WebLogo")=Request.Form("WebLogo")
		rs("Webabout")=Request.Form("Webabout")
			Response.Flush()
			If G("RateTF") = "" Then
				RS("RateTF") = 0
			Else
				RS("RateTF") = G("RateTF")
			End If
			If G("ThumbWidth") = "" Then
				RS("ThumbWidth") = 0
			Else
				RS("ThumbWidth") = G("ThumbWidth")
			End If
			If G("ThumbHeight") = "" Then
				RS("ThumbHeight") = 0
			Else
				RS("ThumbHeight") = G("ThumbHeight")
			End If
			If G("ThumbRate") = "" Then
				RS("ThumbRate") = 0
			Else
				RS("ThumbRate") = G("ThumbRate")
			End If
			If G("MarkComponent") = "" Then
				RS("MarkComponent") = 0
			Else
				RS("MarkComponent") = G("MarkComponent")
			End If
			RS("MarkType") = G("MarkType")
			If G("MarkText") = "" Then
				RS("MarkText") = " "
			Else
				RS("MarkText") = G("MarkText")
			End If
			If G("MarkFontSize") = "" Then
				RS("MarkFontSize") = 12
			Else
				RS("MarkFontSize") = G("MarkFontSize")
			End If
			If G("MarkFontColor") = "" Then
				RS("MarkFontColor") = " "
			Else
				RS("MarkFontColor") = G("MarkFontColor")
			End If
			RS("MarkFontName") = G("MarkFontName")
			RS("MarkFontBond") = G("MarkFontBond")
			If G("MarkPicture") = "" Then
			RS("MarkPicture") = " "
			Else
			RS("MarkPicture") = G("MarkPicture")
			End If
			If G("MarkOpacity") = "" Then
				RS("MarkOpacity") = 0
			Else
				RS("MarkOpacity") = CSng(G("MarkOpacity"))
			End If
			If G("MarkWidth") = "" Then
				RS("MarkWidth") = 0
			Else
				RS("MarkWidth") = G("MarkWidth")
			End If
			If G("MarkHeight") = "" Then
				RS("MarkHeight") = 0
			Else
				RS("MarkHeight") = G("MarkHeight")
			End If
			If G("MarkTranspColor") = "" Then
				RS("MarkTranspColor") = " "
			Else
				RS("MarkTranspColor") = G("MarkTranspColor")
			End If
		
			RS("ThumbComponent") = G("ThumbComponent")
			RS("MarkPosition") = G("MarkPosition")
		Rs.UpDate
      	Rs.Close	
	  	set rs=nothing 
		response.write "<script>alert('修改设置成功!');location.href='sk_config.asp';</script>"'关闭窗口
    	call Main()
end sub
%>



<%Sub Main()
		SqlStr = "select * from SK_Config"
		Set RS = Server.CreateObject("adodb.recordset")
		RS.Open SqlStr, ConnItem, 1, 3
%>

<table width="97%" border="0" align="center" cellpadding="3" cellspacing="1" class="tableBorder">
<form   name="form"  action="?action=save" method="post">
 <tr>
    <td height="33" colspan="3"  style="background:#0D4d7d"><div align="left"><!--#include file="CJ_Pjblog.asp"--></div></td>
  </tr>  <tr>
    <td height="33" colspan="3" class="title"><div align="center"><span style="font-weight: bold">网站基本设置</span></div></td>
  </tr>
  <tr>
    <td colspan="3" class="title2">&nbsp;用于设置网站的基本信息。</td>
    </tr>
  
  <tr>
    <td width="18%" class="table">&nbsp;网站全称</td>
    <td colspan="2" class="table"><input name="WebName"  type="text" id="WebName" class="lostfocus" gf="0" onmouseover='this.className="getfocus"' onmouseout='if (this.gf=="0") this.className="lostfocus"' onblur='this.className="lostfocus";this.gf="0"' onfocus='this.className="getfocus";this.gf="1"'  value="<% =rs("WebName") %>" >
      （显示于首页。如 <a href="http://eexe.net">http://eexe.net</a>”）</td>
  </tr>
  

  

  <tr>
    <td colspan="3" class="title2">&nbsp;</td>
  </tr>
  
  
  
  <tr>
    <td colspan="3" class="title"><span style="font-weight: bold">缩略图水印设置</span></td>
  </tr>
  <tr>
    <td colspan="3" class="table">
	<%
			Response.Write "<body bgcolor=""#FFFFFF"" scroll=yes onLoad=""SetTypeArea(" & RS("MarkType") & ");ShowInfo(" & RS("MarkComponent") & ");ShowThumbInfo(" & RS("ThumbComponent") & ");ShowThumbSetting(" & RS("RateTF") & ")"" >"
		Response.Write "    <tr bgcolor=""#F5F5F5"" >"
		Response.Write "      <td width=""257"" height=""40"" align=""right"" class=""table""><STRONG>生成缩略图组件：</STRONG><BR>"
		Response.Write "      <span class=""STYLE1"">请一定要选择服务器上已安装的组件</span></td>"
		Response.Write "      <td width=""677"" class=""table"">"
		Response.Write "       <select name=""ThumbComponent"" id=""ThumbComponent"" onChange=""ShowThumbInfo(this.value)"" style=""width:50%"">"
		Response.Write "          <option value=0 "
		If RS("ThumbComponent") = "0" Then Response.Write ("selected")
		Response.Write ">关闭 </option>"
		Response.Write "          <option value=1 "
		If RS("ThumbComponent") = "1" Then Response.Write ("selected")
		Response.Write ">AspJpeg组件 " & ExpiredStr(0) & "</option>"
		Response.Write "          <option value=2 "
		If RS("ThumbComponent") = "2" Then Response.Write ("selected")
		Response.Write ">wsImage组件 " & ExpiredStr(1) & "</option>"
		Response.Write "          <option value=3 "
		If RS("ThumbComponent") = "3" Then Response.Write ("selected")
		Response.Write ">SA-ImgWriter组件 " & ExpiredStr(2) & "</option>"
		Response.Write "        </select>"
		Response.Write "      <span id=""ThumbComponentInfo""></span></td>"
		Response.Write "    </tr>"
		Response.Write "    <tr bgcolor=""#F5F5F5"" id =""ThumbSetting"" style=""display:none"">"
		 Response.Write "     <td height=""23"" align=""right"" class=""table""> <input type=""radio"" name=""RateTF"" value=""1"" onClick=""ShowThumbSetting(1);"" "
		 If RS("RateTF") = "1" Then Response.Write ("checked")
		 Response.Write ">"
		 Response.Write "       按比例"
		 Response.Write "       <input type=""radio"" name=""RateTF"" value=""0"" onClick=""ShowThumbSetting(0);"" "
		 If RS("RateTF") = "0" Then Response.Write ("checked")
		 Response.Write ">"
		 Response.Write "     按大小 </td>"
		 Response.Write "     <td width=""677"" height=""50"" class=""table""> <div id =""ThumbSetting0"" style=""display:none"">&nbsp;缩略图宽度："
		Response.Write "          <input type=""text"" name=""ThumbWidth"" size=10 value=""" & RS("ThumbWidth") & """>"
		Response.Write "          象素<br>&nbsp;缩略图高度："
		Response.Write "          <input type=""text"" name=""ThumbHeight"" size=10 value=""" & RS("ThumbHeight") & """>"
		Response.Write "          象素</div>"
		Response.Write "        <div id =""ThumbSetting1"" style=""display:none"">&nbsp;比例："
		Response.Write "          <input type=""text"" name=""ThumbRate"" size=10 value="""
		If Left(RS("ThumbRate"), 1) = "." Then Response.Write ("0" & RS("ThumbRate")) Else Response.Write (RS("ThumbRate"))
		Response.Write """>"
		Response.Write "      <br>&nbsp;如缩小原图的50%,请输入0.5 </div></td>"
		Response.Write "    </tr>"
		Response.Write "    <tr bgcolor=""#F5F5F5"" >"
		Response.Write "      <td height=""40"" align=""right"" class=""table""><strong>图片水印组件：</strong><BR>"
		Response.Write "      <span class=""STYLE1"">请一定要选择服务器上已安装的组件</span></td>"
		Response.Write "      <td width=""677"" class=""table""> <select name=""MarkComponent"" id=""MarkComponent"" onChange=""ShowInfo(this.value)"" style=""width:50%"">"
		Response.Write "          <option value=0 "
		If RS("MarkComponent") = "0" Then Response.Write ("selected")
		Response.Write ">关闭"
		Response.Write "          <option value=1 "
		If RS("MarkComponent") = "1" Then Response.Write ("selected")
		Response.Write ">AspJpeg组件 " & ExpiredStr(0) & "</option>"
		Response.Write "          <option value=2 "
		If RS("MarkComponent") = "2" Then Response.Write ("selected")
		Response.Write ">wsImage组件 " & ExpiredStr(1) & "</option>"
		Response.Write "          <option value=3 "
		If RS("MarkComponent") = "3" Then Response.Write ("selected")
		Response.Write ">SA-ImgWriter组件 " & ExpiredStr(2) & "</option>"
		Response.Write "      </select>  </td>"
		Response.Write "    </tr>"
		Response.Write "    <tr align=""left"" valign=""top"" bgcolor=""#F5F5F5"" id=""WaterMarkSetting"" style=""display:none"" cellpadding=""0"" cellspacing=""0"">"
		Response.Write "      <td colspan=2 class=""table""> <table width=100% border=""0"" cellpadding=""0"" cellspacing=""1""  bordercolor=""e6e6e6"" bgcolor=""#efefef"">"
		Response.Write "          <tr bgcolor=""#FFFFFF"">"
		Response.Write "            <td width=262 height=""26"" align=""right"" class=""table"">水印类型</td>"
		Response.Write "            <td width=""648"" class=""table""> <SELECT name=""MarkType"" id=""MarkType"" onChange=""SetTypeArea(this.value)"">"
		Response.Write "                <OPTION value=""1"" "
		If RS("MarkType") = "1" Then Response.Write ("selected")
		Response.Write ">文字效果</OPTION>"
		Response.Write "                <OPTION value=""2"" "
		If RS("MarkType") = "2" Then Response.Write ("selected")
		Response.Write ">图片效果</OPTION>"
		Response.Write "            </SELECT> </td>"
		Response.Write "          </tr>"
		Response.Write "          <tr bgcolor=""#FFFFFF"">"
		Response.Write "            <td height=""26"" align=""right"" class=""table"">坐标起点位置</td>"
		Response.Write "            <td class=""table""> <SELECT NAME=""MarkPosition"" id=""MarkPosition"">"
		Response.Write "                <option value=""1"" "
		If RS("MarkPosition") = "1" Then Response.Write ("selected")
		Response.Write ">左上</option>"
		Response.Write "                <option value=""2"" "
		If RS("MarkPosition") = "2" Then Response.Write ("selected")
		Response.Write ">左下</option>"
		Response.Write "                <option value=""3"" "
		If RS("MarkPosition") = "3" Then Response.Write ("selected")
		Response.Write ">居中</option>"
		Response.Write "                <option value=""4"" "
		If RS("MarkPosition") = "4" Then Response.Write ("selected")
		Response.Write ">右上</option>"
		Response.Write "                <option value=""5"" "
		If RS("MarkPosition") = "5" Then Response.Write ("selected")
		Response.Write ">右下</option>"
		Response.Write "            </SELECT> </td>"
		Response.Write "          </tr>"
		Response.Write "          <tr>"
		Response.Write "           <td colspan=""2"">"
		Response.Write "           <table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""1"" bgcolor=""#efefef"" id=""textarea"">"
		Response.Write "          <tr bgcolor=""#FFFFFF"">"
		Response.Write "            <td width=""27%"" height=""26"" align=""right"" class=""table"">水印文字信息:</td>"
		Response.Write "            <td width=""70%"" class=""table""> <INPUT TYPE=""text"" NAME=""MarkText"" size=40 value=""" & RS("MarkText") & """>            </td>"
		Response.Write "          </tr>"
		Response.Write "          <tr bgcolor=""#FFFFFF"">"
		Response.Write "            <td height=""26"" align=""right"" class=""table"">字体大小:</td>"
		Response.Write "            <td class=""table""> <INPUT TYPE=""text"" NAME=""MarkFontSize"" size=10 value=""" & RS("MarkFontSize") & """>"
		Response.Write "            <b>px</b> </td>"
		Response.Write "          </tr>"
		Response.Write "          <tr bgcolor=""#FFFFFF"">"
		Response.Write "            <td height=""26"" align=""right"" class=""table"">字体颜色:</td>"
		Response.Write "            <td class=""table""><input  type=""text"" name=""MarkFontColor"" maxlength = 7 size = 7 id=""MarkFontColor"" value=""" & RS("MarkFontColor") & """ readonly>"
		Response.Write "            <input type=""button"" name=""button12"" value=""选择颜色..."" onClick=""OpenThenSetValue('inc/SelectColor.asp',230,190,window,document.form.MarkFontColor);document.form.MarkFontColor.style.color=document.form.MarkFontColor.value;""></td>"
		Response.Write "          </tr>"
		Response.Write "          <tr bgcolor=""#FFFFFF"">"
		Response.Write "            <td height=""26"" align=""right"" class=""table"">字体名称:</td>"
		Response.Write "            <td class=""table""> <SELECT name=""MarkFontName"" id=""MarkFontName"">"
		Response.Write "                <option value=""宋体"" "
		If RS("MarkFontName") = "宋体" Then Response.Write ("selected")
		Response.Write ">宋体</option>"
		Response.Write "                <option value=""楷体_GB2312"" "
		If RS("MarkFontName") = "楷体_GB2312" Then Response.Write ("selected")
		Response.Write ">楷体</option>"
		Response.Write "                <option value=""新宋体"" "
		If RS("MarkFontName") = "新宋体" Then Response.Write ("selected")
		Response.Write ">新宋体</option>"
		Response.Write "                <option value=""黑体"" "
		If RS("MarkFontName") = "黑体" Then Response.Write ("selected")
		Response.Write ">黑体</option>"
		Response.Write "                <option value=""隶书"" "
		If RS("MarkFontName") = "隶书" Then Response.Write ("selected")
		Response.Write ">隶书</option>"
		Response.Write "                <OPTION value=""Andale Mono"" "
		If RS("MarkFontName") = "Andale Mono" Then Response.Write ("selected")
		Response.Write ">Andale"
		Response.Write "                Mono</OPTION>"
		Response.Write "                <OPTION value=""Arial"" "
		If RS("MarkFontName") = "Arial" Then Response.Write ("selected")
		Response.Write ">Arial</OPTION>"
		Response.Write "                <OPTION value=""Arial Black"" "
		If RS("MarkFontName") = "Arial Black" Then Response.Write ("selected")
		Response.Write ">Arial"
		Response.Write "                Black</OPTION>"
		Response.Write "                <OPTION value=""Book Antiqua"" "
		If RS("MarkFontName") = "Book Antiqua" Then Response.Write ("selected")
		Response.Write ">Book"
		Response.Write "                Antiqua</OPTION>"
		Response.Write "                <OPTION value=""Century Gothic"" "
		If RS("MarkFontName") = "Century Gothic" Then Response.Write ("selected")
		Response.Write ">Century"
		Response.Write "                Gothic</OPTION>"
		Response.Write "                <OPTION value=""Comic Sans MS"" "
		If RS("MarkFontName") = "Comic Sans MS" Then Response.Write ("selected")
		Response.Write ">Comic"
		Response.Write "                Sans MS</OPTION>"
		Response.Write "                <OPTION value=""Courier New"" "
		If RS("MarkFontName") = "Courier New" Then Response.Write ("selected")
		Response.Write ">Courier"
		Response.Write "                New</OPTION>"
		Response.Write "                <OPTION value=""Georgia"" "
		If RS("MarkFontName") = "Georgia" Then Response.Write ("selected")
		Response.Write ">Georgia</OPTION>"
		Response.Write "                <OPTION value=""Impact"" "
		If RS("MarkFontName") = "Impact" Then Response.Write ("selected")
		Response.Write ">Impact</OPTION>"
		Response.Write "                <OPTION value=""Tahoma"" "
		If RS("MarkFontName") = "Tahoma" Then Response.Write ("selected")
		Response.Write ">Tahoma</OPTION>"
		Response.Write "                <OPTION value=""Times New Roman"" "
		If RS("MarkFontName") = "Times New Roman" Then Response.Write ("selected")
		Response.Write ">Times"
		Response.Write "                New Roman</OPTION>"
		Response.Write "                <OPTION value=""Trebuchet MS"" "
		If RS("MarkFontName") = "Trebuchet MS" Then Response.Write ("selected")
		Response.Write ">Trebuchet"
		Response.Write "                MS</OPTION>"
		Response.Write "                <OPTION value=""Script MT Bold"" "
		If RS("MarkFontName") = "Script MT Bold" Then Response.Write ("selected")
		Response.Write ">Script"
		Response.Write "                MT Bold</OPTION>"
		Response.Write "                <OPTION value=""Stencil"" "
		If RS("MarkFontName") = "Stencil" Then Response.Write ("selected")
		Response.Write ">Stencil</OPTION>"
		Response.Write "                <OPTION value=""Verdana"" "
		If RS("MarkFontName") = "Verdana" Then Response.Write ("selected")
		Response.Write ">Verdana</OPTION>"
		Response.Write "                <OPTION value=""Lucida Console"" "
		If RS("MarkFontName") = "Lucida Console" Then Response.Write ("selected")
		Response.Write ">Lucida"
		Response.Write "                Console</OPTION>"
		Response.Write "            </SELECT> </td>"
		Response.Write "          </tr>"
		Response.Write "          <tr bgcolor=""#FFFFFF"">"
		Response.Write "            <td height=""26"" align=""right"" class=""table"">字体是否粗体:</td>"
		Response.Write "            <td class=""table""> <SELECT name=""MarkFontBond"" id=""MarkFontBond"">"
		Response.Write "                <OPTION value=0 "
		If RS("MarkFontBond") = "0" Then Response.Write ("selected")
		Response.Write ">否</OPTION>"
		Response.Write "                <OPTION value=1 "
		If RS("MarkFontBond") = "1" Then Response.Write ("selected")
		Response.Write ">是</OPTION>"
		Response.Write "            </SELECT> </td>"
		Response.Write "          </tr>"
		Response.Write "          </table>"
		Response.Write "          </td>"
		Response.Write "          </tr>"
		Response.Write "          <tr>"
		Response.Write "           <td colspan=""2"">"
		Response.Write "           <table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""1"" bgcolor=""#efefef"" id=""picarea"">"
		Response.Write "          <tr bgcolor=""#FFFFFF"">"
		Response.Write "            <td width=""27%"" height=""26"" align=""right"" class=""table"">LOGO图片:<br> </td>"
		Response.Write "            <td width=""70%"" class=""table""> <INPUT TYPE=""text"" NAME=""MarkPicture"" size=40 value=""" & RS("MarkPicture") & """>"
		Response.Write "           <font color=""#FF0000"">列如:/images/logo.gif</font></td>"
		Response.Write "          </tr>"
		Response.Write "          <tr bgcolor=""#FFFFFF"">"
		Response.Write "            <td height=""26"" align=""right"" class=""table"">LOGO图片透明度:</td>"
		Response.Write "            <td class=""table""> <INPUT TYPE=""text"" NAME=""MarkOpacity"" size=10 value="""
		If Left(RS("MarkOpacity"), 1) = "." Then Response.Write ("0" & RS("MarkOpacity")) Else Response.Write (RS("MarkOpacity"))
		Response.Write """>"
		Response.Write "            如50%请填写0.5 </td>"
		Response.Write "          </tr>"
		Response.Write "          <tr bgcolor=""#FFFFFF"">"
		Response.Write "            <td height=""26"" align=""right"" class=""table"">图片去除底色:</td>"
		Response.Write "            <td class=""table""> <INPUT TYPE=""text"" NAME=""MarkTranspColor"" ID=""MarkTranspColor"" maxlength = 7 size = 7 value=""" & RS("MarkTranspColor") & """ >"
		Response.Write "              <input type=""button"" name=""button1"" value=""选择颜色..."" onClick=""OpenThenSetValue('inc/SelectColor.asp',230,190,window,document.form.MarkTranspColor);document.form.MarkTranspColor.style.color=document.form.MarkTranspColor.value;"">"
		Response.Write "            保留为空则水印图片不去除底色。 </td>"
		Response.Write "          </tr>"
		Response.Write "          <tr bgcolor=""#FFFFFF"">"
		Response.Write "            <td height=""26"" align=""right"" class=""table"">图片坐标位置:<br> </td>"
		Response.Write "            <td class=""table""> 　X："
		Response.Write "              <INPUT TYPE=""text"" NAME=""MarkWidth"" size=10 value=""" & RS("MarkWidth") & """>"
		Response.Write "              象素<br>"
		Response.Write "Y:"
		Response.Write "              <INPUT TYPE=""text"" NAME=""MarkHeight"" size=10 value=""" & RS("MarkHeight") & """>"
		Response.Write "            象素  </td>"
		Response.Write "          </tr>"
		Response.Write "          </table>"
		Response.Write "          </td>"
		Response.Write "          </tr>"
				  
		Response.Write "      </table></td>"
		Response.Write "    </tr>"
		Response.Write "</body>"
		
		Response.Write "<script language=""javascript"">"
		Response.Write "if (document.all.MarkFontColor.value!='')"
		Response.Write " document.all.MarkFontColor.style.color=document.all.MarkFontColor.value;"
		Response.Write "if (document.all.MarkTranspColor.value!='')"
		Response.Write " document.all.MarkTranspColor.style.color=document.all.MarkTranspColor.value;"
		Response.Write "function SetTypeArea(TypeID)"
		Response.Write "{"
		Response.Write " if (TypeID==1)"
		Response.Write "  {"
		Response.Write "   document.all.textarea.style.display='';"
		Response.Write "   document.all.picarea.style.display='none';"
		Response.Write "  }"
		Response.Write " else"
		Response.Write "  {"
		Response.Write "   document.all.textarea.style.display='none';"
		Response.Write "   document.all.picarea.style.display='';"
		Response.Write "  }"
		
		Response.Write "}"
		Response.Write "function ShowInfo(ComponentID)"
		Response.Write "{"
		Response.Write "    if(ComponentID == 0)"
		Response.Write "    {"
		Response.Write "        document.all.WaterMarkSetting.style.display = ""none"";"
		Response.Write "    }"
		Response.Write "    else"
		Response.Write "    {"
		Response.Write "        document.all.WaterMarkSetting.style.display = """";"
		Response.Write "    }"
		Response.Write "}"
		Response.Write "function ShowThumbInfo(ThumbComponentID)"
		Response.Write "{"
		Response.Write "    if(ThumbComponentID == 0)"
		Response.Write "    {"
		Response.Write "        document.all.ThumbSetting.style.display = ""none"";"
		Response.Write "    }"
		Response.Write "    else"
		Response.Write "    {"
		Response.Write "        document.all.ThumbSetting.style.display = """";"
		Response.Write "    }"
		Response.Write "}"
		Response.Write "function ShowThumbSetting(ThumbSettingid)"
		Response.Write "{"
		Response.Write "    if(ThumbSettingid == 0)"
		Response.Write "    {"
		Response.Write "        document.all.ThumbSetting1.style.display = ""none"";"
		 Response.Write "       document.all.ThumbSetting0.style.display = """";"
		 Response.Write "   }"
		 Response.Write "   else"
		Response.Write "    {"
		Response.Write "        document.all.ThumbSetting1.style.display = """";"
		Response.Write "        document.all.ThumbSetting0.style.display = ""none"";"
		Response.Write "    }"
		Response.Write "}"
		Response.Write "function CheckForm()"
		Response.Write "{"
		Response.Write "document.form.submit();"
		Response.Write "}"
		
		Response.Write "</script>"
		
	%>	
	</td>
  </tr>
  <tr>
    <td colspan="3" class="title2">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="3" class="title"><div align="center" style="font-weight: bold">目录设置</div></td>
	
    </tr>
  <tr>
    <td colspan="3" class="title2">&nbsp;此设置应用于远程保存文件目录地址。<span class="table"></span></td>
    </tr>
  <tr>
    <td class="table">文章采集目录</td>
	<%
	Dim rs1
	Set rs1 = ConnItem.execute("Select top 1 Dir,MaxFileSize,FileExtName,Timeout from SK_Cj where ID=1")
	%>
    <td colspan="2" class="table"><input name="ArticleDir" type="text" class="lostfocus" id="ArticleDir" onfocus='this.className="getfocus";this.gf="1"' onblur='this.className="lostfocus";this.gf="0"' onmouseover='this.className="getfocus"' onmouseout='if (this.gf=="0") this.className="lostfocus"' value="<% =rs1("Dir") %>" gf="0">      
      <font class="alert">&nbsp;后面必须带"/"符号</font></td>
  </tr>
  <div style="display:none">
  <tr style="display:none">
    <td class="table">&nbsp;图片采集目录</td>
	<%Set rs1 = ConnItem.execute("Select top 1 Dir,MaxFileSize,FileExtName,Timeout from SK_Cj where ID=2")%>
    <td colspan="2" class="table"><input name="photoDir" type="text" class="lostfocus" id="PicDir" onfocus='this.className="getfocus";this.gf="1"' onblur='this.className="lostfocus";this.gf="0"' onmouseover='this.className="getfocus"' onmouseout='if (this.gf=="0") this.className="lostfocus"' value="<% =rs1("dir") %>" gf="0">
      <font class="alert">&nbsp;后面必须带"/"符号</font></td>
  </tr>
  <tr style="display:none">
    <td class="table">软件采集目录</td>
	<%Set rs1 = ConnItem.execute("Select top 1 Dir,MaxFileSize,FileExtName,Timeout from SK_Cj where ID=3")%>
    <td colspan="2" class="table"><input name="DownDir" type="text" class="lostfocus" id="alldir" onfocus='this.className=&quot;getfocus&quot;;this.gf=&quot;1&quot;' onblur='this.className=&quot;lostfocus&quot;;this.gf=&quot;0&quot;' onmouseover='this.className=&quot;getfocus&quot;' onmouseout='if (this.gf==&quot;0&quot;) this.className=&quot;lostfocus&quot;' value="<% =rs1("Dir") %>" gf="0" />
	<%Rs1.close : Set Rs1=Nothing%>
      <font class="alert">&nbsp;后面必须带&quot;/&quot;符号</font></td>
  </tr>
  </div>
  <tr>
    <td height="50" colspan="3" class="title3"><div align="center">
      <input type="submit" name="Submit" value="提交">
    </div></td>
    </tr>
</form>
</table>
</body>
<%
rs.close
set rs=nothing
end sub
	Public Function ExpiredStr(I)
		   Dim ComponentName(3)
			ComponentName(0) = "Persits.Jpeg"
			ComponentName(1) = "wsImage.Resize"
			ComponentName(2) = "SoftArtisans.ImageGen"
			ComponentName(3) = "CreatePreviewImage.cGvbox"
			If IsObjInstalled(ComponentName(I)) Then
				If IsExpired(ComponentName(I)) Then
					ExpiredStr = "，但已过期"
				Else
					ExpiredStr = ""
				End If
			  ExpiredStr = " √支持" & ExpiredStr
			Else
			  ExpiredStr = "×不支持"
			End If
	End Function
		Public Function IsObjInstalled(strClassString)
		On Error Resume Next
		IsObjInstalled = False
		Err = 0
		Dim xTestObj:Set xTestObj = Server.CreateObject(strClassString)
		If 0 = Err Then IsObjInstalled = True
		Set xTestObj = Nothing
		Err = 0
	End Function
	Public Function IsExpired(strClassString)
		On Error Resume Next
		IsExpired = True
		Err = 0
		Dim xTestObj:Set xTestObj = Server.CreateObject(strClassString)
	
		If 0 = Err Then
			Select Case strClassString
				Case "Persits.Jpeg"
					If xTestObjResponse.Expires > Now Then
						IsExpired = False
					End If
				Case "wsImage.Resize"
					If InStr(xTestObj.errorinfo, "已经过期") = 0 Then
						IsExpired = False
					End If
				Case "SoftArtisans.ImageGen"
					xTestObj.CreateImage 500, 500, RGB(255, 255, 255)
					If Err = 0 Then
						IsExpired = False
					End If
			End Select
		End If
		Set xTestObj = Nothing
		Err = 0
	End Function
		'取得Request.Querystring 或 Request.Form 的值
	Public Function G(Str)
	 G = Replace(Replace(Request(Str), "'", ""), """", "")
	End Function
%>

