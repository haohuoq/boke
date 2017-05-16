<!--#include file="../../const.asp" -->
<!--#include file="../../p_conn.asp" -->
<!--#include file="../../common/function.asp" -->
<!--#include file="../../common/library.asp" -->
<!--#include file="../../common/cache.asp" -->
<!--#include file="../../common/checkUser.asp" -->
<!--#include file="../../common/ModSet.asp" -->

<%
'=====================================
'  相册插件设置页面
'    更新时间: 2007-1-16
'=====================================

IF session(CookieName&"_System")=true and memName<>empty and stat_Admin=true Then
session("upadmin") ="guanli"

 
	'获取数据库路径
dim albumSet,albumDBPath
Set albumSet=New ModSet
albumSet.open("album")
albumDBPath=albumSet.getKeyValue("Database")
	'打开数据库
Dim albumConn
Set albumConn=Server.CreateObject("ADODB.Connection")
albumConn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath(albumDBPath)&""
albumConn.Open

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="UTF-8">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<meta http-equiv="Content-Language" content="UTF-8" />
<link rel="stylesheet" rev="stylesheet" href="../../common/control.css" type="text/css" media="all" />
<title>相册管理-内容</title>
</head>
<body class="ContentBody">
<div class="MainDiv">
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
  <tr>
    <th class="CTitle">相册 - <%
     if Request.QueryString("menu")="about" then
	 	Response.Write("关于相册")
	 else
	  	response.write("分类管理及批量上传")
	 end if
	 %>
  </tr>
  <tr>
    <td class="CPanel">
	<div class="SubMenu"><a href="config.asp?menu=about">关于相册</a> | <a href="../../ConContent.asp?Fmenu=Skins&Smenu=PluginsOptions&Plugins=album">基本设置</a> | <a href="config.asp">分类管理</a> | <a href="addpicupload.asp">图片上传</a> | <a href="addconfig.asp">其他设置</a> | <a href="../../ConContent.asp?Fmenu=Skins&Smenu=Plugins">返回插件管理</a></div>
 <% IF Request.QueryString("menu")="about" Then %>
 <table width="100%" border="0" cellpadding="0" cellspacing="0">
	 <tr>
		 <td valign="top"></td>
		 <td valign="top">
		    <div align="left" style="padding:5px;line-height:170%;clear:both;font-size:12px">
		     <strong>当前版本:</strong>  小锋相册3.0版<br/>
		     <strong>更新日期:</strong> <%=DateToStr("2008-12-24","mdy")%><br/>
		     <strong>插件制作:</strong> 小锋<br/>
		     <strong>电子邮件:</strong> wzlycz@163.com<br/>
		     <strong>个人主页:</strong> <a href="http://www.onewz.com" target="_blank" title="峥言锋语"> http://www.onewz.com</a> <br/>
		     <strong>版权声明:</strong> 免费使用，但必须遵循<a href="http://creativecommons.org/licenses/by-nc-sa/2.5/cn/deed.zh" target="_blank" title="知识共享(Creative Commons) 2.5">知识共享(Creative Commons) 2.5</a>。  <br/>
		    </div>
		 </td>
	 </tr>
	</table>
	
 <%  Else  %>
   
    <table width="100%" border="0" cellpadding="0" cellspacing="0">
	 <tr>
		 <td valign="top"></td>
		 <td valign="top">
		    <div align="left" style="padding:5px;line-height:170%;clear:both;font-size:12px">

<form action="albumaction.asp" method="post" style="margin:0px">
<input type="hidden" name="action" value="class"/>
      <div align="left" style="padding:5px;"><%getMsg%>
	   <table border="0" cellpadding="2" cellspacing="1" class="TablePanel">
        <tr align="center">
		  <td width="16" nowrap="nowrap" class="TDHead">&nbsp;</td>
          <td width="100" nowrap="nowrap" class="TDHead">相册类型名称</td>
          <td nowrap="nowrap" class="TDHead" style="width: 345px">分类图片地址</td>
          <td nowrap="nowrap" class="TDHead" style="width: 142px">相片管理</td>
	   </tr>
	   <%
	  dim album_class
      Set album_class=albumConn.execute("select * from blog_album_class order by album_classID asc")
	  do until album_class.eof 
	   %>
	   <tr align="center">
		  <td><input name="selectclassID" type="checkbox" value="<%=album_class("album_classID")%>"/></td>
          <td><input name="classID" type="hidden" value="<%=album_class("album_classID")%>"/><input name="className" type="text" size="20" class="text" value="<%=album_class("album_className")%>"/></td>
	      <td style="width: 345px">
			<input name="classPic" type="text" size="20" class="text" value="<%=album_class("album_classPic")%>" style="width: 332px"/></td>
	      <td style="width: 142px"><a href="editphoto.asp?classid=<%=album_class("album_classID")%>">	管理本类照片 </a></td>
	   </tr>
	   <%
	   album_class.movenext
	   loop
	   %>
	    <tr align="center" bgcolor="#D5DAE0">
        <td colspan="4" class="TDHead" align="left" style="border-top:1px solid #999"><a name="AddLink"></a><img src="../../images/add.gif" style="margin:0px 2px -3px 2px"/>添加类型</td>
       </tr>	
        <tr align="center">
          <td></td>
		  <td><input name="classID" type="hidden" value="-1"/><input name="className" type="text" size="20" class="text"/></td>
          <td style="width: 345px">
			<input name="classPic" type="text" size="20" class="text" value="plugins/album/photo/class.jpg" style="width: 325px"/></td>
          <td style="width: 142px">&nbsp;</td>
	   </tr>
	  </table>
  <div class="SubButton" style="text-align:left;">
    	 <select name="a_class">
			 <option value="SaveAll">保存所有类型</option>
			 <option value="DelSelect">删除所选类型</option>
		 </select>
	  <input type="submit" name="Submit" value="提交" class="button" style="margin-bottom:0px"/>   注意，慎用删除类型，删除可能导致首页随机调用出错
     </div>	</form>
</td></tr></table>
</div>




		    </div>
		 </td>
	 </tr>
	</table>
 
 <% End IF %>
</td></tr></table>

</div>
</body>
</html>
<%

Else
 Response.Redirect("../../default.asp")
End IF

'----------- 显示操作信息 ----------------------------
sub getMsg 
	 if session(CookieName&"_ShowMsg")=true then
	  response.write ("<div id=""msgInfo"" align=""center""><img src=""../../images/Control/aL.gif"" style=""margin-bottom:-11px;""/><span class=""alertTxt"">" & session(CookieName&"_MsgText") & "</span><img src=""../../images/Control/aR.gif"" style=""margin-bottom:-11px;""/></div>")
	  response.write ("<script>setTimeout('hiddenMsg()',3000);function hiddenMsg(){document.getElementById('msgInfo').style.display='none';}</script>")
	  session(CookieName&"_ShowMsg")=false
	  session(CookieName&"_MsgText")=""
	 end If
end sub

%>