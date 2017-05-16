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
Dim albumConn,page,typeid
Set albumConn=Server.CreateObject("ADODB.Connection")
albumConn.ConnectionString="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath(albumDBPath)&""
albumConn.Open
typeid=clng(request.querystring("classid"))
page=clng(request.querystring("page"))
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="UTF-8">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="Content-Language" content="UTF-8" />
<link rel="stylesheet" rev="stylesheet" href="../../common/control.css" type="text/css" media="all" />
<link rel="stylesheet" rev="stylesheet" href="css/css.css" type="text/css" media="all" />
<title>相册管理-内容</title>
<script type="text/javascript" src="js/adminsys.js"></script>
<script type="text/javascript" src="js/bingo.js"></script>
<script type="text/javascript" src="js/pages.js"></script>
<script type="text/javascript">
<!--
var hhid,hpid
var atid="<%=typeid%>"
var ghpage="<%=page%>"
var IsShowPreNext=true
var IsShowFirstLast=true
var IsShowNumSplitBox=true
var dialog = new DialogClass();
window.onload=function(){
_gpage=ghpage?ghpage:1;
	if(atid!=0){hhid=1;
	AJAXCALL("js/saveedt.asp?act=sphoto&typeid="+atid+"&page="+_gpage,"ShSelect");
	hpid=atid;
	}
}

function gotoandplay(p){
	AJAXCALL("js/saveedt.asp?act=sphoto&typeid="+atid+"&page="+p,"ShSelect");
}
function showpedt(edv){
	$("sh"+edv).style.display = "none";
	$("eh"+edv).style.display = "block";
	$("edzin"+edv).select()
	var aa=getElementsByClassName("editinput1","input");
	for (var ii=1; ii<aa.length+1; ii++){
		if (ii!=edv){showmout(ii);}
	}
}
function showmout(edv){
	$("sh"+edv).style.display = "block";
	$("eh"+edv).style.display = "none";
	$("edzin"+edv).value=$("edzen"+edv).value;
}
function editpic(pxu,ttid){
 var edvin=$("edzin"+pxu).value;
 	if(!isAvailLen(edvin,1,14)){
		alert("名称只能7个汉字或14个字母！");
		return false;
		}
		if(edvin==$("edzen"+pxu).value){
		 	alert("你没有做任何操作"); 
		}else{
			AJAXCALL("js/saveedt.asp?act=ephoto&photoid="+ttid+"&etname="+$U(edvin),"Edpic","attach="+pxu);
		}
}
function Edpic(eddoc,id){
	if (eddoc){
		var edbar=eval(eddoc);
		edbar=edbar[0];
		_errzt=edbar.errzt;
		if(_errzt==0){
			alert(esbar.errtx)
		}else{
			$("edzen"+id).value=edbar.edname;
			$("edzin"+id).value=edbar.edname;
			showmout(id);
		}
	}else{alert("系统错误");}
}
function Sptpic(edv,spid){
	if ($("tj"+edv).className=="red"){putsb=false;
		$("tj"+edv).className="";
		$("tj"+edv).title="点击隐藏";
	}else{putsb=true;
		$("tj"+edv).className="red";
		$("tj"+edv).title="点击不隐藏";
	}
	AJAXCALL("js/saveedt.asp?act=putsb&photoid="+spid+"&hots="+putsb,"EdPut","attach="+edv);
}
function comtpic(edv,spid){
	if ($("com"+edv).className=="red"){putsb=false;
		$("com"+edv).className="";
		$("com"+edv).title="点击禁止评论";
	}else{putsb=true;
		$("com"+edv).className="red";
		$("com"+edv).title="点击取消禁止评论";
	}
	AJAXCALL("js/saveedt.asp?act=comsb&photoid="+spid+"&coms="+putsb,"EdPut","attach="+edv);
}
function EdPut(spdoc){
	if (spdoc){alert(spdoc);}
	else{alert("系统错误");}}
	
function editphoto(xgpid,xgcid){
	$("edilai").className="showedit";
	AJAXCALL("js/saveedt.asp?act=edohuan&xgpid="+xgpid+"&xgcid="+xgcid,"EdhOthe");
}
function addshowout(){
hhid=1;$("edilai").innerHTML="";$("edilai").className="showedit1";
AJAXCALL("js/saveedt.asp?act=sphoto&typeid="+hpid+"&page="+_gpage,"ShSelect");
}
function EdhOthe(Eddoc){
	if (Eddoc){
		if(Eddoc=="1"){
			$("edilai").innerHTML="请登陆";
			return false;}
		$("edilai").innerHTML=Eddoc;
		$("myform").action="js/saveedt.asp?act=epicoh";
	}else{$("edilai").innerHTML="系统错误";
	}
}
function Delphoto(dpdoc){
	if (dpdoc){alert(dpdoc);
		AJAXCALL("js/saveedt.asp?act=sphoto&typeid="+atid+"&page="+_gpage,"ShSelect");
	}else{alert("系统错误");}
}
function DePhoto(dpid){
	if(confirm("确实要删除这张图片?")){
	AJAXCALL("savepic.asp?act=dphoto&id="+dpid,"Delphoto");
	}
}
function Eddcheck(postform){
	acv=$F("minname")
	if(document.myform.sel1.value==0) {
		$("spabtex1").innerHTML="请选择类别";
		return false;}
	if(!isAvailLen(acv,1,24)){
		$("addspn").innerHTML="名称只能12个汉字或24个字母！";
		$("minname").focus();return false;}
	if(!isAvailString(acv)){
		$("addspn").innerHTML="名称不要含有非法字符";
		$("minname").focus();return false;}
	if(isNull($("smotu").value)){
		$("smospan").innerHTML="请填写缩略图路径";
		$("smotu").focus();return false;}
	if(isNull($("bigtu").value)){
		$("bigspan").innerHTML="请填写缩略图路径";
		$("bigtu").focus();return false;}
$('buttonm1').disabled=true;
AJAXFORM(postform,"SavePhoto","attach=buttonm1");
}
function SavePhoto(asdoc,qid){
	if (asdoc){
	var asbar=eval(asdoc);
		asbar=asbar[0];
		_Dezt=asbar.Dezt;
		alert(asbar.Detx);
		$(qid).value=asbar.Detx;
}
else{alert("系统错误");}$('buttonm1').disabled=false;}
//-->
</script>
</head>
<body class="ContentBody">
<div class="MainDiv">
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
  <tr>
    <th class="CTitle">相册 - 图片管理
  </tr>
  <tr>
    <td class="CPanel">
	<div class="SubMenu"><a href="config.asp?menu=about">关于相册</a> | <a href="config.asp">分类管理</a> | <a href="addpicupload.asp">图片上传</a> | <a href="addconfig.asp">其他设置</a> | <a href="../../ConContent.asp?Fmenu=Skins&Smenu=Plugins">返回插件管理</a></div>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
<tr>
<td valign="top">
<div align="left" >
<!--class-->
<div id="edilai" class="showedit1"></div>

<DIV class=picshow>
<!--qqqqqq-->
<div class="searchpic" id="showminors"></div>
<!--qqqqqq-->

<div style="clear:both"> </div>
<div class="next_to">			
<div class="next_az" id="selid"></div>
<div class="next_zz" id="shuzi">&nbsp;</div>
<div class="next_ak" id="pages"></div>
</div>
<div style="clear:both"> </div>
</DIV>

<!--class end-->
</div></td></tr></table>
</td>
</tr>
</table>
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