<!--#include file="conn.asp" -->
<%
IF session("upadmin") ="guanli" Then

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="UTF-8">
<head>
<title>图片上传</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="imagetoolbar" content="no"/>
<link rel="stylesheet" rev="stylesheet" href="../../common/control.css" type="text/css" media="all" />
<link type="text/css" href="css/admin.css" rel="stylesheet" />
<script type="text/javascript" src="js/adminsys.js"></script>
<script type="text/javascript" src="js/bingo.js"></script>
<script type="text/javascript">
<!--
var AllowImgFileSize="<%=picsize%>";		//允许上传图片文件的大小 0为无限制  单位：KB 
var AllowImgWidth="<%=picwidth%>";			//允许上传的图片的宽度  0为无限制　单位：px(像素)
var AllowImgHeight="<%=picheight%>";		//允许上传的图片的高度  0为无限制　单位：px(像素)
var abc="<%=img_module%>";
var uplo="<%=up_module%>";

var patn = /\.jpg$|\.bmp$|\.png$|\.jpeg$|\.gif$/i;
function load(Obj,inid,dx){
	if(patn.test(Obj.value)){
		var tempImg=new Image();
		//tempImg.onerror=function(){alert('目标类型错误或路径不存在！');Obj.outerHTML=Obj.outerHTML;};
		tempImg.onreadystatechange=function(){
			if(dx==0){
				if(AllowImgWidth!=0&&AllowImgWidth<this.width ||AllowImgHeight!=0&&AllowImgHeight<this.height ){	
					alert('缩略图的大小不要超过<% = picwidth %>X<% = picheight %>！');
					Obj.outerHTML=Obj.outerHTML;
				}
	 		}else{
				if(AllowImgFileSize!=0&&AllowImgFileSize*1024<this.fileSize){
					alert('单个图片大小不要超过(<% = picsize %>)K！');
					Obj.outerHTML=Obj.outerHTML;
				}
			}
		}
		tempImg.src=Obj.value;
	}else{
		alert("您选择的似乎不是图像文件。");
		$(inid).innerHTML=$(inid).innerHTML;
	}
}



var hhid,hpid
function setid(){
	if(uplo=="0") {
		$("upid").innerHTML="<div class='tipxulog'><span class='red'>上传已经关闭，不能上传</span></div>"
	return false;
       }
	
	if(!$F("picname")){
		$("namespan").innerHTML="请输入图片名称！";
		return false;
	}else{
		$("namespan").innerHTML="";
		}

	tuname=$("picname").value;
	var n = tuname.match(/\d+$/g);
	n == null ?shuxu='' : shuxu= n.join('、');
	tempname=tuname.substr(0 ,(tuname.length-shuxu.length));
	twe=shuxu.length;
	str='';
	if(!$("upcount").value||$("upcount").value==0){$("upcount").value=1;}
	if($("upcount").value><% = MaxNum %>){
		alert("您最多只能同时上传 <% = MaxNum %> 个图片!");
		$("upcount").value = <% = MaxNum %>;
		setid();
	}else{
		for(i=1;i<=$("upcount").value;i++){
			str+="<li><div class='tipxu1'>"+geshu(i,2)+"</div>"
			str+="<div class='timxu1'><input type='text' size='11' class='anc' name='tuname"+i+"' id='tuname"+i+"' value='"+tempname+geshu(Number(shuxu)+Number(i),twe)+"' /></div>";
			if(abc=="0"){
				str+="<div class='tiuxu'>";
				str+="<div>缩略图:<span id='ystu"+(i*2-1)+"'><input type='file' size='42' name='file"+(i*2-1)+"' id='file"+(i*2-1)+"' class='anb' onchange=\"load(this,'ystu"+(i*2-1)+"',0)\" /></span></div>";
			}else{
				str+="<div class='tiuxul'>";
			}
			str+="<div>原始图:<span id='sltu"+(i*2)+"'><input type='file' size='30' name='file"+(i*2)+"' id='file"+(i*2)+"' class='anb' onchange=\"load(this,'sltu"+(i*2)+"',1)\"></span></div>";
			str+="</div><div class='tizxu'><textarea  name='suomin"+i+"' id='suomin"+i+"'class='titext'></textarea></div>";
			str+="</li>";
			}
		strs="<li><div><input type='submit' class='an1' name='button2' id='button2' value='上 传'/>";
		$("upid").innerHTML="<ul>"+str+strs+"</ul>";
		$("myform").target="upframe";
		$("uptu").innerHTML="<iframe name='upframe' id='upframe' width='0' height='0' src='about:blank'></iframe>";
	}
}
function showIndex(){
	var iText=getElementsByClassName("anc","input");
		for(var i=0;i<iText.length;i++){
			if(iText[i].value==''){
				window.alert('第'+eval(i+1)+'个图片名称不能为空！');
				iText[i].style.border = "1px solid #FF0000";
				iText[i].focus();
				return false;
			}else{
                iText[i].style.border = "1px solid #84B0C7";
			}
        }
	var oText=getElementsByClassName("anb","input");
		for(var i=0;i<oText.length;i++){
			if(oText[i].value==''){
				window.alert('第'+eval(i+1)+'个图片地址不能为空！');
                oText[i].style.border = "1px solid #FF0000";
                oText[i].focus();
                return false;
			}else{
                oText[i].style.border = "1px solid #84B0C7";
			}
        }
         $("picname").value=iText[iText.length-1].value;
         $("button2").disabled=true;
}
function killErrors() {    
return true;    
}    
window.onerror = killErrors;
//-->
</script>
</head>
<body class="ContentBody">
<div class="MainDiv">
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
  <tr><th class="CTitle">相册 -  图片上传 </tr>
<tr><td class="CPanel"><div class="SubMenu"><a href="config.asp?menu=about">关于相册</a> | <a href="config.asp">分类管理</a> | <a href="addpicupload.asp">图片上传</a> | <a href="addconfig.asp">其他设置</a> | <a href="../../ConContent.asp?Fmenu=Skins&Smenu=Plugins">返回插件管理</a></div></td></tr>
<tr><td>
<div class="yun" id="uptu"></div>
</td></tr>
<tr><td>

<form id='myform' name='myform' method='post'action="savepic.asp?act=uploadpic" enctype="multipart/form-data" onsubmit="return showIndex();">
<table border="0" cellpadding="2" cellspacing="1" class="TablePanel" style="width:770px">
<tr align="center">
<td colspan="2" nowrap="nowrap" class="TDHead" >选择相册</td>
<td width="374" nowrap="nowrap" class="TDHead" >图片名称</td>
<td width="171" nowrap="nowrap" class="TDHead" >上传图片数</td>
</tr>
	   
<tr align="center">
<td colspan="2">
<select name="sel1" id="sel1" onchange="changeselx('sel1',1,3,0)" >
<option value="0">选择大类</option>
<%set rs=server.createobject("adodb.recordset")
rs.Open "select * from blog_album_class ORDER BY album_classID desc",conn,1,1
do while not rs.eof 
Response.write "<option value="""&rs("album_classID")&""">"&rs("album_className")&"</option>"
rs.movenext
loop
rs.close
set rs=nothing%>
</select>
<span id="spabtex1">请选择</span>
</td>
<td><input id='picname'class="text" type='text' value='' name='picname' style="width:200px"/><span id="namespan"></span></td>
<td >
<input type="text" name="upcount" id="upcount"value="1" class="text" onkeyup="this.value=this.value.replace(/\D/g,'')" style="width:50px"/>
<input type="button" name="Button" class='an1' onclick="setid();" value='设定' />
</td>
</tr>

<tr align="center">
<td width="28" class="TDHead" >序</td>
<td width="182" class="TDHead" >图片名称</td>
<td class="TDHead">图片地址</td>
<td class="TDHead">说明</td>
</tr>

<tr align="center" bgcolor="#D5DAE0">
<td colspan="4">
<div class="tixu" id="upid">
</div>
</td>
</tr>	
</table>
</form>

</td></tr>
</table>
 

</div>
</body>
</html>
<% Else
 Response.Redirect("../../default.asp")
End IF
%>