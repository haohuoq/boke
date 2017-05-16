<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.1//EN" "http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="UTF-8">
<!--#include file="sysfile/config.asp"-->
<!--#include file="sysfile/Function.asp"-->
<%
IF session("upadmin") ="guanli" Then

%>
<%
if request("method")=1 then
	if IsObjInstalled("Scripting.FileSystemobject")=false then
		response.write "<script language=javascript>alert('你的服务器不支持;FSO(Scripting.FileSystemObject)! 不能够修改');location.href='addconfig.asp';</script>"
		response.end
	end if
	logodb=Server.MapPath(""&trim(request("db"))&"")
	if ReportFileStatus(logodb) = false Then
		response.write "<script language=javascript>alert('数据库路径错误！请重新选择');location.href='addconfig.asp';</script>"	
		response.end
	end if
	if trim(request("up_module"))<>0 then
	
		zujian="Adodb.Stream"
	
		if IsObjInstalled(zujian)=false then
		
		response.write "<script language=javascript>alert('no"&zujian&"');location.href='addconfig.asp';</script>"	
		response.end
		end if
	end if
	if trim(request("img_module"))<>0 then
		
			tpzj="Persits.Jpeg"
		
	
	
		if IsObjInstalled(tpzj)=false Then
			response.write "<script language=javascript>alert('no ;"&tpzj&"');location.href='addconfig.asp';</script>"	
			response.end
		elseif trim(request("suiyin"))=2 then 
			logopic=Server.MapPath(""&request("suiimgpath")&"")
			if ReportFileStatus(logopic) = false Then
				response.write "<script language=javascript>alert('你使用图片水印，但水印图片不存在！请重新选择');location.href='addconfig.asp';</script>"	
				response.end
            end if
		end if
	end if


		set fso=Server.CreateObject("Scripting.FileSystemObject")
		set rs=fso.CreateTextFile(Server.MapPath("sysfile/config.asp"),true)
		rs.write "<%"                                                                	   & vbsrlf & vbsrlf
		rs.write " db=" & chr(34) & trim(request("db")) & chr(34) &"		                " & vbcrlf
		rs.write " up_module=" & chr(34) & trim(request("up_module")) & chr(34) &"		    " & vbcrlf
		rs.write " MaxNum=" & chr(34) & trim(request("MaxNum")) & chr(34) &"		    	" & vbcrlf
		rs.write " picsize=" & chr(34) & trim(request("picsize")) & chr(34) &"			    " & vbcrlf
		rs.write " img_module=" & chr(34) & trim(request("img_module")) & chr(34) &"		" & vbcrlf
		rs.write " picerr=" & chr(34) & trim(request("picerr")) & chr(34) & "				" & vbcrlf
		rs.write " pics=" & chr(34) & trim(request("picerr")) & chr(34) & "					" & vbcrlf
		rs.write " picwidth=" & chr(34) & trim(request("picwidth")) & chr(34) & "			" & vbcrlf
		rs.write " picheight="& chr(34) & trim(request("picheight")) & chr(34) &"			" & vbcrlf
		rs.write " small_gz="& chr(34) & trim(request.form("small_gz")) & chr(34) &"		" & vbcrlf
		rs.write " suiyin="& chr(34) & trim(request("suiyin")) & chr(34) &"			        " & vbcrlf
		rs.write " suiwz="& chr(34) & trim(request("suiwz"))& chr(34) & "			        " & vbcrlf
		rs.write " suiwzxx="& chr(34) & trim(request("suiwzxx")) & chr(34)& "			    " & vbcrlf
		rs.write " suiwzsize="& chr(34) & trim(request("suiwzsize"))& chr(34) & "			" & vbcrlf
		rs.write " suiimgys="& chr(34) & trim(request("suiimgys")) & chr(34)& "			    " & vbcrlf
		rs.write " suichu="& chr(34) & trim(request("suichu")) & chr(34)& "			        " & vbcrlf
		rs.write " suiimgpath="& chr(34) & trim(request("suiimgpath"))& chr(34) & "			" & vbcrlf
		rs.write " suiimgtm="& chr(34) & trim(request("suiimgtm"))& chr(34) & "			    " & vbcrlf
		rs.write " suiimgysdel="& chr(34) & trim(request("suiimgysdel"))& chr(34) & "		" & vbcrlf
    	rs.write "%" & ">"
	rs.close
	set rs=nothing
	set fso=nothing
	
	response.write "<script language=javascript>alert('OK!');location.href='addconfig.asp';</script>"
	response.end
end if
%>



<head>
<title>设置</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="imagetoolbar" content="no"/>
<link rel="stylesheet" rev="stylesheet" href="../../common/control.css" type="text/css" media="all" />
<link type="text/css" href="css/admin.css" rel="stylesheet" />
<script type="text/javascript">
<!--
window.onload=function(){
$("colourPalette").style.visibility="hidden";
intocolor();

}
function changeup(vid,spname){
	var objspans = $(spname).getElementsByTagName ("SPAN");
		for (i=0;i<objspans.length;i++){
			if(i==vid){
				objspans[i].style.display = "block";
			}else{
				objspans[i].style.display = "none";
			}
		}
}
function checkweb(theform) {
	var n1
	n1=parseFloat($('MaxNum').value)
	if (n1<"2"||n1>"20"){
		alert("上传限制个数不能小于2个，也不要大于20个");
		$('MaxNum').focus();
		return false; 
	} 
	if ($('MaxNum').value==""){
		alert("上传限制个数不能为空");
		$('MaxNum').focus();
		return false; 
	}
	var pattern = /^#[0-9a-fA-F]{6}$/
	var obj1 = eval("$('suiimgys').value");
	var obj2 = eval("$('suiimgysdel').value");
	if ($('suiimgys').value!=""){
		if(obj1.match(pattern)==null){
			alert("水印图片去除底色错误，请用调色板选择");
			$('suiimgys').focus();
			return false; 
		}
	}
	if ($('suiimgysdel').value!=""){
		if(obj2.match(pattern)==null){
			alert("水印图片去除底色错误，请用调色板选择");
			$('suiimgysdel').focus();
			return false; 
		}
	}
	if ($('dattu').value==""){
		alert("存放大图图片的文件夹不能为空");
		$('dattu').focus();
		return false; 
	}
	if ($('xiaottu').value==""){
		alert("存放小图图片的文件夹不能为空");
		$('xiaottu').focus();
		return false; 
	}
	if ($('dattu').value==$('xiaottu').value){
		alert("小图不能和大图存放在一起\n会造成自动缩略图错误");
		$('xiaottu').focus();
		return false; 
	}
}
function rdl_color(){
	var pattern = /^#[0-9a-fA-F]{6}$/
	var obj1 = eval("$('suiimgys').value");
	if ($('suiimgys').value!=""){
		if(obj1.match(pattern)!=null){
			rdl_change();
			$('syuys').style.backgroundColor=$('suiimgys').value;
		}
	}
}
function rdl_change(){
	shText=$('suiwzxx').value;
	$('ylText').innerHTML=shText;
	if ($('suichu').value==1){
		$('ylText').style.fontWeight="bold";
	}else{
		$('ylText').style.fontWeight="normal";
	}
	sValue=$('suisizename').options[$('suisizename').selectedIndex].value;
	$('ylText').style.color=$('suiimgys').value;
	$('ylText').style.fontSize=$('suiwzsize').value+"px";
	$('ylText').style.fontFamily=sValue;
}
//调色板
var ColorHex=new Array('00','33','66','99','CC','FF')
var SpColorHex=new Array('FF0000','00FF00','0000FF','FFFF00','00FFFF','FF00FF')
var current=null
function intocolor(){
	var colorTable=''
	for (i=0;i<2;i++){
		for (j=0;j<6;j++){
			colorTable=colorTable+'<tr height=12>';
			colorTable=colorTable+'<td width=11 id="#000000" style="background-color:#000000">';    
			if (i==0){
				colorTable=colorTable+'<td width=11 id="#'+ColorHex[j]+ColorHex[j]+ColorHex[j]+'" style="background-color:#'+ColorHex[j]+ColorHex[j]+ColorHex[j]+'">';
			}else{
				colorTable=colorTable+'<td width=11 id="#'+SpColorHex[j]+'" style="background-color:#'+SpColorHex[j]+'">';
			} 
			colorTable=colorTable+'<td width=11 id="#000000" style="background-color:#000000">';
			for (k=0;k<3;k++){
				for (l=0;l<6;l++){
					colorTable=colorTable+'<td width=11 id="#'+ColorHex[k+i*3]+ColorHex[l]+ColorHex[j]+'" style="background-color:#'+ColorHex[k+i*3]+ColorHex[l]+ColorHex[j]+'">';
				}
			}
		}
	}
colorTable='<table width=253 border="0" cellspacing="0" cellpadding="0" style="border:1px #000000 solid;border-bottom:none;border-collapse: collapse">'
           +'<tr height=30><td colspan=21 bgcolor=#cccccc>'
           +'<table cellpadding="0" cellspacing="1" border="0" style="border-collapse: collapse;text-align:left;">'
           +'<tr><td width="3"><td><input type="text" name="DisColor" id="DisColor" size="6" disabled style="border:solid 1px #000000;background-color:#fff"></td>'
           +'<td width="3"><td><input type="text" name="HexColor" id="HexColor" size="7" style="border:inset 1px;font-family:Arial;" value="#000000"></td><td width="120"></td></tr></table></td></table>'
           +'<table border="1" cellspacing="0" cellpadding="0" style="border-collapse:collapse;cursor:pointer;background-color:#000000" onmouseover="doOver(event)" onmouseout="doOut(event)" onclick="doclick(event)">'
           +colorTable+'</table>';          
document.getElementById('colourPalette').innerHTML=colorTable;
}
function doOver(evt){
	var colorsrc = evt.srcElement ? evt.srcElement : evt.target;
	if ((colorsrc.tagName=="TD") && (current!=colorsrc)){
        if (current!=null){current.style.backgroundColor = current._background;}     
        colorsrc._background = colorsrc.style.backgroundColor;
        document.getElementById('DisColor').style.backgroundColor = colorsrc.style.backgroundColor;
        document.getElementById('HexColor').value = colorsrc.id;
        current = colorsrc;
      }
}
function doOut(evt){
    if (current!=null){
    	current.style.backgroundColor = current._background;
    }
}
function Getcolor(img_val,input_val){
	var obj = document.getElementById("colourPalette");
	obj.style.display="block";
	ColorImg = img_val;
	ColorValue = document.getElementById(input_val);
	if (obj){
		obj.style.left = getOffsetLeft(ColorImg) + "px";
		obj.style.top = (getOffsetTop(ColorImg) + ColorImg.offsetHeight) + "px";
	if (obj.style.visibility=="hidden"){
		obj.style.visibility="visible";
	}else {
		obj.style.visibility="hidden";
	}
	}
}
function getOffsetTop(elm) {
	var mOffsetTop = elm.offsetTop;
	var mOffsetParent = elm.offsetParent;
	while(mOffsetParent){
		mOffsetTop += mOffsetParent.offsetTop;
		mOffsetParent = mOffsetParent.offsetParent;
	}
	return mOffsetTop;
}

function getOffsetLeft(elm) {
	var mOffsetLeft = elm.offsetLeft;
	var mOffsetParent = elm.offsetParent;
	while(mOffsetParent) {
		mOffsetLeft += mOffsetParent.offsetLeft;
		mOffsetParent = mOffsetParent.offsetParent;
	}
	return mOffsetLeft;
}
function doclick(evt){
	var colorsrc = evt.srcElement ? evt.srcElement : evt.target;
	if (ColorValue){ColorValue.value = colorsrc.id;}
	if (ColorImg){ColorImg.style.backgroundColor = colorsrc.id;}
	document.getElementById("colourPalette").style.visibility="hidden";
	rdl_change();
}
//-->
</script>
</head>
<body class="ContentBody">
<div class="MainDiv">
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">
  <tr><th class="CTitle">相册 -  上传设置 </tr>
<tr><td class="CPanel"><div class="SubMenu"><a href="config.asp?menu=about">关于相册</a> | <a href="../../ConContent.asp?Fmenu=Skins&Smenu=PluginsOptions&Plugins=album">基本设置</a> |  <a href="config.asp">分类管理</a> | <a href="addpicupload.asp">图片上传</a> | <a href="addconfig.asp">其他设置</a> | <a href="../../ConContent.asp?Fmenu=Skins&Smenu=Plugins">返回插件管理</a></div></td></tr>
<tr><td>
<div id="colourPalette" style=" display:none; visibility:hidden; position: absolute;"></div>
</td></tr>
<tr><td>


<table border="0" cellpadding="2" cellspacing="1" class="TablePanel" style="width:770px">
<tr align="center">
<td colspan="3" nowrap="nowrap" class="TDHead" >相册上传设置</td>
</tr>


<tr align="center" >
<td colspan="3">

<form action method="post" name="myform" id="myform" onsubmit="return checkweb(myform)">
				<div class='tjiao'>
					<ul>
					<li><div class="adxinrz">数据库名：</div>
						<div class="adxinl"><input type="text" name="db"  size="20" maxlength="100" value="<%=db%>"/></div>
						<div class="adxinlz">相对网站根目录路径</div>
					</li>
					<li><div class="adxinrz">上传组件：</div>
						<div class="adxinl">
						<select name="up_module" id="up_module" class="an" onchange="changeup(this.value,'spans')" size="1">
							<option value="0" <%If up_module = "0" then response.write "selected='selected'"%>>关闭</option>
							<option value="1" <%if up_module = "1" then response.write "selected='selected'"%>>无组件上传类</option>
						</select>
						</div>
						<div id="spans" class="adxinlz">
							<span id="w0" <%if up_module <> "0" then response.write "style='display:none'"%>><font color="red"><b>×</b></font>（关闭上传）</span>
							<span id="w1" <%if up_module <> "1" then response.write "style='display:none'"%>><%If IsObjInstalled("Adodb.Stream")=false Then%><font color="red"><b>×</b></font>（服务器不支持）<% else %><b>√</b>（服务器支持）<% end if %></span>							
						</div>
					</li>
					<li><div class="adxinrz">成批上传：</div>
						<div class="adxinl"><input type="text" id="MaxNum" name="MaxNum"  size="20" maxlength="100" value="<%=MaxNum%>" onkeyup="this.value=this.value.replace(/\D/g,'')"/></div>
						<div class="adxinlz">限制可以成批上传图片数</div>
					</li>
					<li><div class="adxinrz">大图上传限制：</div>
						<div class="adxinl"><input type="text" name="picsize" size="20" maxlength="100" value="<%=picsize%>" onkeyup="this.value=this.value.replace(/\D/g,'')"/></div>
						<div class="adxinlz">允许上传单个大图大小：单位 KB, 0 为无限制</div>
					</li>
					<li>
					  <div class="adxinrz">是否支持aspjpeg组件：</div>
						<div class="adxinl"><select name="img_module" id="img_module" class="an" onchange="changeup(this.value,'spana')">
											<option value="1" <%if img_module = "1" then response.write "selected"%>>AspJpeg图片组件</option>
                                            </select>
                        </div>
						<div class="adxinlz"  id="spana">
							
							<span id="wimg1" <%if img_module <> "1" then response.write "style='display:none'" end if%>><%If IsObjInstalled("Persits.Jpeg")=false Then%><font color="red"><b>×</b></font>（服务器不支持）<% else %><b>√</b>（服务器支持）<% end if %></span>
							
						</div>
					</li>
					<li><div class="adxinrz">默认的缩略图：</div>
						<div class="adxinl">
						<input type="text" name="picerr" id="picerr" size="20" maxlength="100" value="<%=picerr%>" onkeyup="this.value=this.value.replace(/\D/g,'')"/>
						</div>
						<div class="adxinlz"> 图片出错和默认缩略图的路径</div>
					</li>
					<li><div class="adxinrz">缩略图的大小：</div>
						<div class="adxinl">
						<input type="text" name="picwidth"  size="7" maxlength="100" value="<%=picwidth%>" onkeyup="this.value=this.value.replace(/\D/g,'')"/>
									X
						<input type="text" name="picheight"  size="7" maxlength="100" value="<%=picheight%>" onkeyup="this.value=this.value.replace(/\D/g,'')"/>
						</div>
						<div class="adxinlz"> 缩略图大小：宽 X 高,</div>
					</li>
					<li><div class="adxinrz">生成缩略图规则：</div>
						<div class="adxinl">
						<select name="small_gz" id="small_gz" class="an" onchange="changeup(this.value,'spansmall')">
							<option value="0"<%if small_gz = "0" then response.write "selected"%>>固定比例</option>
							<option value="1"<%if small_gz = "1" then response.write "selected"%>>等比例缩小</option>
							<option value="2"<%if small_gz = "2" then response.write "selected"%>>锁定高宽比</option>
						  </select></div>
						<div class="adxinlz" id="spansmall">
							<span id="small0"<%if small_gz <> "0" then response.write "style='display:none'" end if%>>不管图片的大小，都按上面的宽X高生成缩略图</span>
							<span id="small1"<%if small_gz <> "1" then response.write "style='display:none'" end if%>>按图像最长的边为基准与宽(上面)等比例生成</span>
							<span id="small2"<%if small_gz <> "2" then response.write "style='display:none'" end if%>>按图像比例生成，并修剪以适合上面的长宽比</span>
						</div>

					</li>
					<li><div class="adxinrz">添加水印设置：</div>
						<div class="adxinl"><select name="suiyin" id="suiyin" class="an" size="1">
							<option value="0" <%if suiyin = "0" then response.write "selected"%>>关闭水印效果</option>
							<option value="1" <%if suiyin = "1" then response.write "selected"%>>文字水印效果</option>
							<option value="2" <%if suiyin = "2" then response.write "selected"%>>图片水印效果</option>
						</select></div>
					</li>
					<li><div class="adxinrz">水印添加位置：</div>
						<div class="adxinl"><select name="suiwz" id="suiwz" class="an">
							<option value="1" <%if suiwz = "1" then response.write "selected"%>>左上</option>
							<option value="2" <%if suiwz = "2" then response.write "selected"%>>左下</option>
							<option value="3" <%if suiwz = "3" then response.write "selected"%>>居中</option>
							<option value="4" <%if suiwz = "4" then response.write "selected"%>>右上</option>
							<option value="5" <%if suiwz = "5" then response.write "selected"%>>右下</option>
						</select></div>
					</li>
					<li><div class="adxinrz">文字水印信息：</div>
						<div class="adxinl"><input type="text"  name="suiwzxx" id="suiwzxx" value="<%=suiwzxx%>" onkeyup="rdl_change()" /></div>
						<div class="adxinlz">
							<div style="position: absolute; width: 250px; height: 50px; z-index: 1; font-family:Batang" id="ylText" name="ylText"></div>
						</div>
					</li>
					<li><div class="adxinrz">水印字体大小：</div>
						<div class="adxinl"><input type="text" id="suiwzsize" name="suiwzsize" value="<%=suiwzsize%>" onkeyup="this.value=this.value.replace(/\D/g,'');rdl_change()"/>px</div>
						<div class="adxinlz"></div>
					</li>
					<li><div class="adxinrz">水印字体颜色：</div>
						<div class="adxinl"><input type="text" name="suiimgys" id="suiimgys"  value="<%=suiimgys%>" onkeyup="rdl_color()"/>
						<img alt="选取颜色!" id="syuys" src="images/rect.gif" width="18" height="17" style="cursor:pointer;background-Color:<%=suiimgys%>; margin-bottom: 3px;" title="选取颜色!" onclick="Getcolor(this,'suiimgys');"/></div>
						<div class="adxinlz"></div>
					</li>
					
					<li><div class="adxinrz">字体是否粗体：</div>
						<div class="adxinl"><select name="suichu" id="suichu" class="an" onchange="rdl_change()">
							<option value="0" <%if suichu = "0" then response.write "selected"%>>否</option>
							<option value="1" <%if suichu = "1" then response.write "selected"%>>是</option>
						</select></div>
					</li>
					<li><div class="adxinrz">图片水印路径：</div>
						<div class="adxinl"><input type="text" name="suiimgpath"  value="<%=suiimgpath%>"/></div>
						<div class="adxinlz">填写LOGO的图片相对路径</div>
					</li>
					<li><div class="adxinrz">图片水印透明度：</div>
						<div class="adxinl"><input type="text" name="suiimgtm"  value="<%=suiimgtm%>"/></div>
						<div class="adxinlz">如60%请填写0.6</div>
					</li>
					<li><div class="adxinrz">水印图片去除底色：</div>
						<div class="adxinl"><input type="text" name="suiimgysdel" id="suiimgysdel"  value="<%=suiimgysdel%>"/>
						<img alt="选取颜色!" src="images/rect.gif" width="18" height="17" style="cursor:pointer;background-Color:<%=suiimgysdel%>; margin-bottom: 3px" title="选取颜色!" onclick="Getcolor(this,'suiimgysdel');"/></div>
						<div class="adxinlz">
						保留为空则水印图片不去除底色</div>
					</li>
					
					<li><div style="text-align:center">
					<input type="hidden" name="method" value="1"/>
					<input type="submit" class="an1" value="修改"/>&nbsp;&nbsp;&nbsp;&nbsp;
					<input type="reset" class="an1" name="Submit2" value="重写"/></div>
					</li>
					</ul>
				</div>
		</form>


</td>
</tr>	
</table>

</td></tr>
</table>
 

</div>
</body>
</html>
<% Else
 Response.Redirect("../../default.asp")
End IF
%>