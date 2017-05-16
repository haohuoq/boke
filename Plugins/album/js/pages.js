window.onload=function(){
externalLinks()
}
function xdaohang(_Count,_PaGe,_Pages){
	$("pages").innerHTML=DaoPage(_PaGe,_Pages);
	$("shuzi").innerHTML="<span class='red'>"+_Count+"</span>P <span class='red'>"+_PaGe+"</span>/<span style='color:blue;'>"+_Pages+"</span>Pages";
	$("selid").innerHTML=selectPage(_PaGe,_Pages);
}
function Postshow(nlei,ghp,soid){
_gp=ghp?ghp:1;
	AJAXCALL("js/bcastr.asp?act="+nlei+"&luid="+soid+"&page="+_gp,"showluhtm");
}
function showluhtm(doc){
if (doc){
	var mbar=eval(doc);
	mbar=mbar[0];
	$("mainpics").innerHTML=mbar.pagehtm;
	_gpage=mbar.page;
	_dcount=mbar.dcount;
	_gtotalPages=mbar.pages;
	$("daopages").innerHTML=DaoPage(_gpage,_gtotalPages);
	$("shuzi").innerHTML="<span class='red'>"+_dcount+"</span>P<span style='color:red;'>"+_gpage+"</span>/<span style='color:blue;'>"+_gtotalPages+"</span>Pages";
	$("selid").innerHTML=selectPage(_gpage,_gtotalPages);
	}
}
function DaoPage(mup,mupcoon){
	var pagectrl='';
	if (mup!=0){
		if(IsShowPreNext){
			if (mup==1){
				pagectrl="<li class=\"b\"><a class=\"off\"><img src=\"css/img/left.gif\"></a><a class=\"off\"><img src=\"css/img/left2.gif\"></a>";
			}else{
				pagectrl="<li class=\"b\"><a href=\"javascript:gotoandplay(1)\" ><img src=\"css/img/left1.gif\"></a>";
				pagectrl+="<a href=\"javascript:gotoandplay("+(mup-1)+")\" ><img src=\"css/img/left3.gif\"></a>";
			}
			}
		if(IsShowNumSplitBox){
			if(IsShowFirstLast){
			if (mupcoon>9&&mup > 5){
				pagectrl+="<li class=\"b\"><a href=\"javascript:gotoandplay(1)\" >1</a> "
				pagectrl+="<li class=\"c\"><a class=\"no\">...</a>"
				}
			}
		ii=mup-4;
		iii=mup+4;		
		if (iii < 9){iii= 9;}
		if (iii > mupcoon||mupcoon<9){iii= mupcoon;}
		if (iii== mupcoon){ii=mupcoon-8}
		if (ii < 1||mupcoon<9){ii=1;}
		for(i=ii;i<(iii+1);i++){
			if (i!=mup){
				pagectrl+="<li class=\"b\"><a href=javascript:gotoandplay("+i+")>"+i+"</a> ";
			}else{
				pagectrl+="<li class=\"b\"><a class=\"on\">"+i+"</a> ";
			}
		}
		if(IsShowFirstLast){
		if (mupcoon > (mup+4)&&mupcoon>9){
			pagectrl+="<li class=\"c\"><a class=\"no\"> ...</a>"
			pagectrl+="<li class=\"b\"><a href=\"javascript:gotoandplay("+mupcoon+")\" >"+mupcoon+"</a> "
			}
			}
		}
		if(IsShowPreNext){
			if (mup<mupcoon){
				pagectrl+="<li class=\"b\"><a href=\"javascript:gotoandplay("+(mup+1)+")\" ><img src=\"css/img/right3.gif\"></a>";
				pagectrl+="<a href=\"javascript:gotoandplay("+mupcoon+")\" ><img src=\"css/img/right1.gif\"></a>";
			}else{
				pagectrl+="<li class=\"b\"><a class=\"off\"><img src=\"css/img/right2.gif\"></a><a class=\"off\"><img src=\"css/img/right.gif\"></a>";
			}
		}
		pagetmp="<ul>"+pagectrl+"</ul>";
	}else{
		pagetmp="&nbsp;";
	}
	return pagetmp;
}
function selectPage(ipg,ipgcoon){
	var sp;
	sp="<select name='hehe' onChange='javascript:gotoandplay(this.options[this.selectedIndex].value)'>";
	for (t=1;t<=ipgcoon;t++){ 
		if (t!=ipg){sp=sp+"<option value='"+t+"'>"+t+"</option>";}
		else{sp=sp+"<option  selected=\"selected\" value='"+t+"'>"+t+"</option>";}}
	sp=sp+"</select>";
return sp;
}
function externalLinks(){
	if (!document.getElementsByTagName) return;
	var anchors = document.getElementsByTagName("a");
	for (var i=0; i<anchors.length; i++) {
	var anchor = anchors[i];
	if (anchor.getAttribute("href") &&anchor.getAttribute("rel") == "external")
//	rel="external"
	anchor.target = "_blank";
	}
}
function check(){
cztx=$("chazhao").value;
    if(cztx=="") {
    alert("请输入关键字！");
    $("chazhao").focus();
	return false;
	}
	if(!isAvailString(cztx)){
	alert("关键字不要含有非法字符");
	$("chazhao").focus();
	return false;
}
}


function ltrim(s){
return s.replace( /^\s*/, "");
}

function rtrim(s){
return s.replace( /\s*$/, "");
}

function trim(s){
return rtrim(ltrim(s));
}

function isAvailString(strString){
	var ch,temp,isCN,isTrue;
	var myReg = /^[0-9a-zA-Z_]+$/;
	isTrue = true;
	for(i = 0; i < strString.length; i++){
		ch = strString.substring(i,i+1);
		temp = escape(ch);
		isCN = (temp.length == 6)? true:false;
		if(!isCN)
		{
			if(!myReg.test(ch))
			{
			isTrue = false;
			break;
			}
		}
	}
return isTrue;
}
function stringLength(str)
{
  var ch,temp,isCN,isTrue;
  var len = 0;
  isTrue = true;
  for(i = 0; i < str.length; i++)
  {
    ch = str.substring(i,i+1);
    temp = escape(ch);
    isCN = (temp.length == 6)? true:false;
    if(isCN)
    {
      len += 2;
    }
    else
    {
      len += 1;
    }
  }
  return len;
} 

function isNull(strInput){
	if(strInput.length == 0 || strInput == "" || strInput == null)
		return true;
	else
		return false;
	}

function isAvailLen(str,intMin,intMax)
{
                var strLen,intMinLen,intMaxLen;
                if(isNull(str))
                        return false;
                if(isNaN(intMin) || isNaN(intMax))
                        return false;
                if(intMin <= 0 || intMin == null || intMin == "")
                        intMinLen = 0;
                else
                        intMinLen = intMin;
                if(intMax <= intMinLen || intMax == null || intMax == "")
                        intMaxLen = intMinLen;
                else
                        intMaxLen = intMax;

                if(this.stringLength(str) >= intMinLen && this.stringLength(str) <= intMaxLen)
                        return true;
                return false;
        }
function do_select(val,vid){
       var lena=$(val).options.length;
       for(var i=0;i<lena;i++)
        {
            var opvalue = $(val).options[i].value;
             if(opvalue == vid)
                {
                        $(val).options[i].selected = true;
                        break;
                }
        }
}
function HTMLEnCode(str)
{
   var s = "";
   if (str.length == 0) return "";
   s = str.replace(/&/g, "&gt;");
   s = s.replace(/</g,   "&lt;");
   s = s.replace(/>/g,   "&gt;");
   s = s.replace(/ /g,   "&nbsp;");
   s = s.replace(/\'/g,  "&#39;");
   s = s.replace(/\"/g,  "&quot;");
   s = s.replace(/\n/g,  "<br/>");
   return s;
}
function HTMLDeCode(str)
{

   if (str.length == 0) return "";
   str = str.replace(/&lt;/g,   "<");
   str = str.replace(/&gt;/g,   ">");
   str = str.replace(/&quot;/g,  "\"");
   str = str.replace(/&nbsp;/g,   " ");
   str = str.replace(/&#39;/g,  "\'");
  str = str.replace(/<br>/g,  "\n");
   return str;
}

