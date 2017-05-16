var content ="<div class='tixu'><ul><li><div class='tipxu'>序</div>"
+"<div class='timxu'>图片名称</div>"
+"<div class='tiuxun'>图片地址</div>"
+"<div class='tizxun' id='topzt'></div></li><li>"
+"<div class='tipxu1' id='paixu'></div>"
+"<div class='timxu1'><input type='text'id='pnames' size='12' class='ane' value='' /></div>"
+"<div class='tiuxun'>"
+"<div>缩略图:<input type='text' id='sphoto'size='58' class='ane' value='' /></div>"
+"<div>原始图:<input type='text'id='bphoto'  size='58' class='ane'  value='' /></div>"
+"</div><div class='tizxun' id='fanzt'></div></li></ul>"
+"<input type='button' class='an1' id='Submita' name='Submitb' disabled='disabled' onclick='picplay();' value='开始' />"
+"<input type='button' class='an1' id='Submitb' name='Submitb' onclick='playclose();' value='取消' /></div>"

function delloadingtext(){
	if(innerPlace.tagName!="DIV"){
		innerPlace.getElementsByTagName("a")[1].innerHTML = innerPlace.getElementsByTagName("a")[1].innerHTML.substr(0,(innerPlace.getElementsByTagName("a")[1].innerHTML.length-loadingtext.length));
	}
}
var onsuccess = "";
var onloading = "";
var onerrors = "页面不存在";
var loadingtext = "<span>..</span>";

var innerPlace="";
function innerloadingtext(obj){
	if(obj.parentNode.getElementsByTagName("a")[0].className=="open"){
		try{
			ulobj = obj.parentNode.getElementsByTagName("ul")[0];
			}catch(e){ulobj = "";}
	if(typeof(ulobj)!="object")
		obj.innerHTML += loadingtext;
	}
}
function switchclass(id,obj){
	if(obj.className=="close"){
		obj.className="open"; 
		try{
			ulobj = obj.parentNode.getElementsByTagName("ul")[0];
			}catch(e){ulobj = "";
		}
	if(typeof(ulobj)=="object"){
		if(ulobj.style.display="none"){
			ulobj.style.display="block";
		}else{
			getValue(id,obj.parentNode);
		}
	}else{
		getValue(id,obj.parentNode);
	}
	}else{
	if(obj.className=="open"){
		obj.className="close";obj.parentNode.getElementsByTagName("ul")[0].style.display="none";
		}
	}
}
function getValue(wname,obj){ //取值
 	AJAXCALL("savepic.asp?act=path&id="+ $U(wname),"getsuccess","isXML=true");
	innerPlace = obj;
}
function getsuccess(xmlResult){ //取值成功后
	var addr = xmlResult.getElementsByTagName("address");
//	var adder = xmlResult.getElementsByTagName("adderr");
	var addrLen = addr.length;
//	if(adder.length!=0){
//	setTimeout("delloadingtext()",200);
//	return false;
//	}
	if(addrLen==0){
		setTimeout("delloadingtext()",200)
		innerPlace.firstChild.className="nonedata";
	}else{
		var subnode="";
		var selfid;
		var name;
		var url;
		for (i=0; i<addrLen; i++){
			selfid = addr[i].firstChild.data.split("|")[0];
			name = addr[i].firstChild.data.split("|")[1];
			url = addr[i].firstChild.data.split("|")[2];
			subnode = subnode + "<li><a href='javascript:void(0)' onclick=\"switchclass(\'"+ selfid +"\',this);innerloadingtext(this.parentNode.getElementsByTagName('a')[1])\" onfocus=\"this.blur()\" class='close'> </a><a class=on  onclick=\"myClick(this,\'"+ url +"\')\">" + name + "</a></li>"; 
		}
	subnode = "<ul>" + subnode + "</ul>";
	setTimeout("delloadingtext()",200);
	innerPlace.innerHTML = innerPlace.innerHTML + subnode;
	}
}
function getting(){ //正在取值时
// innerPlace.innerHTML = "Loading...";
}
function getfailed(){ //取值失败
	alert(onerrors);
}
var oldObj1=null;
function myClick(obj,pul){
	if(oldObj1!=null)oldObj1.style.backgroundColor="";
	obj.style.backgroundColor="lightblue";
	oldObj1=obj;
	loadinp(pul);
}

function changeselx(vid,xid,cooid,zhid){
var dc=$(vid);
	hhid=zhid;
	atid=$(vid).value;
	if(zhid==1){ hpid=atid;	}
	else{ if($("picname")){AJAXCALL("js/saveedt.asp?act=gtname&id="+atid,"Shpname");}
		
		$("spabtex"+xid).innerHTML="";
	}
	
}
function ShSelect(smdoc){
if (smdoc){
	var smbar=eval(smdoc);
	smbar=smbar[0];
	_selst=smbar.selst;
	_spidtx=smbar.spidtx;
	if(hhid==1){
		if(_selst==1){
			$("spantex"+_spidtx).innerHTML=smbar.selhtml;
		}else{
			_dcount=smbar.dcount;
			_gpage=smbar.page;
			_gtotalPages=smbar.pages;
			$("showminors").innerHTML=smbar.pagehtm;
			xdaohang(_dcount,_gpage,_gtotalPages)
		}
	}else{
		$("selx"+_spidtx).innerHTML=smbar.selhtml;
	}
}else{$("showminors").innerHTML="系统错误！";}

}
function Shpname(pndoc){
	if (pndoc){
		$("picname").value=pndoc;
	}
}


function delTr(evt){
	evt = evt? evt: window.event;
	var srcElem = (evt.target)? evt.target: evt.srcElement
  theTD=srcElem.parentNode.parentNode;
 theTD.parentNode.removeChild(theTD);
}


function geshu(suzid,c){
	t=suzid+"";
	while(t.length<c){t="0"+t;}
	return t
}

function DialogClass(){ 
    this.blankImgHandle = null; 
    this.tags = new Array("applet", "iframe","select","object","embed"); 
    this.body_overflow_y = null; 
    this.body_overflow_x = null; 
    
    this.hideAllSelect = function() 
    { 
        
       for( var k=0;k<this.tags.length;k++) 
       { 
           var selects = document.getElementsByTagName(this.tags[k]);     
           for (var i=0;i<selects.length;i++) 
           { 
               selects[i].setAttribute("oldStyleDisplay",selects[i].style.visibility); 
               selects[i].style.visibility = "hidden"; 
           } 
       }        
    } 
     
    this.resetAllSelect = function() 
    {        
       for( var k=0;k<this.tags.length;k++) 
       { 
           var selects = document.getElementsByTagName(this.tags[k]);     
           for (var i=0;i<selects.length;i++) 
           {            
               if (selects[i].getAttribute("oldStyleDisplay")!=null) 
                  selects[i].style.visibility = selects[i].getAttribute("oldStyleDisplay"); 
           } 
        }     
    }         
         
   this.show = function(sdiv) 
   {     

	this.hideAllSelect(); 
             
        var w = document.body.scrollWidth ; 
        var h = document.body.scrollWidth ; 
        this.blankImgHandle = document.createElement("DIV"); 
        with (this.blankImgHandle.style){ 
            position = "absolute"; 
            left     = 0; 
            top      = 0; 
            height   = document.body.offsetHeight+"px"; 
            width    = document.body.offsetWidth+"px"; 
            zIndex   = "5"; 
            filter   = "progid:DXImageTransform.Microsoft.Alpha(style=0,opacity=80)"; 
            opacity  = "0.7"; 
            backgroundColor = "#cccccc";              
        }
        document.body.appendChild(this.blankImgHandle);
	if($(sdiv)) {
		var clientHeight = scrollTop = 0;
		$(sdiv).style.display ="block"

		clientHeight = document.documentElement.clientHeight;
		scrollTop = document.documentElement.scrollTop;
		$(sdiv).style.top =(scrollTop + 50)+"px";
	}   
   }         
         
   this.close = function(ndiv){
	if($(ndiv)) {
	$(ndiv).style.display ="none"
	}            
//      document.body.style.overflowY = this.body_overflow_y;                
//      document.body.style.overflowX = this.body_overflow_x; 
      if (this.blankImgHandle) 
      { 
          this.blankImgHandle.parentNode.removeChild(this.blankImgHandle); 
      }       
      this.resetAllSelect();            
  } 
}
