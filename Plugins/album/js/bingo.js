if(!bingo){
var bingo={"maintainer":"yoyo","build":"0.3.024 Alpha","lastmodify":"2007.9.13"};
function AJAXCALL(url,handleResponse,param){
var config={"isXML":false,"isCache":false,"method":"GET","errorReport":true,"statePoll":null,"postData":null};var attach;
if(param && typeof param=="string"){
var re = /\s/g;param=param.replace(re,"");var tmp=param.split(",");
for(var i=0;i<tmp.length;i++){var pa=tmp[i].substr(0,3);
var ct=tmp[i].split("=")[1];switch (pa){case "isX":if(ct=="true"){config.isXML= true;}else{config.isXML=false;}break;case "isC":if(ct=="true"){config.isCache= true;}else{config.isCache= false;}break;case "met":config.method=ct;break;case "att":attach=ct;break;}}}
else if(param && typeof param=="object"){if(param.isXml){config.isXML=param.isXml;}if(param.isXML){config.isXML=param.isXML;}if(param.isxml){config.isXML=param.isxml;}if(param.isCache){config.isCache=param.isCache;}if(param.iscache){config.isCache=param.iscache;}if(param.method){config.method=param.method;}if(param.attach){attach=param.attach;}if(param.statePoll){config.statePoll=param.statePoll;}}
if(!config.isCache){var nocache=new Date().getTime();if(url.indexOf("?")>0){url+="&nocache="+nocache;}else{url+="?nocache="+nocache;}}
if(config.method=="POST"){var purl=url.split("?");url=purl[0];config.postData=purl[1];}
var newCall=new AjaxCore();
newCall.init(url,handleResponse,config,attach);
}
function AJAXFORM(formID,handleResponse,param){
var config={"isXML":false,"isCache":false,"method":"POST","errorReport":true,"statePoll":null,"postData":null};var attach;
if(param && typeof param=="string"){var re=/\s/g;param=param.replace(re,"");var tmp=param.split(",");for(var i=0;i<tmp.length;i++){var pa=tmp[i].substr(0,3);var ct=tmp[i].split("=")[1];switch (pa){case "isX":if(ct=="true"){config.isXML=true;}else{config.isXML=false;}break;case "att":attach=ct;break;}
}}else if(param && typeof param=="object"){if(param.isXml){config.isXML=param.isXml;}if(param.isxml){config.isXML=param.isxml;}if(param.attach){attach=param.attach;}if(param.statePoll){config.statePoll=param.statePoll;}}
var url=$(formID).action;
config.postData=getForm(formID);
var FormCall=new AjaxCore();
FormCall.init(url,handleResponse,config,attach);
}
function getForm(theFormName){
   var form = document.forms[theFormName];
   if(typeof form.name == "undefined" || form.name===""){form = $(theFormName);}
   var formData = "",element;
   for (var i = 0; i < form.elements.length; i++) {
      element = form.elements[i];if(!element.type){continue;}
      var type=element.type.toLowerCase();    
      if (type == "hidden" ||type == "text" || type == "password" || type == "textarea"){	
       formData += element.name + "=" + $E(element.value) + "&";
	  }else if (element.type.indexOf("select") != -1) {
         for (var j = 0; j < element.options.length; j++) {
            if (element.options[j].selected === true){formData +=  element.name + "=" + $E(element.options[element.selectedIndex].value) + "&";}
         }
      }
      else if(element.type=="checkbox"){
		 if (element.checked){formData += element.name + "=" + $E(element.value) + "&";}
	  }
      else if(element.type=="radio"){
		  if (element.checked === true){formData += element.name + "=" + $E(element.value) + "&";}
	  }
   }
   return formData.substring(0, formData.length - 1);
}
function AjaxCore(){
	var _state,_status,_self=this,_httpRequest,_remoteFile,_handleFunction,_attach;
	var _config={"isXML":false,"isCache":false,"method":"GET","errorReport":true,"statePoll":null,"postData":null};
	this.init=function(remoteUrl,handleFunction,config,attach){
		if(!remoteUrl){_self.alertf("Error:Null URL");return;}
		_remoteFile=remoteUrl;
		_handleFunction=handleFunction||function(){};
		if(config.isXML){_config.isXML=config.isXML;}
		if(config.isCache){_config.isCache=config.isCache;}
		if(config.method){_config.method=config.method;}
		if(config.postData){_config.postData=config.postData;}
		if(config.errorReport){_config.errorReport=config.errorReport;}
		if(config.statePoll && typeof config.statePoll=="function"){_config.statePoll=config.statePoll;
		}
		_attach=attach?attach:null;
		_self.createRequest();
	};
	this.createRequest=function(){
		if (window.XMLHttpRequest){
            _httpRequest = new XMLHttpRequest();
            if (_httpRequest.overrideMimeType) {
                _httpRequest.overrideMimeType('text/xml');
            }
        }else if (window.ActiveXObject) {
            var MSXML=['MSXML2.XMLHTTP.6.0','MSXML2.XMLHTTP.3.0','MSXML2.XMLHTTP.5.0','MSXML2.XMLHTTP.4.0','MSXML2.XMLHTTP', 'Microsoft.XMLHTTP'];
			for(var n=0;n<MSXML.length;n++){ 
				try { 
					_httpRequest=new ActiveXObject(MSXML[n]); 
					break; 
				}catch(e){}
			} 
        }
		with(_httpRequest) {
			onreadystatechange=_self.getResponse;
			open(_config.method,_remoteFile,true);
			if(_config.method=="POST" && _config.postData){
				setRequestHeader("Content-Length",_config.postData.length);   
				setRequestHeader("Content-Type","application/x-www-form-urlencoded");
				send(_config.postData);
			}else{
				var textType=_config.isXML?"text/xml":"text/plain";
				_httpRequest.setRequestHeader('Content-Type',textType);
				if(browser.IE){
                    setRequestHeader("Accept-Encoding", "gzip, deflate");
				}else if(browser.FF){
                    setRequestHeader("Connection","close");
				}
			    send(null);
			}
		}
	};
	this.getResponse=function(){
		_state=_httpRequest.readyState;
		if(_httpRequest.readyState==4 && _httpRequest.status){_status=_httpRequest.status;}
		if(_config.statePoll){_config.statePoll(_httpRequest.readyState);}
		if(_httpRequest.readyState==4 && _httpRequest.status>=400){_self.abort();_self.alertf("ERROR:HTTP response code "+_httpRequest.status);}
		if(_httpRequest.readyState==4 && _httpRequest.status==200){
			var response_content;
			if(_config.isXML){
				response_content=_httpRequest.responseXML;
			}else{
				response_content=_httpRequest.responseText;	
			}		
			if(typeof _handleFunction=="function"){
				 _handleFunction(response_content,_attach);
			}else{
				 eval(_handleFunction+"(response_content,_attach)");/*somebody said:"eval is evil" ^o^*/
			}
		}
	};
	this.abort=function(){_httpRequest.abort();};
	this.state=function(){return _state;};
	this.status=function(){return _status;};
	this.destory=function(){_self.abort();delete(_httpRequest);};
	this.alertf=function(error){if(_config.errorReport){alert(error);}};
}}
/*tools*/
if(!browser){
var browser={};browser.IE=browser.ie=window.navigator.userAgent.indexOf("MSIE")>0;browser.Firefox=browser.firefox=browser.FF=browser.MF=navigator.userAgent.indexOf("Firefox")>0;browser.Gecko=browser.gecko=navigator.userAgent.indexOf("Gecko")>0;browser.Safari=browser.safari=navigator.userAgent.indexOf("Safari")>0;browser.Camino=browser.camino=navigator.userAgent.indexOf("Camino")>0;browser.Opera=browser.opera=navigator.userAgent.indexOf("Opera")>0;browser.other=browser.OT=!(browser.IE || browser.FF || browser.Safari || browser.Camino || browser.Opera);
}
if(!client){var client={left:function(eleWidth){var eleWidth=eleWidth||100;return (parseInt(document.documentElement.clientWidth)-parseInt(eleWidth))/2;},
top:function(eleHeight){var eleHeight=eleHeight||100;return (parseInt(document.documentElement.clientHeight)-parseInt(eleHeight))/2;}};}
if(!getTag){var getTag=function(xmlDoc, tag ,i){
     if(typeof xmlDoc!="object"){alert("Error:Not A XML Document,getTag");return;}
	 var i=i||0;
     var elems = xmlDoc.getElementsByTagName(tag);
	 var elem=elems[i];var rtEle={};
	 rtEle.node=elem;
	 rtEle.text=elem.text||elem.textContent;
	 rtEle.type=elem.nodeType;
     return rtEle;
};}
if(!$){
var $=function(obj){return document.getElementById(obj);};
function $F(obj){return document.getElementById(obj).value;}
function $U(str){return encodeURIComponent(escape(str));}
function $E(str){return encodeURIComponent(str);}
function $C(tag,style){var tag=tag||"div";var c=document.createElement(tag);if(style && typeof style=="object"){bindStyle(c,style);}return c;}var $c=$C;}
if(!$getCSS){var $getCSS=function(id,rule){
  for(var i=0;i<document.styleSheets.length;i++){
		var tmp=browser.IE?document.styleSheets[i].rules:document.styleSheets[i].cssRules;
		for(var k=0;k<tmp.length;k++){
			if(tmp[k].selectorText==id){
				if(rule){
					return tmp[k].style[rule];
				}else{
					return tmp[k].style;
				}			
			}
		}
}};
var $css=$getCSS;var $getcss=$getCSS;}
if(!bindStyle){var bindStyle=function(ele,style){
  for(var p in style){
	  if(p.toString()=="toJSONString"){continue;}
	  if(p.toString()=="float"){
		  if(browser.IE){
			  ele.style.styleFloat=style[p];
		  }else if(browser.FF){
			  ele.style.cssFloat=style[p];
		  }else{
			  ele.style.float=style[p];/*float is a reserved word*/
		  }		  
	  }else if(p.toString()=="alpha"){
		  if(browser.IE){
			  ele.style.filter="Alpha(Opacity="+parseFloat(style[p])*100+")";
		  }else if(browser.FF){
			  ele.style.MozOpacity=style[p];
		  }else{
			  ele.style.opacity=style[p];/* CSS3 */
		  }	
	  }else if(p.toString()=="cursor" && (style[p]=="hand" || style[p]=="pointer")){
           ele.style.cursor="pointer";
      }else{
	       try{ele.style[p]=style[p];}catch(e){}
	  }
  }
};
var bindstyle=bindStyle}
if(!getElementsByClass){var getElementsByClass=function(searchClass,node,tag){
	var classElements=new Array();
	if (!node){node=document;}
	if (!tag){tag='*';}
	var els = node.getElementsByTagName(tag);
	var elsLen = els.length;
	var pattern = new RegExp("(^|\s)"+searchClass+"(\s|$)");
	for(var i = 0,j = 0; i < elsLen; i++) {
		if( pattern.test(els[i].className)){classElements[j] = els[i];j++;}
	}
	return classElements;
};}