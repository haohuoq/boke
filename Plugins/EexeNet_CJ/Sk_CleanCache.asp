<!--#include file="inc/clsCache.asp"-->

<%
   Dim myCache
   Set myCache=new SK_clsCache
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

Call CleanCache
	  
	  Sub CleanCache()
		With Response
		  .Write "<html>"
		  .Write "<head>"
		  .Write "<title>缓存更新</title>"
		  .Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
		  .Write "<link href=""css/Admin_Style.css"" rel=""stylesheet"" type=""text/css"">"
		  .Write "</head>"
		  .Write "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
		
					.Write "    <table width=""95%"" style=""MARGIN-TOP: 3px"" border=""0"" align=""center""cellspacing=""1"" cellpadding=""1"" class=""tableBorder"">"
					.Write "    <tr height=24 class='Title'><td width='550' align='center'><b>更新对象</b></td></tr>"
		.Write "    <tr height=24 class='Title'><td width='40' align=center>"	
		For Each Cacheobj in Application.Contents
			Response.Write(Cacheobj) &"</br >"
			DelCahe(Cacheobj)
		Next
		Application.Lock   
 		Application.Contents.RemoveAll   
  		Application.Unlock
		.Write "</td></tr>"  
		.Write "<script>function back(){alert('所有缓存更新完，按确定返回！');history.back();}setTimeout('back()',800);</script>"
		.Write "</body>"
		.Write "</html>"
    End With
    End Sub     
	'=================================================缓存相关函数=======================
	'不提示,批量清除缓存,参数 PreCacheName-前段匹配
	Public Sub DelCaches(PreCacheName)
	    Dim i
		Dim CacheList:CacheList=split(GetCacheList(PreCacheName),",")
		If UBound(CacheList)>1 Then
			For i=0 to UBound(CacheList)-1
				DelCahe CacheList(i)
			Next
		End IF
	End Sub
	'取得缓存列表 参数 PreCacheName-前段匹配
	Public Function GetCacheList(PreCacheName)
		Dim Cacheobj
		For Each Cacheobj in Application.Contents
		If CStr(Left(Cacheobj,Len(PreCacheName)))=CStr(PreCacheName) Then GetCacheList=GetCacheList&Cacheobj&","
		Next
	End Function
	'清除缓存,参数 MyCaheName-缓存名称
	Public Sub DelCahe(MyCaheName)
		Application.Lock
		Application.Contents.Remove(MyCaheName)
		Application.unLock
	End Sub
%> 
