<!--#include file="conCommon.asp" -->
<!--#include file="common/function.asp" -->
<!--#include file="common/library.asp" -->
<!--#include file="common/cache.asp" -->
<!--#include file="common/checkUser.asp" -->
<!--#include file="common/ubbcode.asp" -->
<!--#include file="common/XML.asp" -->
<!--#include file="common/ModSet.asp" -->
<!--#include file="class/cls_logAction.asp" -->
<!--#include file="class/cls_article.asp" -->
<%
'----------- 释放网站缓存 ---------------------------

Function FreeApplicationMemory
    On Error Resume Next
    Response.Write "释放网站缓存数据列表：<div style='padding:5px 5px 5px 10px;'>"
    Dim Thing,i
    i=0
    For Each Thing IN Application.Contents
    	
        If Left(Thing, Len(CookieName)) = CookieName Then
       		i=i+1
        	if i<30 then
            	Response.Write "<span style='color:#666'>" & thing & "</span><br/>"
            elseif i<31 then
            	Response.Write "<span style='color:#666'>...</span><br/>"
            end if
            
            If IsObject(Application.Contents(Thing)) Then
                Application.Contents(Thing).Close
                Set Application.Contents(Thing) = Nothing
                Application.Contents.Remove(Thing)
            ElseIf IsArray(Application.Contents(Thing)) Then
                Set Application.Contents(Thing) = Nothing
                Application.Contents.Remove(Thing)
            Else
                Application.Contents.Remove(Thing)
            End If
        End If
        
    Next
    response.Write "<br/><span style='color:#666'>共清理了 " & i & " 个缓存数据</span><br/>"
    response.Write "</div>"
End Function
function localurl()
	Dim sUrl
	if len(Request.ServerVariables("QUERY_STRING")) > 0 then
		sUrl = Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" &Request.ServerVariables("QUERY_STRING")
	else
		sUrl = Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")
	end if
	sUrl = "http://" & sUrl
	localurl = sUrl
end function
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>PJBLOG缓存更新程式</title>
<style type="text/css">
	body{ font:13px/160% 'Microsoft YaHei',helvetica,Arial,Tahoma,Sans-Serif; background:#fcfcfc; color:#373933}
</style>
</head>

<body>
<%
				Response.Write "<div style='padding:4px 0px 4px 10px;border: 1px dotted #999;margin:2px;background:#ffffee'>"
		        Application.Lock
		        FreeApplicationMemory
		        Application.UnLock
		        Response.Write "<br/><span><b style='color:#040'>缓存清理完毕...	</b></span>"
		        Response.Write "</div>"
%><br />
<input type="button" value="继续刷新本站缓存" onclick="location.href='<%=localurl()%>'">
</body>
</html>
