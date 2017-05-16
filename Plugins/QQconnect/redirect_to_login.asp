<!--#include file="oauth/class.asp"-->
<!--#include file="oauth/config.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>腾讯QQ一键登录</title>
</head>
<body>
<%
'=========定义
Dim GET_Array,SIGN_Array,Gmethod,i,WB,API_URL,wb_page,wb_json
'=========新建类实例
Set WB=new WBoauth_54bq

'================传参
API_URL=request_token_url
	GET_Array=Array("oauth_consumer_key","oauth_nonce","oauth_signature_method","oauth_timestamp","oauth_version")
	SIGN_Array=GET_Array
	Gmethod="GET"
'=========授权后参数
	wb_page = WB.main()
'WB.die(wb_page) ''debug
if left(wb_page,12)<>"oauth_token=" then WB.die("连接QQ服务器超时，或者您设置的APP ID和APP KEY有误，或者时区设置错误<a href='javascript:void(0);'onload='location.reload()'  target='_self'>再试一次</a>")
	WB.echo("<!--"&wb_page&"-->")
http_restore_query2(wb_page)


If oauth_token<>"" Then
	WB.setcookie "oauth_token_first",oauth_token
	WB.setcookie "oauth_token_secret_first",oauth_token_secret
	WB.echo("<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">")
	WB.echo("<a href='http://openapi.qzone.qq.com/oauth/qzoneoauth_authorize?oauth_token="&oauth_token&"&oauth_consumer_key="&oauth_consumer_key&"&oauth_callback="&oauth_callback&"'  target='_blank'>腾讯QQ一键登录</a>")
	Response.Status="301 Moved Permanently"
	Response.AddHeader "Location", "http://openapi.qzone.qq.com/oauth/qzoneoauth_authorize?oauth_token="&oauth_token&"&oauth_consumer_key="&oauth_consumer_key&"&oauth_callback="&oauth_callback&""
	Response.End
else
	WB.echo("<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">")
	WB.echo("连接QQ服务器超时，或者您设置的APP ID和APP KEY有误，或者时区设置错误<a href='javascript:void(0);'onload='location.reload()'  target='_self'>再试一次</a>")

End If
%>
</body>
</html>