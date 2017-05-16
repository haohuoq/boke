<%
'on error resume next
Session.codepage="65001"
Response.CharSet = "utf-8"
Dim oauth_nonce,oauth_callback,oauth_timestamp,oauth_signature,oauth_token,oauth_token_secret,oauth_verifier,oauth_Cookies_path,oauth_siteurl
'------请修改以下行
Const oauth_consumer_key=""  'App ID	申请地址http://connect.opensns.qq.com/
Const oauth_consumer_secret = ""  'App KEY
Const Time_Zone=0 ' 服务器所在时区与北京时间小时逆差 中国 0 美国 15（仅当提示签名错误时修改，JScript版插件不存在此问题无需修改，亦可用作当服务器时间不准确时进行修正，比如服务器时间比北京时间慢15分钟则写成 Const Time_Zone=0.25 即可。时差在5分钟之内无需修正）
oauth_siteurl="http://www.jinbenli.com/" '博客安装目录，以http://开头，以/结尾
oauth_Cookies_path=""	'当二级域名交互时修改,修改示例 oauth_Cookies_path=".jinbenli.com"
'------以下代码禁止修改
Const oauth_signature_method = "HMAC-SHA1"
Const oauth_version = "1.0"
Const request_token_url="http://openapi.qzone.qq.com/oauth/qzoneoauth_request_token"
Const authorize_url="http://openapi.qzone.qq.com/oauth/qzoneoauth_authorize"
Const access_token_url="http://openapi.qzone.qq.com/oauth/qzoneoauth_access_token"
Const user_info_url ="http://openapi.qzone.qq.com/user/get_user_info"
oauth_callback=strUrlEnCode(oauth_siteurl&"callback.asp")
oauth_token=""
oauth_token_secret=""
%>