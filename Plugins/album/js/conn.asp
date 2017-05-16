<!--#include file="../sysfile/config.asp" -->
<%
	Dim db, myconn,conn
		db="../"&db&""
		myconn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db)
	On Error Resume Next
		Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open myconn

	If Err Then
		err.Clear
		Set Conn = Nothing
		response.write"数据库连接出错，请检查连接字串。"
'		response.write db
		Response.End
	End If
sub CloseConn()
conn.close
set conn=nothing
end sub
Const SafePass="1,1,4,2,3"
'作用：有人通过下载数据库或SQL注入得到了管理员的真正密码后，仍不能进入系统
'
'第1位	是否启用安全密码 为0时则不启用 为1时启用
'第2位	取验证码中的第几位参与运算，取1-4之间的数字
'第3位	取验证码中的第几位参与运算，取1-4之间的数字
'第4位	将取得的两位验证码作什么运算，1为加法运算；2为乘法运算
'第5位	将得到的结果插入到密码的第几位后面
'
'例如安全码参数设置为1，1，3，2，5  即为启用安全码，将验证码的第一位和第三位相乘的结果插入到密码的第五位后面
'如果你登陆时 产生的验证码为3568 管理员密码为TryLogin
'则你应该输入的密码为TryLo18gin
%>