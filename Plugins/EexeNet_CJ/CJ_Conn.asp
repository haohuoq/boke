<%
on error resume next
'===================================
'文件修改于2007-12-30 by http://eexe.net
'作用：链接动态生成的数据库
'================================================
Dim DBPath,ConnStr,conn,CookieName'
DBpath="."
%>
<!--#include file="BlogDB_Conn.asp"-->
<%
''***************以变量请检查是否正确**************
'Const CookieName="PJBlog2"' 这个名字请与你的blog的const.asp设置相同
'
'DBPath = "../../blogDB/pjblog2.asp"  
' 'ACCESS数据库的文件名，请使用相对于本文件的路径
'
''**************************************************** 
 
 
 
 
 ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(DBPath)
 Set conn = Server.CreateObject("ADODB.Connection")
 conn.open ConnStr
    If Err Then Err.Clear:Set conn = Nothing:Response.Write "On The AccessDB connect Err,please check the plugins setup succeed ":Response.End

Sub CloseConn()
    On Error Resume Next
	Conn.close:Set Conn=nothing
End sub
'%>