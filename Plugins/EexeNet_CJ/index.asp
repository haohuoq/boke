<!--#include file="../../const.asp"-->
<!--#include file="../../common/function.asp"-->
<%
'数据库文件读取并转向
'update 2007-12-31
'by eexenet
'
if session("createDBConn")="CreateD" then
	response.Redirect("SK_GetArticle.asp")
else
   Dim dbFile,FileConent,fs,f
	dbFile=Server.MapPath("BlogDB_Conn.asp")
	FileConent="<%CookieName=" &"""" & CookieName &""""& vbcrlf & "DBPath=" &"""" &"../../" & AccessFile & """" &"%" & ">"
	set fs=server.CreateObject("Scripting.FileSystemObject")
	fs.CreateTextFile(dbFile)
	if fs.fileExists(dbFile) then
		set f=fs.OpenTextFile(dbFile,2,false)
		f.write FileConent 
		else
		Response.write("数据库链接文件BLogDB_conn.asp不存在，请确定该文件存在，或者你还没有安装本插件")
		Response.End()
	End if
	f.close
	set fs=nothing
	session("createDBConn")="CreateD"'记录一次数据
 	response.Redirect("SK_GetArticle.asp")
End if

%>
