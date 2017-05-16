<!--#include file="../Cj_conn.asp"-->
<%
dim db,ErrMsg,ErrMsg_lx,Site
Site="skxiu.com"'Cookies名称，如果您在同一地址内有两个或以上个网站系统，请保证名称不同
dim connstrItem
dim dbItem
dim connItem
dbItem="Database/sk_data.mdb"'采集数据库文件的位置 
Call OpenConnItem
Sub OpenConnItem()
	On Error Resume Next
	Set connItem = Server.CreateObject("ADODB.Connection")
	connstrItem="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(dbItem)
	connItem.Open connstrItem
	If Err Then
	   Err.Clear
	   Set ConnItem = Nothing
	   Response.Write "采集数据库连接出错，请检查连接字串。"
	   Response.End
	End If
End Sub
Sub CloseConnItem()
   On Error Resume Next
   ConnItem.close
   Set ConnItem=nothing
End sub
'---------------------CMS数据库------------------------------------------------------------------
Const CMSDataBase=1 '1=整合CMS. 0=单机版 
'-------------------------------------------------------------------------------------------------------------
%>


