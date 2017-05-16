<%
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
%>
<!--#include file="inc/setup.asp"-->
<!--#include file="SK_Session.asp"-->
<%
Dim Rs,Sql,SqlItem,RsItem,Action,FoundErr
Dim HistrolyID,ItemID,ChannelID,ClassID,SpecialID,ArticleID,Title,CollecDate,NewsUrl,Result
Dim  Arr_Histroly,Arr_ArticleID,i_Arr,Del,Flag
Dim MaxPerPage,CurrentPage,AllPage,HistrolyNum,i_His
MaxPerPage=20
FoundErr=False
Del=Trim(Request("Del"))
Action=Trim(Request("Action"))
If Del="Del" Then
   Call DelHistroly()
End If
If FoundErr<>True Then
   Call Main()
else
  Call WriteErrMsg(ErrMsg)
End If
'关闭数据库链接
Call CloseConnItem()
%>

<%Sub Main%>
<html>
<head>
<title>新闻采集系统-history</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" type="text/css" href="css/Admin_Style.css">
<style type="text/css">
.ButtonList {
	BORDER-RIGHT: #000000 2px solid; BORDER-TOP: #ffffff 2px solid; BORDER-LEFT: #ffffff 2px solid; CURSOR: default; BORDER-BOTTOM: #999999 2px solid; BACKGROUND-COLOR: #e6e6e6
}
</style>
<SCRIPT language=javascript>
function unselectall(thisform)
{
    if(thisform.chkAll.checked)
	{
		thisform.chkAll.checked = thisform.chkAll.checked&0;
    } 	
}

function CheckAll(thisform)
{
	for (var i=0;i<thisform.elements.length;i++)
    {
	var e = thisform.elements[i];
	if (e.Name != "chkAll"&&e.disabled!=true)
		e.checked = thisform.chkAll.checked;
    }
}
</script>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table cellpadding="0" cellspacing="1" border="0" width="100%" class="tableBorder" align=center>
  <tr class="topbg">
	<td height="22" colspan="2" align="left"><!--#include file="CJ_Pjblog.asp"--></td>
  </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="tableBorder">
  <tr class="tdbg"> 
    <td height="30" width="65"><strong>管理导航：</strong></td>  
    <td height="30"><a href="Sk_ItemHistroly.asp">管理首页</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="Sk_ItemHistroly.asp?Action=Succeed">成功记录</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="Sk_ItemHistroly.asp?Action=Failure">失败记录</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href="Sk_ItemHistroly.asp?Action=LoseEf">失效记录</a></td>     
  </tr>         
</table>     
<%                             
Set RsItem=server.createobject("adodb.recordset")         
SqlItem="select * from Histroly"
If Action="Succeed"  Then
   SqlItem=SqlItem  &  " Where  Result=True"
   Flag="成 功 记 录"
ElseIf Action="Failure"  Then
   SqlItem=SqlItem  &  " Where  Result=False"
   Flag="失 败 记 录"
ElseIf  Action="LoseEf"  Then
   Flag="失 效 记 录"
   Set Rs=server.createobject("adodb.recordset")
   Sql="Select ArticleID From SK_Article Where Deleted=False"
   Rs.open  Sql,ConnItem,1,1
   If (Not Rs.Eof) And (Not Rs.Bof) Then
      Do While not rs.eof
         Arr_ArticleID=Arr_ArticleID & "," & CStr(rs("ArticleID"))
      Rs.MoveNext
      Loop
   End  If
   Rs.Close
   Set Rs=Nothing
   If Arr_ArticleID<>"" Then
      Arr_ArticleID=0 & Arr_ArticleID
   Else
      Arr_ArticleID=0
   End If
   SqlItem=SqlItem  &  " Where ArticleID Not In("  &  Arr_ArticleID & ")"
Else
   Flag="所 有 记 录"
End  If
%>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="tableBorder" >
    <tr> 
      <td height="22" colspan="2" class="title"> <div align="center"><strong>历  史  记  录 －　<%=Flag%></strong></div></td>
    </tr>
</table>
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="tableBorder">         
   <form name="form1" method="POST" action="?">         
     <tr class="tdbg" style="padding: 0px 2px;">         
      <td width="56" height="22" align="center" bgcolor="#EFEFEF" >选择</td>                 
      <td width="214" align="center" bgcolor="#EFEFEF" >项目名称</td>         
      <td width="435" align="center" bgcolor="#EFEFEF">标题</td>
      <td width="123" height="22" align="center" bgcolor="#EFEFEF" >频道</td> 
      <td width="120" height="22" align="center" bgcolor="#EFEFEF" >栏目</td> 
      <td width="126" align="center" bgcolor="#EFEFEF" >来源</td>        
      <td width="87" align="center" bgcolor="#EFEFEF" >结果</td>
      <td width="93" height="22" align="center" bgcolor="#EFEFEF" >操作</td>         
     </tr>         
<%
If Request("page")<>"" then
    CurrentPage=Cint(Request("Page"))
Else
    CurrentPage=1
End if 
SqlItem=SqlItem  &  " order by HistrolyID DESC"
RsItem.open SqlItem,ConnItem,1,1
If (Not RsItem.Eof) and (Not RsItem.Bof) then
   RsItem.PageSize=MaxPerPage
   Allpage=RsItem.PageCount
   If Currentpage>Allpage Then Currentpage=1
   HistrolyNum=RsItem.RecordCount
   RsItem.MoveFirst
   RsItem.AbsolutePage=CurrentPage
   i_His=0
   Do While not RsItem.Eof
%>
    <tr class="tdbg"  onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#cccccc'" style="padding: 0px 2px;">          
      <td width="56" align="center" style="border-bottom:#999999 1xp solid">          
        <input type="checkbox" value="<%=RsItem("HistrolyID")%>" name="HistrolyID" onClick="unselectall(this.form)" style="border: 0px;background-color: #E1F4EE;">         
      </td>               
      <td width="214" align="center" style="border-bottom:#999999 1xp solid">          
        <%Call Admin_ShowItem_Name(RsItem("ItemID"))%>         
      </td>         
      <td width="435" align="left" style="border-bottom:#999999 1xp solid"><%=RsItem("Title")%>
      </td>  
      <td width="123" align="center" style="border-bottom:#999999 1xp solid"><%Call Admin_ShowChannel_Name(RsItem("ChannelID"))%></td>
      <td width="120" align="center" style="border-bottom:#999999 1xp solid"><%Call Admin_ShowClass_Name(RsItem("ChannelID"),RsItem("ClassID"))%></td>
      <td width="126" align="center" style="border-bottom:#999999 1xp solid"><a href="<%=RsItem("NewsUrl")%>" target=_blank title=<%=RsItem("NewsUrl")%>>点击访问</a></td>     
      <td width="87" align="center" style="border-bottom:#999999 1xp solid">
      <%If RsItem("Result")=True Then
           Response.write "成功"
        ElseIf RsItem("Result")=False Then
           Response.Write "<font color=red>失败</font>"
        Else
           Response.Write "<font color=red>异常</font>"
        End If
      %>
      </td>
      <td width="93" align="center" style="border-bottom:#999999 1xp solid">                           
      <a href="Sk_ItemHistroly.asp?Action=<%=Action%>&Del=Del&HistrolyID=<%=RsItem("HistrolyID")%>" onclick='return confirm("确定要删除此记录吗？");'>删除</a>                 
      </td>         
    </tr>         
<%         
           i_His=i_His+1
           If i_His > MaxPerPage Then
              Exit Do
           End If
        RsItem.Movenext         
   Loop         
%>         
    <tr class="tdbg">          
      <td colspan=8 height="30">       
        <input name="Del" type="hidden" id="Del" value="Del">   
        <input name="Action" type="hidden" id="Action" value="<%=Action%>"> 
        <input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox" style="border: 0px;background-color: #E1F4EE;">全选
      </td>         
    </tr> 
    <tr class="tdbg">          
      <td colspan=8 height="30" align=center>
        <input type="submit" value="清除选中记录" name="DelFlag" onclick='return confirm("确定要清除所选记录吗？");' style="cursor: hand;background-color: #cccccc;">&nbsp;&nbsp;       
        <input type="submit" value="清除失败记录" name="DelFlag" onclick='return confirm("确定要清除所有失败记录吗？");' style="cursor: hand;background-color: #cccccc;">&nbsp;&nbsp;
        <input type="submit" value="清除失效记录" name="DelFlag"  onclick='return confirm("确定要清除所有失效记录吗？");' style="cursor: hand;background-color: #cccccc;">&nbsp;&nbsp;
        <input type="submit" value="清空所有记录" name="DelFlag" onclick='return confirm("确定要清除所有记录吗？");' style="cursor: hand;background-color: #cccccc;">
      </td>         
    </tr> 
    <tr class="tdbg">          
      <td colspan=8 height="30">          
        
      </td>         
    </tr>       
<%Else%>
<tr class="tdbg">
        <td colspan='9' class="tdbg" align="center"><br>系统中暂无历史记录！</td>
</tr> 
<%End  If%>       
<%         
RsItem.Close         
Set RsItem=nothing           
%>         
</form>         
</table>  
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="tableBorder" >
    <tr> 
      <td height="22" colspan="2" class="tdbg">
<%
Response.Write ShowPage("Sk_ItemHistroly.asp?Action="& Action,HistrolyNum,MaxPerPage,True,True," 个记录")
%>

      </td>
    </tr>
</table>
<!--#include file="Admin_ItemFoot.asp"-->           
</body>         
</html>
<%End Sub%>
<%Sub DelHistroly

Dim DelFlag
DelFlag=Trim(Request("DelFlag"))
HistrolyID=Trim(Request("HistrolyID"))
If HistrolyID<>"" Then
   HistrolyID=Replace(HistrolyID," ","")
End If
If DelFlag="清除选中记录" Then
   If HistrolyID="" Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>请选择要删除的记录</li>"
   Else
      HistrolyID=Replace(HistrolyID," ","")
      SqlItem="Delete From [Histroly] Where HistrolyID in(" & HistrolyID & ")"
   End If
ElseIf DelFlag="清除失败记录" Then
   SqlItem="Delete From [Histroly] Where Result=False"
ElseIf DelFlag="清除失效记录" Then
   Set Rs=server.createobject("adodb.recordset")
   Sql="Select ArticleID From SK_Article Where Deleted=False"
   Rs.open Sql,ConnItem,1,1
   If (Not Rs.Eof) And (Not Rs.Bof) Then
      Do While not rs.eof
         Arr_ArticleID=Arr_ArticleID & "," & CStr(rs("ArticleID"))
      Rs.MoveNext
      Loop
   End  If
   Rs.Close
   Set Rs=Nothing
   If Arr_ArticleID<>"" Then
      Arr_ArticleID=0 & Arr_ArticleID
      SqlItem="Delete From [Histroly] Where ArticleID Not In(" & Arr_ArticleID & ")"
   Else
      SqlItem="Delete From [Histroly]"
   End If
ElseIf DelFlag="清空所有记录" Then
   SqlItem="Delete From [Histroly]"
Else
   If HistrolyID="" Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>请选择要删除的记录</li>"
   Else
      HistrolyID=Replace(HistrolyID," ","")
      SqlItem="Delete From [Histroly] Where HistrolyID In(" & HistrolyID & ")"
   End If
End if

If FoundErr<>True Then
   ConnItem.Execute(SqlItem)
End If
End Sub
%>
