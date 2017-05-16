﻿<%
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
option explicit
response.buffer=true	
%>
<!--#include file="inc/setup.asp"-->
<!--#include file="SK_Session.asp"-->
<%
dim Action,Rs,Sql,RsItem,SqlItem,ItemID,ItemName,ClassID,SpecialID,Flag,FoundErr
Dim ObjInstalled,tClass,tSpecial,copyitem(149),i
ObjInstalled=IsObjInstalled("Scripting.FileSystemObject")
Action=trim(request("Action"))
%>      
<html>
<head>
<title>数据采集系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<link rel="stylesheet" type="text/css" href="css/Admin_Style.css">
<style type="text/css">
.ButtonList {
	BORDER-RIGHT: #000000 2px solid; BORDER-TOP: #ffffff 2px solid; BORDER-LEFT: #ffffff 2px solid; CURSOR: default; BORDER-BOTTOM: #999999 2px solid; BACKGROUND-COLOR: #e6e6e6
}
</style>
<SCRIPT language=javascript>
    function unselectall(thisform){
        if(thisform.chkAll.checked){
            thisform.chkAll.checked = thisform.chkAll.checked&0;
        }   
    }
    function CheckAll(thisform){
        for (var i=0;i<thisform.elements.length;i++){
            var e = thisform.elements[i];
            if (e.Name !="chkAll"&&e.disabled!=true)
                e.checked = thisform.chkAll.checked;
        }
    }
</script>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="tableBorder">
  <tr class='topbg'> 
    <td height="22" colspan="2" align="left" ><!--#include file="CJ_PJblog.asp"--></td>
  </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="tableBorder">
  <tr class="tdbg"> 
    <td width="65" height="30"><strong>管理导航：</strong></td>
    <td height="30"><a href="Sk_ItemDatabase.asp">管理首页</a> | <a href="Sk_ItemDatabase.asp?Action=Compact">数据库压缩</a> | <a href="Sk_ItemDatabase.asp?Action=Backup">数据库备份</a> | <a href="Sk_ItemDatabase.asp?Action=Restore">数据库恢复</a> | <a href="Sk_ItemDatabase.asp?Action=LeadOut">项目导出</a></td>
  </tr>
</table>
<% 
if Action="Compact" or Action="CompactData" then
	call ShowCompact()
elseif Action="Backup" or Action="BackupData" then
        call ShowBackup()
elseif Action="Restore" or Action="RestoreData" then
        call ShowRestore()
elseif Action="LeadOut" or Action="LeadOutData" then 
	call ShowLeadOut()
elseif  Action="LeadIn" or Action="ShowLeadInData" or Action="LeadInData"  Then
    call  ShowLeadIn()
elseif  Action="ShowUpData" or Action="UpData"  Then
    call  ShowUpData()
Else
    call main()
End  If
call closeconnitem()
%>
<!--#include file="Admin_ItemFoot.asp"-->
</body>
</html>

<%Sub Main%>
<br>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="tableBorder">
  <tr>
	<td colspan="2" align="center" class="title" height=22><b>数据库管理说明</b></td>
	</tr>
  <tr>
	<td colspan="2" align="left" class="tdbg" height="100">
	<br>
	<p align="center">欢迎使用SK采集系统，使用数据库管理功能之前请阅读此说明。</p>
	<p>
	1<span lang="zh-cn">、数据库压缩：</span><p>
	&nbsp;&nbsp; <span lang="zh-cn">由于使用了历史记录，数据库的记录数会越来越多，使用压缩功能将会使数据库体积减少不少。</span><p>
	2<span lang="zh-cn">、数据库备份：</span><p>
	&nbsp;&nbsp; <span lang="zh-cn">备份数据以防意外。</span><p>
	3<span lang="zh-cn">、数据库恢复：</span><p>
	<span lang="en-us">&nbsp;&nbsp; </span>使用本功能可以恢复数据库，前提是有数据库备份。<p>
	4<span lang="zh-cn">、项目导出：</span><p>
	&nbsp;&nbsp; 
	是不是经常有朋友问这个怎么操作、那个怎么操作？虽然你很热情，但是久了也不能保证还有那份热情，别急，使用项目导出功能将项目数据导出到一个干净的数据库中，让你的朋友下载，然后使用项目导入功能，是不是什么事情都解决了。<p>
	<span lang="en-us">5</span>、项目导入：<p>
	<span lang="en-us">&nbsp;&nbsp; </span>和朋友交流本系统项目的设置心得，这可是少不了的哦。<p>
	<span lang="en-us">6</span>、检查更新数据：<p>
	<span lang="en-us">&nbsp;&nbsp; </span>在使用项目导入功能后必须使用本功能更新数据，否则不能正常采集。<p>
	　</td>
	</tr>
</table>
<%End Sub%>

<%Sub  ShowCompact
If  Action="Compact"  Then
%>
<br>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="tableBorder">
<form method="post" action="Sk_ItemDatabase.asp?Action=CompactData">
  <tr>
    <td colspan="2" align="center" class="title" height=22><b>数据库压缩</b></td>
  </tr>
  <tr class="tdbg">
    <td align="center" valign="middle" height="100">
      <br>
      <font color="#FF6600"><b>注：</b></font>压缩前，建议先备份数据库，以免发生意外错误。 <br>
    </td>
  </tr>
  <tr class="tdbg">
    <td align="center">
      <br>
      <input name="submit" type=submit value=" &nbsp;压缩数据库&nbsp; " <%If ObjInstalled=False Then response.Write "disabled"%> style="cursor: hand;background-color: #cccccc;">
   <%
   If ObjInstalled=False Then
      Response.Write "<br><b><font color=red>您的服务器不支持 FSO(Scripting.FileSystemObject)组件! 不能使用本功能</font></b>"
   End if
   %>
    </td>
  </tr>
</form>
</table>
<%
Else
   Call CompactData()
end if
%>
<%End  Sub%>

<%Sub ShowBackup
If Action="Backup" Then
%>
<br>
<form method="post" action="Sk_ItemDatabase.asp?Action=BackupData">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="tableBorder">
  <tr>
	<td colspan="2" align="center" class="title" height=22><b>数据库备份</b></td>
	</tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="tableBorder">
  <tr>
	<td width='200' height='33' align='right' class="tdbg">备份目录：</td>
	<td class="tdbg"><input type=text size=20 name="BackPath" value="Databackup"></td>
	<td class="tdbg">相对路径目录，如目录不存在，将自动创建</td>
  </tr>
  <tr>
	<td width='200' height='34' align='right' class="tdbg">备份名称：</td>
	<td height='34' class="tdbg"><input type=text size=20 name="BackMdb" value="<%=Date()%>"></td>
	<td height='34' class="tdbg">不用输入文件名后缀（默认为“.asa”）。如有同名文件，将覆盖</td>
  </tr>
  <tr align='center'>
        <td height='40' colspan='3' class="tdbg"><input name='submit' type=submit value=' 开始备份 ' <%If ObjInstalled=false Then response.Write "disabled"%>></td>
  </tr>
  <%If ObjInstalled=false Then
	Response.Write "<b><font color=red>你的服务器不支持 FSO(Scripting.FileSystemObject)组件! 不能使用本功能</font></b>"
  end if
  %>
</table>
</form>
<%Else
   Call BackUpData()
End If
End Sub%>

<%Sub ShowRestore
If Action="Restore" Then
%>
<br>
<form method="post" action="Sk_ItemDatabase.asp?Action=RestoreData">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="tableBorder">
  <tr>
	<td colspan="2" align="center" class="title" height=22><b>数据库恢复</b></td>
	</tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="tableBorder">
  <tr>
	<td height='100' align='center' class="tdbg">备份数据库路径（相对）：<input name="RestorePath" type=text id="RestorePath" value="Databackup\<%=Date()%>.asa" size=50 maxlength="200">
        </td>
  </tr>
  <tr>
        <td align='center' class="tdbg">
        <input name='submit' type=submit value=' 开始恢复 ' <%If ObjInstalled=false Then response.Write "disabled"%>>
  <%If ObjInstalled=false Then
	Response.Write "<b><font color=red>你的服务器不支持 FSO(Scripting.FileSystemObject)组件! 不能使用本功能</font></b>"
  end if
  %>
        </td>
  </tr>
</table>
</form>
<%Else
   Call RestoreData()
End If
End Sub%>


<%Sub  ShowLeadOut
If  Action="LeadOut"  Then
%>
<br>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="tableBorder">
  <tr>
     <td colspan="2" align="center" class="title" height=22><b>项目导出</b></td>
  </tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="tableBorder">
<form method="post" name="myform" action="Sk_ItemDatabase.asp?Action=LeadOutData">
   <tr class="tdbg">         
      <td width="5%" height="22" align="center" class=ButtonList>选择</td>                 
      <td width="10%" align="center" class=ButtonList>项目名称</td>
      <td width="10%" align="center" class=ButtonList>所属频道</td> 
      <td width="10%" align="center" class=ButtonList>所属栏目</td> 
      <td width="10%" align="center" class=ButtonList>所属专题</td>   
      <td width="5%" align="center" class=ButtonList>状态</td>      
   </tr>         
<%
   Set RsItem=server.createobject("adodb.recordset")         
   SqlItem="select ItemID,ItemName,ChannelID,ClassID,SpecialID,Flag from Item order by ItemID ASC"         
   RsItem.open SqlItem,ConnItem,1,1
   If (Not RsItem.Eof)  And (Not RsItem.Bof) then
%>
   <%
      Do While Not RsItem.Eof
   %>
   <tr class="tdbg">         
      <td width="5%" height="22" align="center"><input type="checkbox" value=<%=RsItem("ItemID")%> name="ItemID" onClick="unselectall(this.form)" style="border: 0px;background-color: #E1F4EE;"></td>                 
      <td width="10%" align="left"><%=RsItem("ItemName")%></td>
      <td width="10%" height="22" align="center"><%Call Admin_ShowChannel_Name(RsItem("ChannelID"))%></td> 
      <td width="10%" align="center"><%Call Admin_ShowClass_Name(RsItem("ChannelID"),RsItem("ClassID"))%></td>   
      <td width="10%" align="center"><%Call Admin_ShowSpecial_Name(RsItem("ChannelID"),RsItem("SpecialID"))%></td>   
      <td width="5%" align="center">
      <%
         If RsItem("Flag")=True Then
            Response.write "√"
         Else
            Response.write "<font color=red>×</font>"
         End If
      %>
      </td>      
   </tr>   
      <%
         RsItem.MoveNext
      Loop
      %>
   <tr class="tdbg">
      <td colspan=7 height="52" align="center">
      <input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox" >全选&nbsp;&nbsp;&nbsp;&nbsp;
      导出到数据库：<input type="text" name="LeadOutMdb" size="30" value="Database/sk_LeadOut.mdb">
      <input type="submit" name="submit" value="导出" style="cursor: hand;background-color: #cccccc;">
      </td>
   </tr>
<%
   Else
%>
      <tr class="tdbg">
        <td colspan='9' class="tdbg" align="center"><br>系统中暂无采集项目！</td>
      </tr>
<%
      End If
   RsItem.Close
   Set RsItem=Nothing
%>
</form>
</table>
<%
Else
   Call LeadOutData()
End If
%>
<%End  Sub%>

<%Sub  ShowLeadIn
If  Action="LeadIn" Then
%>
<br>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="tableBorder">
  <tr>
	<td colspan="2" align="center" class="title" height=22><b>项目导入</b></td>
	</tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="tableBorder">
<form method="post" action="Sk_ItemDatabase.asp?Action=ShowLeadInData">
   <tr class="tdbg">
      <td align="center" valign="middle" height="100">
        <br>
	数据库位置：
	<input name="LeadInMdb" type="text" id="LeadInMdb" size="23" value="Databackup/LeadIn.mdb">
      </td>
   </tr>
   <tr class="tdbg">
      <td align="center">
        <input name="submit" type=submit value=" 下&nbsp;一&nbsp;步 " style="cursor: hand;background-color: #cccccc;">
      </td>
   </tr>
</form>
</table>
<%
ElseIf Action="ShowLeadInData" Then
   Call ShowLeadInData()
Else
   Call LeadInData()
End if
End  Sub%>

<%Sub ShowUpData
If Action="ShowUpData" Then
%>
<br>
<form method="post" action="Sk_ItemDatabase.asp?Action=UpData">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="tableBorder">
  <tr>
	<td colspan="2" align="center" class="title" height=22><b>检查更新数据</b></td>
	</tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="tableBorder">
  <tr>
	<td height="34" align='left' class="tdbg">请选择要更新的数据</td>
  </tr>
  <tr>
	<td align='left' class="tdbg">
        <input type="checkbox" name="ChannelData" value="yes">频道部分数据&nbsp;&nbsp;
        <input type="checkbox" name="ItemData" value="yes" checked disabled>项目数据
        </td>
  </tr>
  <tr align='center'>
        <td height='40' colspan='3' class="tdbg"><input name='submit' type=submit value=' 开始更新 '></td>
  </tr>
</table>
</form>
<%Else
   Call UpData()
End If
End Sub%>

<%
sub CompactData()
        '关闭数据库链接
        Call CloseConnItem() 
	Dim fso, Engine,strDBPath,DBPath,DbTemp
	DBPath = server.mappath(DbItem)'数据库文件
        Randomize timer
	DbTemp = CStr(Clng(8999999*Rnd+1000000))
	if instr(DBPath,"/") then 
		strDBPath = left(DBPath,instrrev(DBPath,"/"))
	else
		strDBPath = left(DBPath,instrrev(DBPath,"\"))
	end if
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(DBPath) Then
		Set Engine = CreateObject("JRO.JetEngine")
		Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBPath," Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & DbTemp & ".asa"
		fso.CopyFile strDBPath & DbTemp & ".asa",DBPath
		fso.DeleteFile(strDBPath & DbTemp & ".asa")
		ErrMsg="<br>数据库压缩成功!"
	Else
                FoundErr=True
		ErrMsg="<br><li>数据库没有找到!</li>"
	End If
	Set fso = nothing
	Set Engine = nothing
        If FoundErr=True Then
           Call WriteErrMsg(ErrMsg)
        else
           Call WriteSucced(ErrMsg)
        End If
end sub

sub BackUpData()
	Dim fso,BackPath,BackMdb
        BackPath=Trim(Request("BackPath"))
        BackMdb=Trim(Request("BackMdb"))
        If BackPath="" Then
           FoundErr=True
           ErrMsg="<br><li>请指定备份目录！</li>"
        else
           BackPath=Replace(BackPath," ","")
        End If
        If BackMdb="" Then
           FoundErr=True
           ErrMsg=ErrMsg & "<br><li>请指定备份文件名</li>"
        Else
           BackMdb=Replace(BackMdb," ","")
        End If
        If FoundErr<>True Then
           Set fso = Server.CreateObject("Scripting.FileSystemObject")
	   If fso.FolderExists(server.mappath(BackPath))=False Then
              fso.CreateFolder(server.mappath(BackPath))
           End If
           If fso.FileExists(server.mappath(BackPath & "/" & BackMdb & ".asa"))=True then
              fso.DeleteFile(server.mappath(BackPath & "/" & BackMdb & ".asa"))
           End If
           fso.copyfile server.mappath(DbItem),server.mappath(BackPath & "/" & BackMdb & ".asa")
	   If fso.FileExists(server.mappath(BackPath & "/" & BackMdb & ".asa"))=True Then
              ErrMsg="<br>数据库备份成功!"
              ErrMsg=ErrMsg & "<br>数据库备份为：" & BackPath & "/" & BackMdb & ".asa"
           Else
              FoundErr=True
              ErrMsg="<br><li>数据库备份失败!</li>"
	   End If
	   Set fso = nothing
	End If
        If FoundErr=True Then
           Call WriteErrMsg(ErrMsg)
        Else
           Call WriteSucced(ErrMsg)
        End If
end sub


sub RestoreData()
	Dim fso,RestorePath
        RestorePath=Trim(Request("RestorePath"))
        If RestorePath="" Then
           FoundErr=True
           ErrMsg="<br><li>请指定原备份的数据库文件名！</li>"
        else
           RestorePath=Replace(RestorePath," ","")
        End If

        If FoundErr<>True Then
           Set fso = Server.CreateObject("Scripting.FileSystemObject")
           If fso.FileExists(server.mappath(RestorePath))=True then
              fso.copyfile server.mappath(RestorePath),server.mappath(DbItem)
              ErrMsg="<br>数据库恢复成功!"
           Else
              FoundErr=True
              ErrMsg=ErrMsg & "<br><li>数据库：" & RestorePath & " 不存在！"
           End If
	   Set fso = nothing
	End If
        If FoundErr=True Then
           Call WriteErrMsg(ErrMsg)
        Else
           Call WriteSucced(ErrMsg)
        End If
end sub

Sub LeadOutData
   Dim fso,ItemMdb,ItemMdbPath,LeadOutMdb,RsF,SqlF,RsLead,SqlLead,ItemIDTemp
   LeadOutMdb=trim(request.form("LeadOutMdb"))
   ItemID=trim(request.form("ItemID"))
   ItemMdb=DbItem
   ItemMdbPath=Left(DbItem,Instrrev(DbItem,"/")-1)
   If Instr(ItemMdb,"/")>0 Then
      ItemMdbPath=Left(ItemMdb,InstrRev(ItemMdb,"/"))
   End If
   If LeadOutMdb="" then
      FoundErr=True
      ErrMsg="<br><li>数据库地址不能为空！</li>"
   End If
   If ItemID="" Then
      FoundErr=True
      ErrMsg=ErrMsg & "<br><li>请选择要导出的项目</li>"
   Else
      ItemID=Replace(ItemID," ","")
   End If
   If FoundErr<>True And ObjInstalled<>False Then
      Set fso = Server.CreateObject("Scripting.FileSystemObject")
      If fso.FileExists(Server.MapPath(LeadOutMdb)) Then
      Else
         '不存在则创建
         If fso.FileExists(Server.MapPath(ItemMdbPath & "ItemTemp.mdb")) Then
            fso.CopyFile Server.MapPath(ItemMdbPath & "ItemTemp.mdb"),Server.MapPath(LeadOutMdb)
         Else
            FoundErr=True
            ErrMsg=ErrMsg& "<br>用于导出项目的数据库：ItemTemp.mdb不存在！"
         End If
      End If
      set fso=nothing
   End If

   If FoundErr<>True Then
      dim connstrLead,connLead
      Set connLead = Server.CreateObject("ADODB.Connection")
      connstrLead="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(LeadOutMdb)
      connLead.Open connstrLead
      If Err Then
         err.Clear
         FoundErr=True
	 ErrMsg=ErrMsg & "<br>数据库连接出错，请确认数据库是否存在。"
      End If

      If FoundErr<>True Then
         ConnLead.execute("Delete From Item")
         ConnLead.execute("Delete From Filters")
         Set RsItem=server.createobject("adodb.recordset")         
         SqlItem="select * from Item where ItemID in(" & ItemID & ") order by ItemID DESC"         
         RsItem.open SqlItem,ConnItem,1,1
         If Not RsItem.Eof then
            Do while Not RsItem.Eof
               '打开数据库
               Set RsLead=server.createobject("adodb.recordset")         
               SqlLead="select * from Item"         
               RsLead.open SqlLead,ConnLead,1,3
               RsLead.AddNew
               for i=1 to Ubound(copyitem)
			   RsLead(i)=RsItem(i)
			   next
               RsLead.Update
               RsLead.Close
               Set RsLead=Nothing

               '过滤信息
               Set RsF=server.createobject("adodb.recordset")         
               SqlF="select * from Filters Where ItemID=" & RsItem("ItemID") & " order by ItemID DESC"         
               RsF.open SqlF,ConnItem,1,1              
               If Not RsF.Eof then
                  Do While Not RsF.Eof
                     Set RsLead=server.createobject("adodb.recordset")         
                     SqlLead="select * from Filters"         
                     RsLead.open SqlLead,ConnLead,1,3
                     RsLead.AddNew
                     RsLead("ItemID")=ItemIDTemp
                     RsLead("FilterName")=RsF("FilterName")
                     RsLead("FilterObject")=RsF("FilterObject")
                     RsLead("FilterType")=RsF("FilterType")
                     RsLead("FilterContent")=RsF("FilterContent")
                     RsLead("FisString")=RsF("FisString")
                     RsLead("FioString")=RsF("FioString")
                     RsLead("FilterRep")=RsF("FilterRep")
                     RsLead("Flag")=RsF("Flag")
                     RsLead("PublicTf")=RsF("PublicTf")
                     RsLead.Update
                     RsLead.Close
                     Set RsLead=Nothing
                  RsF.MoveNext
                  Loop
               End If
               RsF.Close
               Set RsF=Nothing

            RsItem.MoveNext
            Loop
         End If
         RsItem.Close
         Set RsItem=Nothing
      End If
      ConnLead.close
      set connlead=nothing

   End If
   If FoundErr<>True Then
      ErrMsg="<br>数据导出成功"
      ErrMsg=ErrMsg & "<br>数据导出为：" & LeadOutMdb
      Call WriteSucced(ErrMsg)
   Else
      Call WriteErrMsg(ErrMsg)
   End If
End Sub

Sub ShowLeadInData
   Dim LeadInMdb,connstrLead,connLead,RsLead,SqlLead
   LeadInMdb=Trim(Request("LeadInMdb"))
   If LeadInMdb="" Then
      FoundErr=True 
      ErrMsg="<br><li>数据库地址不能为空！</li>"
   End If
   If FoundErr<>True Then
      On error resume next
      Set connLead = Server.CreateObject("ADODB.Connection")
      connstrLead="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(LeadInMdb)
      connLead.Open connstrLead
      If Err Then
         err.Clear
         FoundErr=True
         ErrMsg=ErrMsg & "<br><li>数据库连接出错，请确认数据库是否存在。</li>"
      End If
      If FoundErr<>True Then
         Set RsLead=server.createobject("adodb.recordset")         
         SqlLead="select ItemID,ItemName,ChannelID,ClassID,SpecialID,Flag from Item order by ItemID DESC"         
         RsLead.open SqlLead,ConnLead,1,1
         If Not RsLead.Eof then
%>
<br>
<form method="post" action="Sk_ItemDatabase.asp?Action=LeadInData">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="tableBorder">
  <tr>
	<td colspan="2" align="center" class="title" height=22><b>项目导入</b></td>
	</tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="tableBorder">
   <tr class="tdbg">         
      <td width="5%" height="22" align="center" class=ButtonList>选择</td>                 
      <td width="10%" align="center" class=ButtonList>项目名称</td>
      <td width="10%" height="22" align="center" class=ButtonList>所属频道</td> 
      <td width="10%" height="22" align="center" class=ButtonList>所属栏目</td> 
      <td width="10%" align="center" class=ButtonList>所属专题</td>   
      <td width="5%" align="center" class=ButtonList>状态</td>      
   </tr>         
<%
            Do While Not RsLead.Eof
%>
   <tr class="tdbg">         
      <td width="5%" height="22" align="center"><input type="checkbox" value=<%=RsLead("ItemID")%> name="ItemID" onClick="unselectall(this.form)" style="border: 0px;background-color: #E1F4EE;"></td>                 
      <td width="10%" align="left"><%=RsLead("ItemName")%></td>
      <td width="10%" height="22" align="center"><%Call Admin_ShowChannel_Name(RsLead("ChannelID"))%></td> 
      <td width="10%" height="22" align="center"><%Call Admin_ShowClass_Name(RsLead("ChannelID"),RsLead("ClassID"))%></td> 
      <td width="10%" align="center"><%Call Admin_ShowSpecial_Name(RsLead("ChannelID"),RsLead("SpecialID"))%></td>
      <td width="5%" align="center">
      <%If RsLead("Flag")=True Then
           Response.write "√"
        Else
           Response.Write "<font color=red>×</font>"
      End If%>
      </td>            
   </tr>   
<%
            RsLead.MoveNext
            Loop
%>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="tableBorder">
   <tr class="tdbg">
     <td align="center">
	<input name="LeadInMdb" type="hidden" value="<%=LeadInMdb%>">
        <input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox" >全选&nbsp;&nbsp;&nbsp;&nbsp;
	<input name="step" type="hidden" value="1">
	<input name="submit" type=submit value=" 确&nbsp;&nbsp;&nbsp;定 " style="cursor: hand;background-color: #cccccc;">
     </td>
   </tr>
</table>
</form>
<%
         Else
            FoundErr=True
            Errmsg=ErrMsg &  "<br>无任何记录！"
         End If
         RsLead.Close
         Set RsLead=Nothing
      End If
      connLead.close
      set connlead=nothing
   End If
   If FoundErr=True Then
      Call WriteErrMsg(ErrMsg)
   End If
End Sub

Sub UpData()
'要更新的内容：
'频道目录
'保存图片
'
'频道数据（还少了专题）
   Dim rsCount,sqlCount,aCount,bCount,Arr_Channel,i_Channel,sqlItem
   Set Rs=ConnItem.execute("Select ChannelID,ChannelDir,ChannelName from PE_Channel Where ModuleType=1")
   If Not Rs.Eof Then
      Arr_Channel=Rs.GetRows()
   End If
   Set Rs=Nothing
   If IsArray(Arr_Channel)=True Then
      For i_Channel=0 To Ubound(Arr_Channel,2)
         If Trim(Request("ChannelData"))="yes" Then
            Set rsCount= Server.CreateObject("ADODB.Recordset")
            sqlCount="select count(ArticleID) from PE_Article where ChannelID=" & Arr_Channel(0,i_Channel) & " And Passed=-1 and deleted=0"
            rsCount.open sqlCount,ConnItem,1,1
            If rsCount.Eof Then
               aCount=0
            Else
               aCount=rsCount(0)
            End If
            rsCount.Close

            sqlCount="select count(ArticleID) from PE_Article where ChannelID=" & Arr_Channel(0,i_Channel) & " And Passed=0 and deleted=0"
            rsCount.open sqlCount,conn,1,1
            If rsCount.Eof Then
               bCount=0
            Else
               bCount=rsCount(0)
            End If
            rsCount.Close
            set rsCount=Nothing
            ConnItem.execute("Update [PE_Channel] Set ItemCount=" & aCount+bCount & " where ChannelID=" & Arr_Channel(0,i_Channel))
            ErrMsg=ErrMsg & "<br><b>" & Arr_Channel(2,i_Channel) & "</b> 文章总数：" & aCount+bCount & " 已审核数：" & aCount & " 未审核数：" & bCount
         End If
         SqlItem="Update [Item] Set ChannelDir='" & Arr_Channel(1,i_Channel) & "'"
         If ObjInstalled=False Then
            SqlItem=SqlItem & ",SaveFiles=False"
         End If
         SqlItem=SqlItem & " where ChannelID=" & Arr_Channel(0,i_Channel)
         ConnItem.Execute(SqlItem)
      Next
      If Request("ChannelData")="yes" Then
         ErrMsg=ErrMsg & "<br>频道数据更新完毕"
      End If
   End If
'项目数据(未完成)
   ErrMsg=ErrMsg & "<br>检查项目数据"
   Set RsItem=server.createobject("adodb.recordset")         
   SqlItem="select * from Item"
   RsItem.open SqlItem,ConnItem,1,1    
   If Not RsItem.Eof Then
      Do While (Not RsItem.Eof) and (Not RsItem.Bof)
         FoundErr=False
         ErrMsg=ErrMsg & "<br><b>" & RsItem("ItemName") & "</b> 项目数据： "
         If RsItem("ItemName")="" or isnull(RsItem("ItemName")) Then
            FoundErr=True
            ErrMsg=ErrMsg & "项目名称"
         End If
         If RsItem("ChannelID")="" or RsItem("ChannelID")=0 or IsNull(RsItem("ChannelID")) Then
            FoundErr=True
            ErrMsg=ErrMsg & " 频道"
         else
            If RsItem("ClassID")="" or RsItem("ClassID")=0 Or IsNull(RsItem("ClassID")) Then
               FoundErr=True
               ErrMsg=ErrMsg & " 栏目"
            Else
               set tClass=ConnItem.execute("select C.Child,C.LinkUrl From PE_Class C inner join PE_Channel D on C.ChannelID=D.ChannelID Where C.ChannelID="  & RsItem("ChannelID") & " and C.ClassID=" & RsItem("ClassID"))
               If tClass.bof and tClass.eof then
                  FoundErr=True
                  ErrMsg=ErrMsg & " 栏目"
               Else
                  if tClass(0)>0 then
                     FoundErr=True
                     ErrMsg=ErrMsg & " 栏目"
                  End if
                  If tClass(1)<>"" then
                     FoundErr=True
                     ErrMsg=ErrMsg & " 栏目"
                  End if
               End If
               Set tClass=Nothing
            End If
            If IsNumeric(RsItem("SpecialID"))=False Then
               FoundErr=True
               ErrMsg=ErrMsg & " 专题"
            Else
               If RsItem("SpecialID")<>0 Then
                  set tSpecial=ConnItem.execute("select SpecialID From PE_Special Where ChannelID="  & RsItem("ChannelID"))
                  If tSpecial.bof and tSpecial.eof then
                     FoundErr=True
                     ErrMsg=ErrMsg & " 专题"
                  End If
                  Set tSpecial=Nothing
               End If
            End If
         End If
         If RsItem("WebName")="" or IsNull(RsItem("WebName")) Then
            FoundErr=True
            ErrMsg=ErrMsg & " 网站名称"
         End If
         If RsItem("WebUrl")="" Or IsNull(RsItem("WebUrl")) Then
            FoundErr=True
            ErrMsg=ErrMsg & "、网站地址"
         End If
         If RsItem("LoginType")="" or IsNull(RsItem("LoginType")) Then
            FoundErr=True
            ErrMsg=ErrMsg & " 网站登录类型"
         else
            If RsItem("LoginType")=1 Then
               If RsItem("LoginUrl")="" or RsItem("LoginPostUrl")="" or Instr(RsItem("LoginUser"),"=")=0 or Instr(RsItem("LoginPass"),"=")=0 or RsItem("LoginFalse")="" then
                  FoundErr=True
                  ErrMsg=ErrMsg & " 网站登录参数"
               End If
            End If
         End If
         If RsItem("ListPaingType")="" or IsNull(RsItem("ListPaingType")) Then
            FoundErr=True
            ErrMsg=ErrMsg & " 列表分页类型"
         Else
            If RsItem("ListPaingType")=0 Then
               If RsItem("ListStr")="" or IsNull(RsItem("ListStr"))=True Then
                  FoundErr=True
                  ErrMsg=ErrMsg & " 列表索引"
               End If
            ElseIf RsItem("ListPaingType")=1 Then
               If RsItem("ListStr")="" Then
                  FoundErr=True
                  ErrMsg=ErrMsg & " 列表索引"
               End If
               If RsItem("LPsString")="" or IsNull(RsItem("LPsString")) or RsItem("LPoString")="" Or IsNull(RsItem("LPoString")) then
                  FoundErr=True
                  ErrMsg=ErrMsg & " 列表分页标记"
               end If
               If IsNull(RsItem("ListPaingStr1"))<>True and Len(RsItem("ListPaingStr1"))<15 Then
                  FoundErr=True
                  ErrMsg=ErrMsg & " 索引分页重定向"
               End If
            ElseIf RsItem("ListPaingType")=2 Then
               If Len(RsItem("ListPaingStr2"))<15 or Instr(RsItem("ListPaingStr2"),"{$ID}")=0 Then
                  FoundErr=True
                  ErrMsg=ErrMsg & " 列表原始字符"
               End If
               If IsNumeric(RsItem("ListPaingID1"))=False or IsNumeric(RsItem("ListPaingID2"))=False or (RsItem("ListPaingID1")=0 and RsItem("ListPaingID2")=0) Then
                  FoundErr=True
                  ErrMsg=ErrMsg & " 列表分页范围"
               End If
            ElseIf RsItem("ListPaingType")=3 Then
               If RsItem("ListPaingStr3")="" Then
                  FoundErr=True
                  ErrMsg=ErrMsg & " 列表手动分页标记"
               End If
            Else
               FoundErr=True
               ErrMsg=ErrMsg & " 列表分页类型"
            End If
         End If
         If RsItem("LsString")="" or IsNull(RsItem("LsString")) Or RsItem("LoString")="" Or IsNull(RsItem("LoString")) Then
            FoundErr=True
            ErrMsg=ErrMsg & " 列表标记设置"
         End If
         If RsItem("HsString")="" or IsNull(RsItem("HsString")) Or RsItem("HoString")="" Or IsNull(RsItem("HoString")) Then
            FoundErr=True
            ErrMsg=ErrMsg & " 链接标记设置"
         End If
         If IsNull(RsItem("HttpUrlType")) Then
            FoundErr=True
            ErrMsg=ErrMsg & " 链接类型"
         Else
            If RsItem("HttpUrlType")=1 and Len(Rsitem("HttpUrlStr"))<15 Then
               FoundErr=True
               ErrMsg=ErrMsg & " 链接字符"
            End If
         End If
         If RsItem("TsString")="" or IsNull(RsItem("TsString")) Or RsItem("ToString")="" Or IsNull(RsItem("ToString")) Then
            FoundErr=True
            ErrMsg=ErrMsg & " 标题标记设置"
         End If
         If RsItem("CsString")="" or IsNull(RsItem("CsString")) Or RsItem("CoString")="" Or IsNull(RsItem("CoString")) Then
            FoundErr=True
            ErrMsg=ErrMsg & " 正文标记设置"
         End If
         If IsNull(RsItem("DateType")) Then
            FoundErr=True
            ErrMsg=ErrMsg & " 时间标记设置"
         Else
            If RsItem("DateType")=1 Then
               If RsItem("DsString")="" or IsNull(RsItem("DsString")) Or RsItem("DoString")="" Or IsNull(RsItem("DoString")) Then
                  FoundErr=True
                  ErrMsg=ErrMsg & " 时间标记设置"
               End If
            End If
         End If
         If IsNull(RsItem("AuthorType")) Then
            FoundErr=True
            ErrMsg=ErrMsg & " 作者标记设置"
         Else
            If RsItem("AuthorType")=1 Then
               If RsItem("AsString")="" or IsNull(RsItem("AsString")) Or RsItem("AoString")="" Or IsNull(RsItem("AoString")) Then
                  FoundErr=True
                  ErrMsg=ErrMsg & " 作者标记设置"
               End If
            ElseIf RsItem("AuthorType")=2 Then
               If RsItem("AuthorStr")="" or IsNull(RsItem("AuthorStr")) Then
                  FoundErr=True
                  ErrMsg=ErrMsg & " 指定作者设置"
               End If 
            End If
         End If
         If RsItem("CopyFromType")=1 Then
            If RsItem("FsString")="" or RsItem("FoString")="" Then
               FoundErr=True
               ErrMsg=ErrMsg & " 来源标记设置"
            End If
         ElseIf RsItem("CopyFromType")=2 Then
            If RsItem("CopyFromStr")="" Then
               FoundErr=True
               ErrMsg=ErrMsg & " 指定来源设置"
            End If 
         End If
         If RsItem("KeyType")=1 Then
            If RsItem("KsString")="" or RsItem("KoString")="" Then
               FoundErr=True
               ErrMsg=ErrMsg & " 关键字标记设置"
            End If
         ElseIf RsItem("KeyType")=2 Then
            If RsItem("KeyStr")="" Then
               FoundErr=True
               ErrMsg=ErrMsg & " 指定关键字设置"
            End If 
         End If
         If RsItem("NewsPaingType")=1 Then
            If RsItem("NPsString")="" or RsItem("NPoString")="" Then
               FoundErr=True
               ErrMsg=ErrMsg & " 正文分页标记设置"
            End If
            If RsItem("NewsPaingStr")<>"" And Len(RsItem("NewsPaingStr"))<15 Then
               FoundErr=True
               ErrMsg=ErrMsg & " 正文分页绝对链接"
            End If
         ElseIf RsItem("NewsPaingType")=2 Then
            FoundErr=True
            ErrMsg=ErrMsg & " 正文分页类型"
         End If
         If RsItem("PaginationType")=1 Then
            If RsItem("MaxCharPerPage")=0 Then
               FoundErr=True
               ErrMsg=ErrMsg & " 自动分页的每页字符数"
            End If
         End If
         If RsItem("ReadLevel")<>5 And RsItem("ReadLevel")<>9 And RsItem("ReadLevel")<>99 And RsItem("ReadLevel")<>999 And RsItem("ReadLevel")<>9999 Then
            FoundErr=True
            ErrMsg=ErrMsg & " 文章阅读等级"
         End If
         If RsItem("Stars")<>0 And RsItem("Stars")<>1 And RsItem("Stars")<>2 And RsItem("Stars")<>3 And RsItem("Stars")<>4 And RsItem("Stars")<>5 Then
            FoundErr=True
            ErrMsg=ErrMsg & " 文章评分等级"
         End if
         If IsNumeric(Rsitem("ReadPoint"))=False Then
            FoundErr=True
            ErrMsg=ErrMsg & " 文章阅读点数"
         end If
         If IsNumeric(Rsitem("Hits"))=False Then
            FoundErr=True
            ErrMsg=ErrMsg & " 点击数初始值"
         end If
         If RsItem("UpDateType")=2 Then
            If IsDate(RsItem("UpDateTime"))=False Then
               FoundErr=True
               ErrMsg=ErrMsg & " 自定义时间"
            End If
         End If
         If RsItem("InputerType")=1 Then
            If RsItem("Inputer")="" Then
               FoundErr=True
               ErrMsg=ErrMsg & " 自定义录入者"
            End If
         End If
         If RsItem("EditorType")=1 Then
            If RsItem("Editor")="" Then
               FoundErr=True
               ErrMsg=ErrMsg & " 自定义编辑"
            End If
         End If
         If IsNumeric(RsItem("SkinID"))=False Then
            FoundErr=True
            ErrMsg=ErrMsg & " 配色风格"
         End If
         If IsNumeric(RsItem("TemplateID"))=False Then
            FoundErr=True
            ErrMsg=ErrMsg & " 设计模板"
         End If
         If IsNumeric(RsItem("CollecListNum"))=False Then
            FoundErr=True
            ErrMsg=ErrMsg & " 列表深度"
         End If
         If IsNumeric(RsItem("CollecNewsNum"))=False Then
            FoundErr=True
            ErrMsg=ErrMsg & " 新闻数量"
         End If
         If FoundErr=False Then
            ErrMsg=ErrMsg & " 状态--正常"
         Else
            ErrMsg=ErrMsg & " 设置不正确 状态--<font color=red>异常</font>"
         End If
         foundErr=False
         RsItem.movenext
      Loop
   end if
   rsItem.close
   set rsItem=nothing
   Call WriteSucced(ErrMsg)
End Sub
%>
