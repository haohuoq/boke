<!--#include file="conn.asp"-->
<!--#include file="Exif.asp"-->
<%
dim rs
Set rs=Server.CreateObject("ADODB.Recordset")
Function Format_Time(s_Time, n_Flag) 
Dim y, m, d, h, mi, s 
Format_Time = "" 
If IsDate(s_Time) = False Then Exit Function 
y = cstr(year(s_Time)) 
m = cstr(month(s_Time)) 
If len(m) = 1 Then m = "0" & m 
d = cstr(day(s_Time)) 
If len(d) = 1 Then d = "0" & d 
h = cstr(hour(s_Time)) 
If len(h) = 1 Then h = "0" & h 
mi = cstr(minute(s_Time)) 
If len(mi) = 1 Then mi = "0" & mi 
s = cstr(second(s_Time)) 
If len(s) = 1 Then s = "0" & s 
Select Case n_Flag 
Case 1 
' yyyy-mm-dd hh:mm:ss 
Format_Time = y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s 
Case 2 
' yyyy-mm-dd 
Format_Time = y & "-" & m & "-" & d 
Case 3 
' hh:mm:ss 
Format_Time = h & ":" & mi & ":" & s 
Case 4 
' yyyy年mm月dd日 
Format_Time = y & "年" & m & "月" & d & "日" 
Case 5 
' yyyymmdd 
Format_Time = y & m & d 
case 6 
'yyyymmddhhmmss 
format_time= y & m & d & h & mi & s 
End Select 
End Function 

function textToHtml(str)
    dim result
    dim l
    if isNULL(str) then 
       textToHtml=""
       exit function
    end if
    l=len(str)
    result=""
 dim i
 for i = 1 to l
     select case mid(str,i,1)
            case "'" '省略号
                 result=result+"&#39"
            case "<" '小于号
                 result=result+"&lt;"
            case ">" '大于号
                 result=result+"&gt;"
              case chr(13) '新行符
                 result=result+"<br>"
            case chr(34) '引号
                 result=result+"&quot;"
            case "&" '与号
                 result=result+"&amp;" 
              case chr(32) '不换行空格号           
                 'result=result+"&nbsp;"
                 if i+1<=l and i-1>0 then
                    if mid(str,i+1,1)=chr(32) or mid(str,i+1,1)=chr(9) or mid(str,i-1,1)=chr(32) or mid(str,i-1,1)=chr(9)  then                       
                       result=result+"&nbsp;"
                    else
                       result=result+" "
                    end if
                 else
                    result=result+"&nbsp;"                     
                 end if
            case chr(9) '跳格符
                 result=result+"    "
            case else
                 result=result+mid(str,i,1)
         end select
       next 
       textToHtml=result
   end function



'更具id查找相片并更新浏览次数，禁止刷新浏览次数
Dim xcid,xcclass
xcid = unescape(Request.QueryString("xcid"))
xcclass = unescape(Request.QueryString("xcclass"))
SQL="select * from blog_album where album_class="&xcclass&" and album_ID="&xcid
rs.open sql,conn,1,3
 
if session("ip")="" and session("xcid")="" then 
rs("Countnum")=rs("Countnum")+1  
rs.update
Session("ip")=Request.ServerVariables("REMOTE_ADDR")
session("xcid")=xcid
elseif session("ip")=Request.ServerVariables("REMOTE_ADDR") and session("xcid")<>xcid then
rs("Countnum")=rs("Countnum")+1  
rs.update
session("xcid")=xcid
end if
' 查找上一张 和下一张的图片id
dim prers,nextrs,prep,nextp
set prers=server.CreateObject("adodb.recordset")
sql="select top 1 album_ID from blog_album where album_class="&xcclass&" and album_ID> "&xcid&" order by album_ID asc"
prers.Open SQL,conn,3,1
if prers.eof then
prep=0
else
prep=prers("album_ID")
end if

set nextrs=server.CreateObject("adodb.recordset")
sql="select top 1 album_ID from blog_album where album_class="&xcclass&" and album_ID< "&xcid&" order by album_ID desc"
nextrs.Open SQL,conn,3,1
if nextrs.eof then
nextp=0
else
nextp=nextrs("album_ID")
end if


'统计分类相册里有多少张图片并查找当前图片位置
Dim tjclass,i
set tjclass=Server.CreateObject("Adodb.Recordset")
sql="select album_ID from blog_album where album_class="&xcclass&" order by album_ID desc"
tjclass.Open SQL,conn,3,1
i=1
do while not tjclass.eof
if tjclass("album_ID") = rs("album_ID") then
exit do
end if
tjclass.movenext
i=i+1
loop
dim pid,ptitle,pinfo,purl,pclass,ptime,phidden,pnum,ptj,pcount,pcomm
pid=rs("album_ID")
ptitle=rs("album_Title")
pinfo=rs("album_Messager")
purl=rs("album_Url")
pclass=rs("album_class")
ptime=format_time(rs("album_Time"),2)
pcount=rs("countnum")
if rs("album_Hidden") then phidden=0 else phidden=1 end if
if rs("album_Comment") then pcom=0 else pcom=1 end if
ptj=i
pnum= tjclass.recordcount

''''''''''''''''''''''
Dim photocom,sj,nei,commid
set photocom=Server.CreateObject("Adodb.Recordset")
sql="select * from blog_album_Comment where album_ID= "&pid&" order by album_CommentID asc"
photocom.Open SQL,conn,1,1


Dim str ,str2
If photocom.eof Then
str=""
else
do until photocom.eof
sj=format_time(photocom("album_CommentTIME"),2)
nei=textToHtml(photocom("album_CommentMessager"))
commid=photocom("album_CommentID")
str=str & "<div class='comment' style='width:90%'><DIV class=commenttop><A name=comm_"&photocom("album_ID")&" href='javascript:addQuote("""&photocom("album_CommentUSER")&""",""commcontent_"&photocom("album_ID")&""")'><IMG style='MARGIN: 0px 4px -3px 0px' border=0 alt='' src='images/icon_quote.gif'></A><A href='member.asp?action=view&amp;memName="&photocom("album_CommentUSER")&"'><STRONG>"&photocom("album_CommentUSER")&"</STRONG></A> <SPAN class=commentinfo>[ "&sj&" ]</SPAN></DIV><DIV id='commcontent_"&commid&"' class='commentcontent'>"&nei&" </DIV></div>"	

str2=str2 & "<div class='comment' style='width:90%'><DIV class=commenttop><A name=comm_"&photocom("album_ID")&" href='javascript:addQuote("""&photocom("album_CommentUSER")&""",""commcontent_"&photocom("album_ID")&""")'><IMG style='MARGIN: 0px 4px -3px 0px' border=0 alt='' src='images/icon_quote.gif'></A><A href='member.asp?action=view&amp;memName="&photocom("album_CommentUSER")&"'><STRONG>"&photocom("album_CommentUSER")&"</STRONG></A> <SPAN class=commentinfo>[ "&sj&" | "&photocom("album_CommentIP")&" | <a href='plugins/album/albumaction.asp?action=delcomms&amp;commID="&commid&"' onclick='if (!window.confirm(""Delete ?"")) {return false}'><img src='images/del1.gif' alt='Delete' border='0'/></a> ]</SPAN></DIV><DIV id='commcontent_"&commid&"' class='commentcontent'>"&nei&" </DIV></div>"
photocom.movenext
loop
str=escape(str)
str2=escape(str2)
end if
photocom.Close
Set photocom=Nothing
	

'''''''''''''''''''''''

dim F_Name,Exif,ExifSplit,noexif
dim CameraMake,CameraModel,DateTime,ImageDimension,Software,ISOSpeed,FStop,ExposureTime,Flash,ExposureBias,FocalLength,MeteringMode
F_Name="/"&purl&""
Exif=GetImageExifInfo(F_Name)
    if trim(Exif) <> "" then
        ExifSplit=Split(Exif,"|")
        CameraMake=ExifSplit(0)
        CameraModel=ExifSplit(1)
        DateTime=ExifSplit(2)
        ImageDimension=ExifSplit(3)
        Software=ExifSplit(4)
        ISOSpeed=ExifSplit(5)
        FStop=ExifSplit(6)
        ExposureTime=ExifSplit(7)
        Flash=ExifSplit(8)
        ExposureBias=ExifSplit(9)
        FocalLength=ExifSplit(10)
        MeteringMode=ExifSplit(11)
noexif=1
response.write (escape("{pid:'"&pid&"' , ptitle : '"&ptitle&"' , pinfo : '"&pinfo&"' , pclass :'"&pclass&"' , purl :'"&purl&"' , ptime :'"&ptime&"', pcount :'"&pcount&"' , phidden :'"&phidden&"', pcom :'"&pcom&"', pnum :'"&pnum&"' , prep :'"&prep&"' , nextp :'"&nextp&"' , ptj : '"&ptj&"' , pexif : '"&noexif&"',CameraMake:'"&CameraMake&"' , CameraModel : '"&CameraModel&"' , DateTime : '"&DateTime&"' , ImageDimension :'"&ImageDimension&"' , Software :'"&Software&"' , ISOSpeed :'"&ISOSpeed&"', FStop :'"&FStop&"' , ExposureTime :'"&ExposureTime&"', Flash :'"&Flash&"', ExposureBias :'"&ExposureBias&"' , FocalLength :'"&FocalLength&"' , MeteringMode :'"&MeteringMode&"', pcomm :'"&str&"', dcomm :'"&str2&"' }")) 
     
else
noexif=0
response.write (escape("{pid:'"&pid&"' , ptitle : '"&ptitle&"' , pinfo : '"&pinfo&"' , pclass :'"&pclass&"' , purl :'"&purl&"' , ptime :'"&ptime&"', pcount :'"&pcount&"' , phidden :'"&phidden&"', pcom :'"&pcom&"', pnum :'"&pnum&"' , prep :'"&prep&"' , nextp :'"&nextp&"' , ptj : '"&ptj&"' , pexif : '"&noexif&"', pcomm :'"&str&"', dcomm :'"&str2&"'}"))
end if



tjclass.close
set tjclass=nothing
rs.Close
Set rs=Nothing
conn.Close
Set conn=Nothing

%>
