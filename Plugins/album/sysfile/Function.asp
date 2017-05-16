<%
'==================================================================
'过程名：CreateDIR
'作  用：建立目录的程序，如果有多级目录，则一级一级的创建
'参  数：LocalPath ------目录的路径
'返回值：True  ----创建成功
'        false-----创建不成功
'==================================================================
Function CreateDIR(LocalPath)
    On Error Resume Next
    LocalPath = Replace(LocalPath, "\", "/")
    Set FileObject = server.CreateObject("Scripting.FileSystemObject")
    patharr = Split(LocalPath, "/") 
    path_level = UBound(patharr)
    For I = 0 To path_level
        If I = 0 Then pathtmp = patharr(0) & "/" Else pathtmp = pathtmp & patharr(I) & "/"
        cpath = Left(pathtmp, Len(pathtmp) - 1)
        If Not FileObject.FolderExists(cpath) Then FileObject.CreateFolder cpath
    Next
    Set FileObject = Nothing
    If Err.Number <> 0 Then
        CreateDIR = False
        Err.Clear
    Else
        CreateDIR = True
    End If
End Function
'==================================================================
'过程名：RegExpTest
'作  用：提取字符串最后的数字
'参  数：ConStr ----要提取字符串
'返回值：True  ----提取的数字字符
'==================================================================
Function RegExpTest(ConStr)
	Set Re = New Regexp 
	Re.IgnoreCase = True
	Re.Global = True
	Re.Pattern ="\d+$"
	Set Matches =Re.Execute(ConStr)
	For Each Match in Matches
		TempStr=Match.Value
	Next
	RegExpTest =TempStr
end Function
'==================================================================
'过程名：copynumber
'作  用：格式化数字
'参  数：f ----要格式化的数字
'		 s ----格式化的位数
'		 n ----格式化形式
'			0-- s是格式化的位数
'			1-- 取s的字符数
'返回值：  ----格式化的数字字符
'==================================================================
function copynumber(f,s,n)
if n=1 then
	copynumber=String(len(s)-len(f),"0")&f
else
	if s>len(f) then
		copynumber=String(s-len(f),"0")&f
	else
		copynumber=f
	end if
end if
end function
'==================================================================
'过程名：IsObjInstalled
'作  用：检查组件是否已经安装
'参  数：strClassString ----组件名
'返回值：True  ----已经安装
'        False ----没有安装
'==================================================================
Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function

'==================================================================
'过程名：ReportFileStatus
'作  用：判断文件是否存在
'参  数：FileName-----文件详细路径及名称
'返回值：True  ----文件存在
'        False ----文件不存在
'==================================================================
Function ReportFileStatus(FileName)
    Dim FSO,msg
    Set FSO = CreateObject("Scripting.FileSystemObject")
    If (FSO.FileExists(FileName)) Then
        ReportFileStatus =True
    Else
        ReportFileStatus = False
    End If
	Set FSO=nothing
End Function
'==================================================================
'过程名：Fsodelwj
'作  用：fso 删除文件
'参  数：wjurl----文件路径
'返回值：没有
'==================================================================

Function Fsodelwj(wjurl)
if wjurl="" then
		exit function
end if	
if Left(wjurl, 7)<>"http://" then
	delpic=Server.MapPath(wjurl)
	Set fso = CreateObject("Scripting.FileSystemObject")
	if fso.FileExists(delpic) = True Then
		set nf=fso.GetFile(delpic)
		nf.delete
	end if
Set fso = nothing
end if

End Function
'==================================================================
'过程名：FormatName
'作  用：格式化文件名
'参  数：FileExt----文件后缀名
'返回值：按年+月+日+随机数+后缀名的文件名
'==================================================================

Function FormatName(FileExt)
	Dim RanNum
	Randomize
	RanNum = Int(900*rnd)+100
	FormatName = Year(now)&Month(now)&Day(now)&Hour(now)&Minute(now)&Second(now)&RanNum&"."&FileExt
End Function

'==================================================
'过程名：Saveimage
'作  用：保存远程的文件到本地
'参  数：RemoteFileUrl -------- 远程文件URL
'        LocalFileName --------- 本地文件名
'        imgtype-----1 图片文件，0 所有文件
'返回值：True  ------------------- 保存成功
'        false ----------------- 保存不成功
'==================================================
function saveimage(RemoteFileUrl,LocalFileName,imgtype)
	dim geturl,objstream,imgs,httpimgs
	if Left(RemoteFileUrl, 7)<>"http://" then
		saveimage=False 
		exit function
	end if	
	GIFflType = chrb(71) & chrb(73) & chrb(70) 
    BMPflType = chrb(66) & chrb(77)
    JPGflType = Chrb(255) & Chrb(216) & Chrb(255)
	saveimage=True
	set httpimgs=server.createobject("MSXML2.XMLHTTP.3.0")
	With httpimgs
		httpimgs.open "Get",RemoteFileUrl,false, "", ""
		httpimgs.send()
		if httpimgs.readystate<>4 or httpimgs.Status > 300  then
			saveimage=False 
			exit function
		end if
		imgs=httpimgs.responsebody
	End With
	set httpimgs=nothing
	if imgtype=1 then
		if midb(imgs, 1, 3) <>JPGflType then
		elseif midb(imgs, 1, 3) <> GIFflType then
		elseif midb(imgs, 1, 2) <> BMPflType then
			saveimage=False 
			exit function
		end if
	end if
	set objstream = server.createobject("adodb.stream")
	objstream.type =1
	objstream.open
	objstream.write imgs
	objstream.savetofile server.MapPath(LocalFileName),2
	objstream.close()'关闭对象
	if Err.number<>0 then
		saveimage=False
		Exit Function
		Err.Clear
	end If
	set objstream=nothing
end function

'==================================================
'过程名：Xiaotu
'作  用：生成缩略图 (采用ASPJPEG组件)
'参  数：SaveFile-------------大图路径
'      ：xiaopath-------------小图路径
'	   ：FileName-------------大图名称
'      ：S_Width------------小图最宽值
'      ：S_Height-----------小图最高值
'      ：Nametype---1 xiaopath必须是完整路径文件名，
'					0 xiaopath是小图路径，生成的缩略图名是大图前面+S
'      ：shtype------------原图是否添加水印
'					0 不添加水印
'					1 文字水印
'					2 图片水印
'返回值：缩略图的路径
'==================================================
Function  Xiaotu(SaveFile,xiaopath,FileName,S_Width,S_Height,Nametype,shtype)
	Dim Jpeg,Path
	Set Jpeg = Server.CreateObject("Persits.Jpeg")
	Path =Server.MapPath(SaveFile)
	Jpeg.Open Path
	wid=Jpeg.OriginalWidth
	hei=Jpeg.OriginalHeight
	bili1=S_Width/S_Height
	Select Case (small_gz)
		Case 0
			Jpeg.Width = S_Width
			Jpeg.Height = S_Height
		Case 1
			if wid>hei then
				Jpeg.Width=S_Width
				Jpeg.Height=hei/wid*S_Width
			else
				Jpeg.Height=S_Width
				Jpeg.Width=wid/hei*S_Width
			end if
		Case 2
			if wid>hei then
				Jpeg.Width=S_Width
				jpeg.Height=S_Width/wid*hei
				gao=(Jpeg.Height-S_Height)/2
				jpeg.crop 0,gao,S_Width,S_Height+gao
			else
				Jpeg.Height=S_Width
				jpeg.Width=S_Width/hei*wid
				bian=(Jpeg.Width-S_Height)/2
				jpeg.crop bian,0,S_Height+bian,S_Width
			end if
		End Select
	if Nametype=1 then
		xiaova=xiaopath
	else
		xiaova=xiaopath&"/"&"S_"&FileName
	end if
	Jpeg.Save Server.MapPath(xiaova)
	Set Jpeg = Nothing
	Select Case (shtype)
		Case 1
			call Picwzsy(SaveFile,suiimgys,suiwzsize,suiwzxx)
		Case 2
			call Pictpsy(SaveFile,suiimgpath,suiwz,suiimgysdel,suiimgtm)
		End Select

'参  数：SaveFile-----------------图片路径
'      ：SuiColor-------------水印字体颜色
'	   ：SuiSize------------水印文字的大小
'      ：SuiName--------------文字水印信息

	xiaotu=xiaova
end Function

'==================================================
'过程名：Pictpsy
'作  用：生成图片水印 (采用ASPJPEG组件)
'参  数：SaveFile-----------------图片路径
'      ：SuiPath------------水印图片的路径
'      ：SuiWeiz------------水印图片的位置
'      ：SuiDelt----------水印图片去除底色
'      ：suitm--------------水印图片透明度
'返回值：没有
'==================================================
Function Pictpsy(SaveFile,SuiPath,SuiWeiz,SuiDelt,suitm)
	Set Photo = Server.CreateObject("Persits.Jpeg")
	Photo.Open Server.MapPath(SaveFile)
	daWidth=Photo.width
	daHeight=Photo.height
	Set fso = CreateObject("Scripting.FileSystemObject")
	if fso.FileExists(SuiPath) =True Then
		exit function
		Set fso = Nothing
	end if
	Set Logo = Server.CreateObject("Persits.Jpeg")
	Logo.Open Server.MapPath(SuiPath)
	loWidth=Logo.Width
	loHeight=Logo.Height

	Select Case (SuiWeiz) 
		Case 1									'//左上
			sui_x=0
			sui_y=0	
		Case 2									'//左下
			sui_x=0
			sui_y=daHeight-loHeight
		Case 3									'//居中
			sui_x=(daWidth-loWidth)/2
			sui_y=(daHeight-loHeight)/2
		Case 4									'//右上
			sui_x=daWidth-loWidth
			sui_y=0
		Case 5									'//右下
			sui_x=daWidth-loWidth
			sui_y=daHeight-loHeight

	End Select
	If SuiDelt="" Then
		Photo.DrawImage sui_x, sui_y, Logo
	else
		delColor = Replace(SuiDelt,"#","&H")
		Photo.DrawImage sui_x, sui_y, Logo , suitm,delColor,0
								'加入LOGO水印效里的关键所在，“sui_x”为位于图片的X坐标，
								'“sui_y”为位于图片的Y坐标，“suiimgtm”为淡化程度，
								'“delColor”为去除LOGO水印图中的某个颜色，从而达到完成透明的效果，
								'“100”为去除颜色周边的相应过渡色的百分比。
	end if

	Photo.Save Server.MapPath(SaveFile)
	Set fso = Nothing
	Set Photo = Nothing
	Set Logo = Nothing
end Function
'==================================================
'过程名：Picwzsy
'作  用：生成文字水印 (采用ASPJPEG组件)
'参  数：SaveFile-----------------图片路径
'      ：SuiColor-------------水印字体颜色
'	   ：SuiSize------------水印文字的大小
'      ：SuiName--------------文字水印信息
'返回值：没有
'==================================================
Function Picwzsy(SaveFile,SuiColor,SuiSize,SuiName)
	Dim Jpeg
	Set Jpeg = Server.CreateObject("Persits.Jpeg")
	Jpeg.Open Server.MapPath(SaveFile)
	Wid=Jpeg.Width
	Hei=Jpeg.Height
	FontColor = Replace(SuiColor,"#","&h")	
	Jpeg.Canvas.Font.Color = FontColor			'//水印字体颜色
	Jpeg.Canvas.Font.Size = SuiSize			    '//水印文字的大小
	Jpeg.Canvas.Font.Family =  "Tahoma"  		    '//字体
	if SuiChut="1" then
		Jpeg.Canvas.Font.Bold = True 			'//水印文字是否为粗体，True=粗体 False=正常。
	else
		Jpeg.Canvas.Font.Bold = False
	end if
	
	if Jpeg.OriginalWidth=>200 or Jpeg.OriginalHeight=>200 then
	    'SuiName= "图片来自峥言锋语|Photo By Onewz.Com"
	    Jpeg.Canvas.Pen.Color = &H000000' black 颜色 
        Jpeg.Canvas.Pen.Width = 25 '画笔宽度 
        Jpeg.Canvas.Brush.Solid = False '是否加粗处理 
        Jpeg.Canvas.Line 1, Jpeg.OriginalHeight-5, Jpeg.OriginalWidth, Jpeg.OriginalHeight-5
 		Jpeg.Canvas.Print Jpeg.OriginalWidth-108,Jpeg.OriginalHeight-16,CStr(Split(SuiName,"|")(1))
		Jpeg.Canvas.Print 5,Jpeg.OriginalHeight-16,CStr(Split(SuiName,"|")(0))
	 else
	

	Jpeg.Canvas.Print 5, Jpeg.OriginalHeight-16, SuiName
	end if
	Jpeg.Save Server.MapPath(SaveFile)
	Set Jpeg = Nothing 
end Function

function HTMLEncode2(fString)
if fString<>"" then
	fString = Replace(fString, ">", "&gt;")
	fString = Replace(fString, "<", "&lt;")
	fString = Replace(fString, CHR(34), "&quot;")
	fString = Replace(fString, CHR(13), "")
	fString = Replace(fString, CHR(10), "<br>")
	fString = Replace(fString, CHR(32), "&nbsp;")
end if
	HTMLEncode2 = fString
end function
'----------------------------------------------------------------------
'转发时请保留此声明信息,这段声明不并会影响你的速度!
'*******************    无组件上传类   ********************************
'文件属性：例如上传文件为c:\myfile\doc.txt
'FileName    文件名       字符串    "doc.txt"
'FileSize    文件大小     数值       1210
'FileType    文件类型     字符串    "text/plain"
'FileExt     文件扩展名   字符串    "txt"
'FilePath    文件原路径   字符串    "c:\myfile"
'由于Scripting.Dictionary区分大小写,所以在网页及ASP页的项目名都要相同的大小
'写,如果人习惯用大写或小写,为了防止出错的话,可以把
'sFormName = Mid (sinfo,iFindStart,iFindEnd-iFindStart)
'改为
'(小写者)sFormName = LCase(Mid (sinfo,iFindStart,iFindEnd-iFindStart))
'(大写者)sFormName = UCase(Mid (sinfo,iFindStart,iFindEnd-iFindStart))
'**********************************************************************
'----------------------------------------------------------------------
dim oUpFileStream

Class upload_file
  
dim Form,File,Version
  
Private Sub Class_Initialize 
	'定义变量
	dim RequestBinDate,sStart,bCrLf,sInfo,iInfoStart,iInfoEnd,tStream,iStart,oFileInfo
	dim iFileSize,sFilePath,sFileType,sFormvalue,sFileName
	dim iFindStart,iFindEnd
	dim iFormStart,iFormEnd,sFormName
	'代码开始
	Version="无组件上传类 Version 0.96"
	set Form = Server.CreateObject("Scripting.Dictionary")
	set File = Server.CreateObject("Scripting.Dictionary")
	if Request.TotalBytes < 1 then Exit Sub
	set tStream = Server.CreateObject("adodb.stream")
	set oUpFileStream = Server.CreateObject("adodb.stream")
	oUpFileStream.Type = 1
	oUpFileStream.Mode = 3
	oUpFileStream.Open 
	oUpFileStream.Write Request.BinaryRead(Request.TotalBytes)
	oUpFileStream.Position=0
	RequestBinDate = oUpFileStream.Read 
	iFormEnd = oUpFileStream.Size
	bCrLf = chrB(13) & chrB(10)
	'取得每个项目之间的分隔符
	sStart = MidB(RequestBinDate,1, InStrB(1,RequestBinDate,bCrLf)-1)
	iStart = LenB (sStart)
	iFormStart = iStart+2
	'分解项目
  Do
    iInfoEnd = InStrB(iFormStart,RequestBinDate,bCrLf & bCrLf)+3
    tStream.Type = 1
    tStream.Mode = 3
    tStream.Open
    oUpFileStream.Position = iFormStart
    oUpFileStream.CopyTo tStream,iInfoEnd-iFormStart
    tStream.Position = 0
    tStream.Type = 2
    tStream.Charset ="UTF-8"
    sInfo = tStream.ReadText      
    '取得表单项目名称
    iFormStart = InStrB(iInfoEnd,RequestBinDate,sStart)-1
    iFindStart = InStr(22,sInfo,"name=""",1)+6
    iFindEnd = InStr(iFindStart,sInfo,"""",1)
    sFormName = Mid (sinfo,iFindStart,iFindEnd-iFindStart)
    '如果是文件
	if InStr (45,sInfo,"filename=""",1) > 0 then
		set oFileInfo= new FileInfo
		'取得文件属性
		iFindStart = InStr(iFindEnd,sInfo,"filename=""",1)+10
		iFindEnd = InStr(iFindStart,sInfo,"""",1)
		sFileName = Mid (sinfo,iFindStart,iFindEnd-iFindStart)
		oFileInfo.FileName = GetFileName(sFileName)
		oFileInfo.FilePath = GetFilePath(sFileName)
		oFileInfo.FileExt = GetFileExt(sFileName)
		iFindStart = InStr(iFindEnd,sInfo,"Content-Type: ",1)+14
		iFindEnd = InStr(iFindStart,sInfo,vbCr)
		oFileInfo.FileType = Mid (sinfo,iFindStart,iFindEnd-iFindStart)
		oFileInfo.FileStart = iInfoEnd
		oFileInfo.FileSize = iFormStart -iInfoEnd -2
		oFileInfo.FormName = sFormName
		file.add sFormName,oFileInfo
	else
		'如果是表单项目
		tStream.Close
		tStream.Type = 1
		tStream.Mode = 3
		tStream.Open
		oUpFileStream.Position = iInfoEnd 
		oUpFileStream.CopyTo tStream,iFormStart-iInfoEnd-2
		tStream.Position = 0
		tStream.Type = 2
		tStream.Charset = "UTF-8"
		sFormvalue = tStream.ReadText 
		form.Add sFormName,sFormvalue
	end if
	tStream.Close
    iFormStart = iFormStart+iStart+2
    '如果到文件尾了就退出
    loop until (iFormStart+2) = iFormEnd 
	RequestBinDate=""
	set tStream = nothing
End Sub

Private Sub Class_Terminate  
  '清除变量及对像
	if not Request.TotalBytes<1 then
		oUpFileStream.Close
		set oUpFileStream =nothing
    end if
	Form.RemoveAll
	File.RemoveAll
	set Form=nothing
	set File=nothing
End Sub
   
 '取得文件路径
Private function GetFilePath(FullPath)
	If FullPath <> "" Then
		GetFilePath = left(FullPath,InStrRev(FullPath, "\"))
	Else
		GetFilePath = ""
	End If
End function
 
'取得文件名
Private function GetFileName(FullPath)
  If FullPath <> "" Then
    GetFileName = mid(FullPath,InStrRev(FullPath, "\")+1)
    Else
    GetFileName = ""
  End If
End function

'取得扩展名
Private function GetFileExt(FullPath)
  If FullPath <> "" Then
    GetFileExt = mid(FullPath,InStrRev(FullPath, ".")+1)
    Else
    GetFileExt = ""
  End If
End function

End Class

'文件属性类
Class FileInfo
  dim FormName,FileName,FilePath,FileSize,FileType,FileStart,FileExt
  Private Sub Class_Initialize 
    FileName = ""
    FilePath = ""
    FileSize = 0
    FileStart= 0
    FormName = ""
    FileType = ""
    FileExt = ""
  End Sub
  
'保存文件方法
Public function SaveToFile(FullPath)
	dim oFileStream,ErrorChar,i
	SaveToFile=1
    if trim(fullpath)="" or right(fullpath,1)="/" then exit function
    set oFileStream=CreateObject("Adodb.Stream")
    oFileStream.Type=1
    oFileStream.Mode=3
    oFileStream.Open
    oUpFileStream.position=FileStart
    oUpFileStream.copyto oFileStream,FileSize
    oFileStream.SaveToFile FullPath,2
	oFileStream.Close
	set oFileStream=nothing 
	SaveToFile=0
end function
End Class


''----------------------------------------------------
''格式化数字
''----------------------------------------------------
function doublenum(Num)
    if Num > 9 then 
        doublenum = Num 
    else 
        doublenum = "0"& Num
    end if
end function
'*************************************
'日期转换函数
'*************************************

Function DateToStr(DateTime, ShowType)
    Dim DateMonth, DateDay, DateHour, DateMinute, DateWeek, DateSecond
    Dim FullWeekday, shortWeekday, Fullmonth, Shortmonth, TimeZone1, TimeZone2
    TimeZone1 = "+0800"
    TimeZone2 = "+08:00"
    FullWeekday = Array("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
    shortWeekday = Array("Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat")
    Fullmonth = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
    Shortmonth = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

    DateMonth = Month(DateTime)
    DateDay = Day(DateTime)
    DateHour = Hour(DateTime)
    DateMinute = Minute(DateTime)
    DateWeek = Weekday(DateTime)
    DateSecond = Second(DateTime)
    If Len(DateMonth)<2 Then DateMonth = "0"&DateMonth
    If Len(DateDay)<2 Then DateDay = "0"&DateDay
    If Len(DateMinute)<2 Then DateMinute = "0"&DateMinute
    Select Case ShowType
        Case "Y-m-d"
            DateToStr = Year(DateTime)&"-"&DateMonth&"-"&DateDay
        Case "Y-m-d H:I A"
            Dim DateAMPM
            If DateHour>12 Then
                DateHour = DateHour -12
                DateAMPM = "PM"
            Else
                DateHour = DateHour
                DateAMPM = "AM"
            End If
            If Len(DateHour)<2 Then DateHour = "0"&DateHour
            DateToStr = Year(DateTime)&"-"&DateMonth&"-"&DateDay&" "&DateHour&":"&DateMinute&" "&DateAMPM
        Case "Y-m-d H:I:S"
            If Len(DateHour)<2 Then DateHour = "0"&DateHour
            If Len(DateSecond)<2 Then DateSecond = "0"&DateSecond
            DateToStr = Year(DateTime)&"-"&DateMonth&"-"&DateDay&" "&DateHour&":"&DateMinute&":"&DateSecond
        Case "YmdHIS"
            DateSecond = Second(DateTime)
            If Len(DateHour)<2 Then DateHour = "0"&DateHour
            If Len(DateSecond)<2 Then DateSecond = "0"&DateSecond
            DateToStr = Year(DateTime)&DateMonth&DateDay&DateHour&DateMinute&DateSecond
        Case "ym"
            DateToStr = Right(Year(DateTime), 2)&DateMonth
        Case "d"
            DateToStr = DateDay
        Case "ymd"
            DateToStr = Right(Year(DateTime), 4)&DateMonth&DateDay
        Case "mdy"
            Dim DayEnd
            Select Case DateDay
            Case 1
                DayEnd = "st"
            Case 2
                DayEnd = "nd"
            Case 3
                DayEnd = "rd"
            Case Else
                DayEnd = "th"
        End Select
        DateToStr = Fullmonth(DateMonth -1)&" "&DateDay&DayEnd&" "&Right(Year(DateTime), 4)
        Case "w,d m y H:I:S"
            DateSecond = Second(DateTime)
            If Len(DateHour)<2 Then DateHour = "0"&DateHour
            If Len(DateSecond)<2 Then DateSecond = "0"&DateSecond
            DateToStr = shortWeekday(DateWeek -1)&","&DateDay&" "& Left(Fullmonth(DateMonth -1), 3) &" "&Right(Year(DateTime), 4)&" "&DateHour&":"&DateMinute&":"&DateSecond&" "&TimeZone1
        Case "y-m-dTH:I:S"
            If Len(DateHour)<2 Then DateHour = "0"&DateHour
            If Len(DateSecond)<2 Then DateSecond = "0"&DateSecond
            DateToStr = Year(DateTime)&"-"&DateMonth&"-"&DateDay&"T"&DateHour&":"&DateMinute&":"&DateSecond&TimeZone2
        Case Else
            If Len(DateHour)<2 Then DateHour = "0"&DateHour
            DateToStr = Year(DateTime)&"-"&DateMonth&"-"&DateDay&" "&DateHour&":"&DateMinute
    End Select
End Function

'个人代码风格注释（变量名中第一个小写字母表表示变量类型）
'i:为Integer型;
's:为String;
Function U2UTF8(Byval a_iNum)
    Dim sResult,sUTF8
    Dim iTemp,iHexNum,i

    iHexNum = Trim(a_iNum)

    If iHexNum = "" Then
        Exit Function
    End If

    sResult = ""

    If (iHexNum < 128) Then
        sResult = sResult & iHexNum
    ElseIf (iHexNum < 2048) Then
        sResult = ChrB(&H80 + (iHexNum And &H3F))
        iHexNum = iHexNum \ &H40
        sResult = ChrB(&HC0 + (iHexNum And &H1F)) & sResult
    ElseIf (iHexNum < 65536) Then
        sResult = ChrB(&H80 + (iHexNum And &H3F))
        iHexNum = iHexNum \ &H40
        sResult = ChrB(&H80 + (iHexNum And &H3F)) & sResult
        iHexNum = iHexNum \ &H40
        sResult = ChrB(&HE0 + (iHexNum And &HF)) & sResult
    End If

    U2UTF8 = sResult
End Function

Function GB2UTF(Byval a_sStr)
    Dim sGB,sResult,sTemp
    Dim iLen,iUnicode,iTemp,i

    sGB = Trim(a_sStr)
    iLen = Len(sGB)
    For i = 1 To iLen
         sTemp = Mid(sGB,i,1)
         iTemp = Asc(sTemp)

         If (iTemp>127 OR iTemp<0) Then
             iUnicode = AscW(sTemp)
             If iUnicode<0 Then
                 iUnicode = iUnicode + 65536
             End If
        Else
            iUnicode = iTemp
        End If

        sResult = sResult & U2UTF8(iUnicode)
    Next

    GB2UTF = sResult
End Function

'调用方法 Response.BinaryWrite(GB2UTF("中国人"))



%>
