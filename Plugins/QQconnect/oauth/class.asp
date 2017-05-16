<!--#include file="sha1.asp"-->
<%
Class WBoauth_54bq
	'Private Sub class_initialize
	'End Sub
	Public Function re_nonce()
		oauth_nonce = Get_nonce(32)
		oauth_timestamp = DateDiff("s","01/01/1970 08:00:00",Now())+Time_Zone*60*60
	End Function
	Public Function main()
		re_nonce()
		Dim i,GET_data,POST_data,Geturl,SIGN_data
		SIGN_data=Get_parameter(SIGN_Array)
		GET_data=Get_parameter(GET_Array)
		If Gmethod<>"GET" Then
			POST_data=Get_parameter(POST_Array)
		Else
		
		End If
		oauth_signature=Get_signature(SIGN_data)
		Geturl=Get_url(oauth_signature,GET_data)
		main=Get_Page(Gmethod,Geturl,POST_data)
	End Function
	'签名参数
	Private Function Get_parameter(myArray)	'获取sign串
		Get_parameter=""
		If IsArray(myArray) Then 
			for i=LBound(myArray) To Ubound(myArray)
				execute("Get_parameter=Get_parameter&""&"&myArray(i)&"=""&"&myArray(i)&"")
			Next
		End If 
		If Left(Get_parameter,1)="&" Then Get_parameter=Right(Get_parameter,Len(Get_parameter)-1)
	End Function
	'签名
	Private Function Get_signature(parameter)
		Get_signature = strUrlEnCode(b64_hmac_sha1(oauth_consumer_secret&"&"&oauth_token_secret,Gmethod&"&" & strUrlEnCode(API_URL) & "&" & strUrlEnCode(parameter)))
'		WB.die(parameter)
	End Function
	'签名
	public Function qianming(parameter)
		qianming = strUrlEnCode(b64_hmac_sha1(oauth_consumer_secret&"&"&oauth_token_secret,Gmethod&"&" & strUrlEnCode(API_URL) & "&" & strUrlEnCode(parameter)))
	End Function
	
	Private Function Get_url(oauth_signature,parameter)
		Get_url =  API_URL&"?"&parameter & "&oauth_signature=" & oauth_signature
	End Function
	public Function echo(str)
		response.write str
	End Function
	public Function die(str)
		response.write str
		response.end
	End Function

	Public Function GetContent(str,start,last)
		If Instr(lcase(str),lcase(start))>0 then
			GetContent=Right(str,Len(str)-Instr(lcase(str),lcase(start))-Len(start)+1)
			If last<>"" Then
			GetContent=Left(GetContent,Instr(lcase(GetContent),lcase(last))-1)
			End if
		Else
			GetContent="nothing"
		End if
	End Function
	
	Private Function Get_Page(Gmethod,url,value)
		On Error Resume Next
		Dim Retrieval
		Set Retrieval=CreateObject("Microsoft.XMLHTTP")
		With Retrieval
		.Open Gmethod, url, False, "", ""
		.setRequestHeader "CONTENT-TYPE","application/x-www-form-urlencoded"
		.Send(value)
		Get_Page=BytesToBstr(.ResponseBody)
		End With
		Set Retrieval=Nothing
		If err.number <> 0   then
			response.clare()
			Err.Raise   6       '   引发溢出错误.
			Get_Page=CStr(Err.Number)	'	Err.Description	Err.Source
			Err.Clear       '   清除该错误。
		End If
	End Function
	Private Function BytesToBstr(body)
		dim objstream
		set objstream=Server.CreateObject("adodb.stream")
		objstream.Type=1
		objstream.Mode=3
		objstream.Open
		objstream.Write body
		objstream.Position=0
		objstream.Type=2
		objstream.Charset="utf-8"
		BytesToBstr=objstream.ReadText
		objstream.Close
		set objstream=nothing
	End Function
	Public Function setcookie(key,value)
		Response.Cookies("shmshz")(key)=value
		Response.Cookies("shmshz").Path = "/"
		If oauth_Cookies_path<>"" Then
			Response.Cookies("shmshz").domain = oauth_Cookies_path
		End If
	End Function
	Public Function cookiepath()
		Response.Cookies("shmshz").Path = "/"
	End Function
	Public Function getcookie(key)
		getcookie=Request.Cookies("shmshz")(key)
	End Function
End Class

'=========通用函数
Function Get_nonce(digit) '生成随机数字串
randomize
Get_nonce=int((99998999)*Rnd + 1000)
End Function
'=========生成随机串
	Function Get_nonce2(byVal maxLen)
        Dim strNewPass
        Dim whatsNext, upper, lower, intCounter
        Randomize
        For intCounter = 1 To maxLen
        whatsNext = Int((1 - 0 + 1) * Rnd + 0)
        If whatsNext = 0 Then
        upper = 122
        lower = 100
        Else
        upper = 57
        lower = 48
        End If
        strNewPass = strNewPass & Chr(Int((upper - lower + 1) * Rnd + lower))
        Next
        Get_nonce2 = strNewPass
	End Function
'=========返回参数转json http_build_query
Function http_restore_query(str)
	Dim tmpa,tmpb,tmpc,i,tmpd,tmpz:tmpz=0
	If Len(str)<90 Then tmpz=2
	tmpa="{"
	if instr(str,"&") >0 then
		tmpb=split(str,"&")
		For i=LBound(tmpb) to UBound(tmpb)
			if instr(tmpb(i),"=") >0 then
				tmpc = split(tmpb(i),"=")
				tmpa = tmpa & """"&tmpc(0)&""":""" & tmpc(1) & ""","
			end if	
		next
	else
		if instr(str,"=") >0 then
			tmpd = split(str,"=")
			tmpa = tmpa & """"&tmpd(0)&""":""" & tmpd(1) & ""","
		end if	
	end if
	tmpa = tmpa & """rett"":"&tmpz&"}"
	http_restore_query=tmpa
End Function
'=========返回参数转json http_build_query2

Function http_restore_query2(str)
	Dim tmpa,tmpb,tmpc,i,tmpd,tmpz:tmpz=0
	if instr(str,"&") >0 then
		tmpb=split(str,"&")
		For i=LBound(tmpb) to UBound(tmpb)
			if instr(tmpb(i),"=") >0 then
				tmpc = split(tmpb(i),"=")
				tmpd=tmpc(0)&"="""&tmpc(1)&""""
				Execute tmpd
			end if	
		next
	else
		if instr(str,"=") >0 then
			tmpd = split(str,"=")
			tmpe=tmpd(0)&"="""&tmpd(1)&""""
			Execute tmpe
		end if	
	end if
End Function

Function strUrlEnCode(byVal strUrl)
	Dim tmp
	tmp = Server.URLEncode(strUrl)
	tmp = Replace(tmp,"%5F","_")
	tmp = Replace(tmp,"%2E",".")
	strUrlEnCode = Replace(tmp,"%2D","-")
End Function
%>
