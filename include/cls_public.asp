<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights ReservePingY.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：cls_public.asp
'= 摘    要：共用类文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-02-27
'====================================================================

Class cls_Public
	Public SysInfo, SysStat(5)
	Public Mem_Info(5), Mem_GroupSetting
	Public IsMember

	Private tmpFsoObj
	Private EA_Ini
	
	'*****************************
	'初始化环境
	'*****************************
	Private Sub Class_Initialize()
		Dim vTemp
		Dim sIniFilePath
		
		Set EA_Ini		= New cls_Ini
		sIniFilePath	= Server.MapPath (SystemFolder & "include/config.ini")
		EA_Ini.OpenFile	= sIniFilePath
		
		If EA_Ini.IsTrue Then 
			If Application(sCacheName & "IsFlush") <> 1 Then
				vTemp	= EA_Ini.ReadNode("System", "Info")
				SysInfo	= Split(vTemp, ",")

				If UBound(SysInfo) < 26 Then FoundErr = True
			
				SysStat(0) = EA_Ini.ReadNode("System", "Column_Total")
				SysStat(1) = EA_Ini.ReadNode("System", "Topic_Total")
				SysStat(2) = EA_Ini.ReadNode("System", "M_Topic_Total")
				SysStat(3) = EA_Ini.ReadNode("System", "User_Total")
				SysStat(4) = EA_Ini.ReadNode("System", "Review_Total")
			Else
				FoundErr = True
			End If
		Else
			FoundErr = True
		End If

		If FoundErr Then 
			vTemp = EA_DBO.Get_System_Info()

			If IsArray(vTemp) Then 
				SysInfo		= Split(vTemp(5, 0), ",")
				
				SysStat(0)	= vTemp(0, 0)
				SysStat(1)	= vTemp(1, 0)
				SysStat(2)	= vTemp(2, 0)
				SysStat(3)	= vTemp(3, 0)
				SysStat(4)	= vTemp(4, 0)
				
				Call EA_Ini.WriteNode("System", "Column_Total", SysStat(0))
				Call EA_Ini.WriteNode("System", "Topic_Total", SysStat(1))
				Call EA_Ini.WriteNode("System", "M_Topic_Total", SysStat(2))
				Call EA_Ini.WriteNode("System", "User_Total", SysStat(3))
				Call EA_Ini.WriteNode("System", "Review_Total", SysStat(4))

				vTemp = SysInfo
				vTemp(14) = ""
				vTemp = Join(vTemp, ",")

				EA_Ini.WriteNode "System", "Info", vTemp
				EA_Ini.Save
				
				Application.Lock 
				Application(sCacheName & "IsFlush") = 0
				Application.UnLock 

				FoundErr = False
			Else
				ErrMsg = "加载站点配置数据错误，系统已关闭。"
				Call ShowErrMsg(0, 0)
			End If
		End If

		EA_Ini.Close
		
		Call Chk_IsMember
		Call Chk_LockIp()
	End Sub
	
	'*********************
	'关闭对象过程
	'*********************
	Public Sub Close_Obj()
		On Error Resume Next

		Erase SysInfo
		Erase SysStat
		Erase Mem_Info
		Erase Mem_GroupSetting

		If IsObject(EA_Temp) Then 
			EA_Temp.Close
			Set EA_Temp=Nothing
		End If
		
		If IsObject(EA_Ini) Then
			EA_Ini.Close
			Set EA_Ini=Nothing
		End If

		EA_DBO.Close
		Set EA_DBO=Nothing

		If IsObject(EA_M_DBO) Then
			EA_M_DBO.Close_DB
			Set EA_M_DBO=Nothing
		End If

		If IsObject(tmpFsoObj) Then Set tmpFsoObj = Nothing
	End Sub
	
	'**********************
	'检测是否屏蔽ip过程
	'**********************
	Public Sub Chk_LockIp()
		Dim Ip
		Dim Temp
		Ip=Get_UserIp
		Ip=FormatIp(Ip)

		Temp=EA_DBO.Get_Ip_LockInfo(Ip)
		If IsArray(Temp) Then 
			ErrMsg="您的来访ip已被屏蔽，请与管理员联系。"
			Call ShowErrMsg(0,0)
		End If
	End Sub
	
	'************************
	'检测是否为会员过程
	'************************
	Public Function Chk_IsMember()
		Dim Temp,vTemp
		
		If Len(Session("UserData"))>0 Then 
			IsMember=True
		Else
			If Len(Request.Cookies("UserData")) Then 
				Session("UserData")=Request.Cookies("UserData")
				IsMember=True
			Else
				IsMember=False
			End If
		End If
		
		If IsMember Then 
			vTemp=Split(Session("UserData"),",")
			Mem_Info(0)=vTemp(0)
			Mem_Info(1)=vTemp(1)
			Mem_Info(2)=vTemp(2)
			Mem_Info(3)=vTemp(3)
			Mem_Info(4)=vTemp(4)
			Mem_Info(5)=vTemp(5)
			
			Temp=EA_DBO.Get_MemberLoginInfo(vTemp(0))
			If Not IsArray(Temp) Then 
				IsMember=False
				Session("UserData")			 = Empty
				Response.Cookies("UserData") = Empty
			Else
				If CLng(vTemp(4))<> CLng(Temp(16,0)) Then 
					IsMember=False
					Session("UserData")			 = Empty
					Response.Cookies("UserData") = Empty
				Else
					Call Get_Member_GroupSetting(Mem_Info(3))
				End If
			End If
		End If
		
		Chk_IsMember=IsMember
	End Function
	
	'***********************************
	'读取会员组配置信息过程
	'输入参数：
	'	1、组id
	'***********************************
	Public Sub Get_Member_GroupSetting(GroupId)
		Dim vTemp,TempArray

		vTemp=EA_Ini.ReadNode("GroupSetting","Group_"&GroupId)
		
		If vTemp="" Then 
			TempArray=EA_DBO.Get_Group_Setting(GroupId)
			If IsArray(TempArray) Then 
				Call EA_Ini.WriteNode("GroupSetting","Group_"&GroupId,TempArray(0,0)&","&Abs(TempArray(1,0))&","&TempArray(2,0))
				EA_Ini.Save
			Else
				If Not EA_Ini.IsNode("GroupSetting","Group_1") Then 
					TempArray=EA_DBO.Get_Group_Setting(1)
					If IsArray(TempArray) Then 
						Call EA_Ini.WriteNode("GroupSetting","Group_1",TempArray(0,0)&","&Abs(TempArray(1,0))&","&TempArray(2,0))
						EA_Ini.Save
						GroupId=1
					Else
						ErrMsg="系统读取会员信息时发生错误，系统已关闭。"
						Call ShowErrMsg(0,0)
					End If
				Else
					GroupId=1
				End If
			End If
			
			Get_Member_GroupSetting GroupId
		Else
			Mem_GroupSetting=Split(vTemp,",")
		End If

		EA_Ini.Close
	End Sub
	
	'**********************************
	'显示错误信息提示过程
	'输入参数：
	'	1、错误号
	'	2、显示类型
	'**********************************
	Public Sub ShowErrMsg(ErrNum,Types)
		Response.Clear
		Select Case CInt(Types)
		Case 0
			Response.Write "<font style='font-family:Verdana;font-size:11px'>"&ErrMsg&"</font>"
		Case 1
			Response.Redirect SystemFolder&"error.asp?errnum="&ErrNum
		Case 2
			Response.Write "<script language=""JavaScript"">"&vbcrlf
			Response.Write "alert("""&ErrMsg&""");"&vbcrlf
			Response.Write "history.go(-1);"&vbcrlf
			Response.Write "</script>"&vbcrlf
		End Select
		Response.End
	End Sub
	
	'****************************
	'显示成功信息提示过程
	'输入参数：
	'	1、成功号
	'	2、显示类型
	'****************************
	Public Sub ShowSusMsg(SusNum,Note)
		Response.Clear
		Response.Redirect SystemFolder&"success.asp?susnum="&SusNum&"&note="&Note
		Response.End
	End Sub
	
	'********************
	'检测是否外部提交数据过程
	'********************
	Public Sub Chk_Post()
		Dim Server_V1,Server_V2
		
		Server_V1=Cstr(Request.ServerVariables("HTTP_REFERER"))
		Server_V2=Cstr(Request.ServerVariables("SERVER_NAME"))
		
		If Mid(Server_V1,8,Len(Server_V2))<>Server_V2 Then Call ShowErrMsg(9,1)
	End Sub
	
	'****************************************************
	'检测HTML文件是否存在
	'输入参数：
	'	1、HTML文件地址
	'****************************************************
	Public Function Chk_IsExistsHtmlFile(ByVal sFilePath)
		If Not IsObject(tmpFsoObj) Then Set tmpFsoObj = CreateObject("Scripting.FileSystemObject")

		sFilePath=Server.MapPath (sFilePath)
		
		Chk_IsExistsHtmlFile = tmpFsoObj.FileExists (sFilePath)
	End Function

	'***********************************************
	'输入参数：
	'	1、HTML文件地址
	'	2、文件内容
	'***********************************************
	Public Sub Save_HtmlFile(sFilePath,sPageContent)
		On Error Resume Next
		Err.Clear

		Dim FileName
		Dim S

		Set S = Server.CreateObject("ADOD" & "B.S" & "TREAM")
		FileName=Server.MapPath(sFilePath)

		With S
			.Open
			.Charset = "utf-8"
			.WriteText sPageContent
			.SaveToFile FileName,2
			.Close
		End With

		Set S = Nothing

		If Err Then
			Response.Write "在写入文件 " & sFilePath & " 时发生错误，可能是由于服务器没有打开写入权限，请联系你的服务器管理员。"
			Response.End
		End If
	End Sub

	'*********************************
	'根据指定名称生成目录
	'*********************************
	Public Sub MakeNewsDir(foldername)
		If Not IsObject(tmpFsoObj) Then Set tmpFsoObj = CreateObject("Scripting.FileSystemObject")

		tmpFsoObj.CreateFolder foldername
	End Sub

	'***********************************
	'检查某一目录是否存在
	'***********************************
	Public Function CheckDir(FolderPath)
		If Not IsObject(tmpFsoObj) Then Set tmpFsoObj = CreateObject("Scripting.FileSystemObject")

		folderpath=Server.MapPath(".")&"\"&folderpath
	    If tmpFsoObj.FolderExists(FolderPath) Then
	       CheckDir = True
	    Else
	       CheckDir = False
	    End If
	End Function
	
	'***************************************
	'检查定时开关状态过程
	'输入参数：
	'	1、时间字符串
	'***************************************
	Public Function Chk_SystemTimer(TimeStr)
		Dim TimeArray
		Dim i

		FoundErr=False
		TimeArray=Split(TimeStr,"|")
		
		If UBound(TimeArray)<>1 Then 
			ErrMsg="定时关闭参数格式错误，请与管理员联系。"
			FoundErr=True
		Else
			TimeArray(0)=SafeRequest(0,TimeArray(0),0,1,0)
			TimeArray(1)=SafeRequest(0,TimeArray(1),0,23,0)
			
			If TimeArray(0)>TimeArray(1) Then 
				ErrMsg="定时关闭参数错误，请与管理员联系。"
				FoundErr=True
			End If

			If TimeArray(0)<=Hour(Now()) And TimeArray(1)>=Hour(Now()) Then 
				FoundErr=False
			Else
				ErrMsg=SysInfo(2)
				FoundErr=True
			End If
		End If

		Chk_SystemTimer=FoundErr
	End Function
	
	'************************************
	'截取文字长度函数
	'输入参数：
	'	1、文字内容
	'	2、文字最大长度
	'************************************
	Public Function Cut_Title(Title,TLen)
		Dim k,i,d,c
		Dim iStr
		Dim ForTotal

		If CDbl(TLen) > 0 Then
			k=0	
			d=StrLen(Title)
			iStr=""
			ForTotal = Len(Title)

			For i=1 To ForTotal
				c=Abs(Asc(Mid(Title,i,1)))
				If c>255 Then
					k=k+2
				Else
					k=k+1
				End If
				iStr=iStr&Mid(Title,i,1)
				If CLng(k)>CLng(TLen) Then 
					iStr=iStr&".."
					Exit For
				End If
			Next

			Cut_Title=iStr
		Else
			Cut_Title=""
		End If
	End Function
	
	'*******************************
	'检测文字长度函数
	'输入参数：
	'	1、文字内容
	'*******************************
	Private Function StrLen(strText)
		Dim k,i,c
		Dim ForTotal

		k=0	
		ForTotal = Len(strText)

		For i=1 To ForTotal
			c=Abs(Asc(Mid(strText,i,1)))
			If c>255 Then
				k=k+2
			Else
				k=k+1
			End If	    
		Next
		StrLen=k
	End Function 
	
	'************************************************
	'标题颜色处理函数
	'输入参数：
	'	1、颜色号
	'	2、标题文字
	'************************************************
	Public Function Add_ArticleColor(ColorCode,Title)
		Dim TempStr
		
		If Not IsNumeric(ColorCode) Then
			TempStr=Title
		Else
			Select Case CInt(ColorCode)
			Case 1
				TempStr="<span style=""color: #FF0000;"">"&Title&"</span>"
			Case 2
				TempStr="<span style=""color: #25B825;"">"&Title&"</span>"
			Case 3
				TempStr="<span style=""color: #0066CC;"">"&Title&"</span>"
			Case Else
				TempStr=Title
			End select
		End If
		
		Add_ArticleColor=TempStr
	End Function
	
	'*******************************************
	'显示文章类型函数
	'输入参数：
	'	1、是否为图片文章
	'	2、是否为推荐文章
	'*******************************************
	Public Function Chk_ArticleType(IsImg,IsTop)
		Dim TempStr
		If CBool(IsTop) Then 
			TempStr=TempStr&"<img src="""&SystemFolder&"images/public/article_top.gif"" alt=""推荐文章"" />"
		Else
			If CBool(IsImg) Then 
				TempStr=TempStr&"<img src="""&SystemFolder&"images/public/article_img.gif"" alt=""图片文章"" />"
			Else
				TempStr=TempStr&"<img src="""&SystemFolder&"images/public/article_normal.gif"" alt=""普通文章"" />"
			End If
		End If
		
		Chk_ArticleType=TempStr
	End Function
	
	'****************************************
	'检测文章是否为新文章
	'输入参数：
	'	1、发表时间
	'****************************************
	Public Function Chk_ArticleTime(PostTime)
		If DateDiff("h",PostTime,Now())<=24 Then Chk_ArticleTime="&nbsp;<img src="""&SystemFolder&"images/public/new.gif"" border=0 align=absmiddle alt=""24小时内新文章"">"
	End Function

	'**************************************************
	'替换文章正文中的内部连接函数
	'输入参数：
	'	1、文章内容
	'	2、文章地址[栏目id]
	'**************************************************
	Public Function Cov_InsideLink(ByRef StrContent,ColumnId)
		Dim i
		Dim TempArray
		Dim WordIndex
		Dim ForTotal
		
		TempArray=EA_DBO.Get_InsideLink_ByColumn(ColumnId)
		If IsArray(TempArray) Then 
			ForTotal = UBound(TempArray,2)

			For i=0 To ForTotal
				WordIndex=InStr(1,StrContent,TempArray(0,i))
				
				If WordIndex>0 Then StrContent=Replace(StrContent,TempArray(0,i),"<a href="""&TempArray(1,i)&""" rel=""external"" class=""a_link"">"&TempArray(0,i)&"</a>")
			Next
		End If
		
		Cov_InsideLink=StrContent
	End Function
	
	'************************************************
	'转换栏目路径函数
	'输入参数：
	'	1、栏目id
	'	2、路径类型
	'************************************************
	Public Function Cov_ColumnPath(ColumnId,PathType)
		If PathType=1 Then 
			Cov_ColumnPath=SystemFolder&"list.asp?classid="&ColumnId
		Else
			Cov_ColumnPath=SystemFolder&"articlelist/article_"&ColumnId&"_adddate_desc_1.htm"
		End If
	End Function 
	
	'**************************************************************
	'转换文章路径函数
	'输入参数：
	'	1、文章id
	'	2、文章发表时间
	'	3、路径类型
	'**************************************************************
	Public Function Cov_ArticlePath(ArticleId,ArticleTime,PathType)
		If PathType=1 Then 
			Cov_ArticlePath=SystemFolder&"article.asp?articleid="&ArticleId
		Else
			Cov_ArticlePath=SystemFolder&"articleview/" & Year(ArticleTime) & "-" & Month(ArticleTime) & "-" & Day(ArticleTime) & "/article_view_"&ArticleId&".htm"
		End If
	End Function
	
	Public Function Get_NavByColumnCode(sCode)
		Dim StepNum
		Dim TempStr,TempArray
		Dim i
		Dim ForTotal
		
		StepNum=Len(sCode)/4
		
		TempArray=EA_DBO.Get_Nav_List(StepNum,sCode)
		If IsArray(TempArray) Then 
			ForTotal = UBound(TempArray,2)

			For i=0 To ForTotal
				TempStr=TempStr&" -=> <a href="""&Cov_ColumnPath(TempArray(0,i),SysInfo(18))&"""><b>"&TempArray(1,i)&"</b></a>"
			Next
		End If
		
		Get_NavByColumnCode=TempStr
	End Function
	
	'*****************************************
	'简单HTML代码过滤函数
	'输入参数：
	'	1、待过滤字符串
	'*****************************************
	Public Function Base_HTMLFilter(sInputStr)
		If Len(sInputStr)>0 Then 
			sInputStr=Replace(sInputStr,Chr(13)&Chr(10),vbcrlf)
		End If
		
		Base_HTMLFilter=sInputStr
	End Function
	
	'*****************************************
	'全HTML代码过滤函数
	'输入参数：
	'	1、待过滤字符串
	'*****************************************
	Public Function Full_HTMLFilter(sInputStr)
		If Len(sInputStr)>0 Then 
			sInputStr=Replace(sInputStr, ">", "&gt;")
			sInputStr=Replace(sInputStr, "<", "&lt;")
			sInputStr=Replace(sInputStr, """", "&quot;")
			sInputStr=Replace(sInputStr, CHR(32), "&nbsp;")
			sInputStr=Replace(sInputStr, CHR(9), "&nbsp;")
			sInputStr=Replace(sInputStr, CHR(34), "&quot;")
			sInputStr=Replace(sInputStr, CHR(39), "&#39;")
			sInputStr=Replace(sInputStr, CHR(13), "")
			sInputStr=Replace(sInputStr, CHR(10) & CHR(10), "</P><P> ")
			sInputStr=Replace(sInputStr, CHR(10), "<BR>")

			Dim re
			Set re=new RegExp
			re.IgnoreCase =true
			re.Global=True
			re.Pattern="(javascript)"
			sInputStr=re.Replace(sInputStr,"<I>&#106avascript</I>")
			re.Pattern="(jscript:)"
			sInputStr=re.Replace(sInputStr,"<I>&#106script:</I>")
			re.Pattern="(js:)"
			sInputStr=re.Replace(sInputStr,"<I>&#106s:</I>")
			re.Pattern="(value)"
			sInputStr=re.Replace(sInputStr,"<I>&#118alue</I>")
			re.Pattern="(about:)"
			sInputStr=re.Replace(sInputStr,"<I>about&#58</I>")
			re.Pattern="(file:)"
			sInputStr=re.Replace(sInputStr,"<I>file&#58</I>")
			re.Pattern="(document.cookie)"
			sInputStr=re.Replace(sInputStr,"<I>documents&#46cookie</I>")
			re.Pattern="(vbscript:)"
			sInputStr=re.Replace(sInputStr,"<I>&#118bscript:</I>")
			re.Pattern="(vbs:)"
			sInputStr=re.Replace(sInputStr,"<I>&#118bs:</I>")
			re.Pattern="(on(mouse|exit|error|click|key))"
			sInputStr=re.Replace(sInputStr,"<I>&#111n$2</I>")
			sInputStr=BadWords_Filter(sInputStr)
		End If
		
		Full_HTMLFilter = sInputStr
	End Function

	'***************************************
	'全HTML格式清除转换函数
	'输入参数：
	'	1、待过滤字符串
	'***************************************
	Public Function Clean_HTMLFilter(sInputStr)
		Dim objRegExp
        sInputStr = sInputStr & ""
        Set objRegExp = new RegExp
        objRegExp.Global = True
        objRegExp.Pattern = "(<[^>]*>)"
        sInputStr = objRegExp.Replace(sInputStr,"")

		objRegExp.Pattern="\[NextPage([^\]])*\]"
		sInputStr=objRegExp.Replace(sInputStr,"")

		objRegExp.Pattern="\&.+?\;"
		sInputStr=objRegExp.Replace(sInputStr,"")

		sInputStr=Replace(sInputStr, CHR(13), "")
		sInputStr=Replace(sInputStr, CHR(10), "")
		sInputStr=Replace(sInputStr, CHR(32), "")
		sInputStr=Replace(sInputStr, CHR(9), "")
		sInputStr=Replace(sInputStr, CHR(8), "")

		Clean_HTMLFilter=sInputStr
		Set objRegExp = Nothing
	End Function

	'***************************************
	'HTML过滤逆转换函数
	'输入参数：
	'	1、待转换字符串
	'***************************************
	Public Function Un_Base_HTMLFilter(sInputStr)
		If Len(sInputStr)>0 Then 
			sInputStr = Replace(sInputStr, "</P><P> ", "&nbsp;")
			sInputStr = Replace(sInputStr, "<BR>", "&nbsp;")
		End If
		
	    Un_Base_HTMLFilter = sInputStr
	End Function

	'***************************************
	'HTML过滤逆转换函数
	'输入参数：
	'	1、待转换字符串
	'***************************************
	Public Function Un_Full_HTMLFilter(sInputStr)
		If Len(sInputStr)>0 Then 
			sInputStr = Replace(sInputStr, "</P><P> ", CHR(10) & CHR(10))
			sInputStr = Replace(sInputStr, "<BR>", CHR(10))
		End If
		
	    Un_Full_HTMLFilter = sInputStr
	End Function
	
	'****************************************
	'屏蔽字符过滤函数
	'输入参数：
	'	1、待过滤内容
	'****************************************
	Public Function BadWords_Filter(strText)
		Dim str_FilterContent
		Dim BadWord_Array
		Dim Tmp,i,TempArray
		Dim ForTotal
		
		TempArray=EA_DBO.Get_System_Info()
		If IsArray(TempArray) Then str_FilterContent=TempArray(7,0)
		
		If Not(IsNull(str_FilterContent) Or Not IsNull(strText)) Then
			BadWord_Array = Split(str_FilterContent, ";")
			ForTotal = UBound(BadWord_Array)

			For i = 0 To ForTotal
				Tmp=Split(BadWord_Array(i),"==")
				
				strText = Replace(strText, Tmp(0), Tmp(1)) 
			Next
		End If
		
		BadWords_Filter = strText
	End Function

	Public function DealJsText(Str)
		if not isnull(Str) then
			Dim re,po,ii

			Str = Replace(Str, CHR(9), "&nbsp;")
			Str = Replace(Str, CHR(39), "&#39;")
			Str = Replace(Str, CHR(13), "")
			Str = Replace(Str, CHR(10) & CHR(13), "</P><P> ")
			Str = Replace(Str, CHR(10), "")
			Str = Replace(Str, "‘", "&#39;")
			Str = Replace(Str, "’", "&#39;")
			'网友冷情圣郎提供
			Str = Replace(Str, "\", "\\")
			Str = Replace(Str, CHR(32), " ")
			Str = Replace(Str, CHR(34), "\""")
			Str = Replace(Str, CHR(39), "'")

			Set re=new RegExp
			re.IgnoreCase =true
			re.Global=True
			po=0
			ii=0

			re.Pattern="(javascript)"
			Str=re.Replace(Str,"<I>&#106avascript</I>")
			re.Pattern="(jscript:)"
			Str=re.Replace(Str,"<I>&#106script:</I>")
			re.Pattern="(js:)"
			Str=re.Replace(Str,"<I>&#106s:</I>")
			re.Pattern="(</SCRIPT>)"
			Str=re.Replace(Str,"&lt;/script&gt;")
			re.Pattern="(<SCRIPT)"
			Str=re.Replace(Str,"&lt;script")

			DealJsText = Str

			Set re=Nothing
		End if
	end Function
	
	'****************************************************
	'检测数据提交间隔时间函数
	'输入参数：
	'	1、间隔时间
	'	2、间隔符
	'	3、对照时间
	'****************************************************
	Public Function Chk_PostTime(iSpace,sSplit,sSourTime)
		Dim Flag
		Flag=False

		If Not IsDate(sSourTime) Then
			Flag=False
		Else
			If DateDiff(sSplit,sSourTime,Now())<iSpace Then 
				Flag=True
			Else
				Flag=False
			End If
		End If

		Chk_PostTime=Flag
	End Function
	
	'*************************************************************************************
	'全功能安全过滤函数
	'输入参数：
	'	1、请求方式
	'	2、请求名
	'	3、值类型
	'	4、默认值
	'	5、过滤类型
	'*************************************************************************************
	Public Function SafeRequest(Requester,RequestName,RequestType,DefaultValue,FilterType)
		Dim TempValue
		
		Select Case Requester
		Case 0
			TempValue=RequestName
		Case 1
			TempValue=Request(RequestName)
		Case 2
			TempValue=Request.Form (RequestName)
		Case 3
			TempValue=Request.QueryString (RequestName)
		Case 4
			TempValue=Request.Cookies (RequestName)
		Case 5
			TempValue=RequestName
		End Select
			
		Select Case RequestType
		Case 0
			If Not IsNumeric(TempValue) Or Len(TempValue)<=0 Then 
				TempValue=CDbl(DefaultValue)
			Else
				TempValue=CDbl(TempValue)
			End If
		Case 1
			Select Case FilterType
			Case 0
				TempValue=Replace(TempValue,"'","&#39;")
				If iDataBaseType>0 Then	TempValue=Replace(TempValue,";","；")
				TempValue=Replace(TempValue,"select","Ｓelect",1,-1,1)
			Case 1
				TempValue=Replace(TempValue,"'","&#39;")
				Call Base_HTMLFilter(TempValue)
			Case 2
				TempValue=Replace(TempValue,"'","&#39;")
				Call Full_HTMLFilter(TempValue)
			Case 3
				Call Clean_HTMLFilter(TempValue)
			End Select
		Case 2
			If Not IsDate(TempValue) Or Len(TempValue)<=0 Then 
				TempValue=CDate(DefaultValue)
			Else
				TempValue=CDate(TempValue)
			End If
		End Select
		
		SafeRequest=TempValue
	End function
	
	'***************************
	'获取来访用户IP函数
	'***************************
	Public Function Get_UserIp()
		Dim Ip,Tmp
		Dim i,IsErr
		Dim ForTotal

		IsErr=False
		
		Ip=Request.ServerVariables("REMOTE_ADDR")
		If Len(Ip)<=0 Then Ip=Request.ServerVariables("HTTP_X_ForWARDED_For")
		
		If Len(Ip)>15 Then 
			IsErr=True
		Else
			Tmp=Split(Ip,".")
			If Ubound(Tmp)=3 Then 
				ForTotal = Ubound(Tmp)

				For i=0 To ForTotal
					If Len(Tmp(i))>3 Then IsErr=True
				Next
			Else
				IsErr=True
			End If
		End If
		
		If IsErr Then 
			Get_UserIp="1.1.1.1"
		Else
			Get_UserIp=Ip
		End If
	End Function
	
	'*******************************
	'格式化ip字符串函数
	'输入参数：
	'	1、ip字符串
	'*******************************
	Public Function FormatIp(IpStr)
		Dim Tmp,i
		Dim ForTotal
		
		Tmp=Split(IpStr,".")
		ForTotal = Ubound(Tmp)

		For i=0 To ForTotal
			If Len(Tmp(i))<3 Then Tmp(i)=Right("000"&Tmp(i),3)
		Next
		
		IpStr=Join(Tmp,",")
		
		FormatIp=Replace(IpStr,",","")
	End Function

	'************************************************
	'统计页总数函数
	'输入参数：
	'	1、每页记录数
	'	2、记录总数
	'************************************************
	Public Function Stat_Page_Total(PageSize,ReCount)
		If ReCount Mod PageSize=0 Then
			Stat_Page_Total= CLng(ReCount \ PageSize)
		Else
			Stat_Page_Total= CLng((ReCount \ PageSize)+1)
		End If
	End Function
End Class
%>