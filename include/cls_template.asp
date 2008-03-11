<!--#Include File="tag_friend.asp"-->
<!--#Include File="tag_placard.asp"-->
<!--#Include File="tag_comment.asp"-->
<!--#Include File="tag_vote.asp"-->
<!--#Include File="tag_column.asp"-->
<!--#Include File="tag_info.asp"-->
<!--#Include File="tag_adsense.asp"-->
<!--#Include File="tag_topic.asp"-->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：cls_template.asp
'= 摘    要：模版类文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-03-05
'====================================================================

Class cls_Template
	Public Title,Nav
	Public TemplatePath
	Public PageArray(4)

	Public P_Prefix, P_Suffix

	Private LastPosition
	Private i
	Private S
	
	'*****************************
	'对象类初始化过程
	'*****************************
	Private Sub Class_Initialize()
		TemplatePath = "templates/"
		P_Prefix = "<!--"
		P_Suffix = "-->"
	End Sub

	Public Sub Close()
		Erase PageArray
	End Sub
	
	'***********************************************
	'加载模版过程
	'输入参数：
	'	1、模版id
	'***********************************************
	Public Function Load_Template(TemplateId, TemplateType)
		Dim Temp
		Temp = EA_DBO.Get_Template_Info(TemplateId, TemplateType)

		If IsArray(Temp) Then 
			TemplateId = Temp(2, 0)

			PageArray(0) = EA_DBO.Get_Theme_Name(TemplateId)(0, 0)		'template name
			PageArray(1) = EA_DBO.Get_Template_Info(0, 1)(0, 0)			'template css
			PageArray(2) = EA_DBO.Get_Template_Info(0, 2)(0, 0)			'template head
			PageArray(3) = EA_DBO.Get_Template_Info(0, 3)(0, 0)			'template foot
			Load_Template= Temp(0, 0)
		Else
			ErrMsg = Replace(SysMsg(9), "$1", "")
			ErrMsg = Replace(ErrMsg, "$2", Err.Description)
			Call EA_Pub.ShowErrMsg(0, 0)
		End If
	End Function

	Public Function Load_Template_File(ByRef sFileName)
		Err.Clear 
		On Error Resume Next
		
		Set S = Server.CreateObject("ADOD" & "B.S" & "TREAM")
		With S
			.Mode = 3
			.Type = 2
			.Open
			.LoadFromFile(Server.MapPath(TemplatePath&sFileName))
			Load_Template_File = EA_Pub.Bytes2bStr(.ReadText, "gb2312")
			.Close
		End With
		Set S = Nothing
		
		If Err Then 
			ErrMsg = Replace(SysMsg(9), "$1", sFileName)
			ErrMsg = Replace(ErrMsg, "$2", Err.Description)
			Call EA_Pub.ShowErrMsg(0, 0)
		End If
	End Function


	Public Function ChkBlock (ByRef sBlockName,ByRef sContent)
		Dim sBlockBeginStr,sBlockEndStr

		sBlockBeginStr	= P_Prefix & sBlockName & " Begin" & P_Suffix
		sBlockEndStr	= P_Prefix & sBlockName & " End" & P_Suffix

		If InStr(1, sContent, sBlockBeginStr) And InStr(1, sContent, sBlockEndStr) Then
			ChkBlock = True
		Else
			ChkBlock = False
		End If
	End Function

	Public Function GetBlock(ByRef sBlockName,ByRef sContent)
		Dim iBlockBegin,iBlockEnd
		Dim sBlockBeginStr,sBlockEndStr

		sBlockBeginStr	= P_Prefix & sBlockName & " Begin" & P_Suffix
		sBlockEndStr	= P_Prefix & sBlockName & " End" & P_Suffix

		iBlockBegin	= InStr(1,sContent,sBlockBeginStr)
		If iBlockBegin > 0 Then
			iBlockEnd = InStr(iBlockBegin,sContent,sBlockEndStr)

			GetBlock  = Mid(sContent,iBlockBegin + Len(sBlockBeginStr),iBlockEnd - (iBlockBegin + Len(sBlockBeginStr)))
			
			sContent  = Left(sContent,iBlockBegin-1) & VBCrlf & P_Prefix & sBlockName & "s" & P_Suffix & VBCrlf &  Right(sContent,Len(sContent)-(iBlockEnd+Len(sBlockEndStr)-1))
		End If
	End Function

	Public Sub SetBlock(ByRef sBlockName,ByRef sBlockContent,ByRef sContent)
		sContent = Replace(sContent & "", P_Prefix & sBlockName & "s" & P_Suffix, sBlockContent & VBCrlf & P_Prefix & sBlockName & "s" & P_Suffix)
	End Sub

	Public Sub CloseBlock(ByRef sBlockName,ByRef sContent)
		sContent = Replace(sContent & "", P_Prefix & sBlockName & "s" & P_Suffix, "")
	End Sub

	Public Sub SetVariable(ByRef sVariableName,ByRef sVariableContent,ByRef sContent)
		If InStr(sContent, P_Prefix & sVariableName & P_Suffix) > 0 Then sContent = Replace(sContent & "", P_Prefix & sVariableName & P_Suffix, sVariableContent & "")
	End Sub

	Public Function GetParameter(ParameterName, ByRef PageStr)
		Dim TempStr, PageLen
		Dim CurrentTag, StartTag, EndTag
		Dim ParameterArray
		Dim ParameterPrefix, ParameterSuffix

		CurrentTag	= 0
		StartTag	= 1
		PageLen		= Len(PageStr)

		ParameterPrefix = P_Prefix & ParameterName & "("
		ParameterSuffix = ")" & P_Suffix

		CurrentTag	= InStr(StartTag, PageStr, ParameterPrefix)

		If CurrentTag <> 0 Then
			StartTag = CurrentTag
			EndTag	 = InStr(StartTag, PageStr, ParameterSuffix)

			If EndTag <> 0 Then
				TempStr = Mid(PageStr, StartTag + Len(ParameterPrefix), EndTag - StartTag - Len(ParameterPrefix))

				ParameterArray = Split(TempStr, ",")

				PageStr = Left(PageStr, StartTag - 1) & Right(PageStr, (Len(PageStr) - EndTag - Len(ParameterSuffix)) + 1)

				LastPosition = StartTag - 1
			End If
		End If
		
		GetParameter = ParameterArray
	End Function

	Public Sub SetSingle(ByRef Content, ByRef PageStr)
		PageStr = Left(PageStr, LastPosition) & Content & Right(PageStr, Len(PageStr) - LastPosition)
	End Sub

	Public Function ChkTag_Prefix (sTag, ByRef sPageContent)
		If InStr(sPageContent, P_Prefix & sTag & ".") > 0 Then
			ChkTag_Prefix = True
		Else
			ChkTag_Prefix = False
		End If
	End Function

	Public Function ChkTag (sTag, ByRef sPageContent)
		If InStr(sPageContent, P_Prefix & sTag) > 0 Then
			ChkTag = True
		Else
			ChkTag = False
		End If
	End Function
	
	Public Function Replace_PublicTag(ByRef PageContent)
		PageContent = PageContent & ""

		EA_Pub.SysInfo(16) = Replace(EA_Pub.Full_HTMLFilter(EA_Pub.SysInfo(16)), "|", ",")
		EA_Pub.SysInfo(17) = EA_Pub.Full_HTMLFilter(EA_Pub.SysInfo(17))

		SetVariable "Page.Head", PageArray(2), PageContent
		SetVariable "Page.Foot", PageArray(3), PageContent
		SetVariable "Page.CSS", PageArray(1), PageContent

		SetVariable "Page.Title", Title, PageContent
		SetVariable "Page.Keyword", EA_Pub.SysInfo(16), PageContent
		SetVariable "Page.Description", EA_Pub.SysInfo(17), PageContent

		SetVariable "Page.Nav", Nav, PageContent
		SetVariable "Page.Path", SystemFolder, PageContent
		

		PageContent = Replace(PageContent, "</title>", "</title>" & Chr(13) & Chr(10) & "<meta name=""generator"" content=""NB文章系统(NBArticle) " & SysVersion & """ />", 1, -1, 0)
		

		If ChkTag_Prefix("Info", PageContent) Then Call MakeInfo(PageContent)
		If ChkTag_Prefix("Placard", PageContent) Then Call MakePlaCard(PageContent)
		If ChkTag_Prefix("Vote", PageContent) Then Call MakeVote(PageContent)
		If ChkTag_Prefix("Friend", PageContent) Then Call MakeFriend(PageContent)
		If ChkTag_Prefix("AdSense", PageContent) Then Call MakeAdSense(PageContent)
		If ChkTag_Prefix("Topic", PageContent) Then Call MakeTopic(PageContent)
		If ChkTag_Prefix("Column", PageContent) Then Call MakeColumn(PageContent)
		If ChkTag_Prefix("Comment", PageContent) Then Call MakeComment(PageContent)


		Replace_PublicTag=PageContent
	End Function
	
	Public Function Load_MemberTopPost()
		Dim TempStr,TempArray
		Dim i
		Dim ForTotal
		
		TempArray=EA_DBO.Get_MemberTopPostList()
		If IsArray(TempArray) Then
			TempStr="<table>"
			TempStr=TempStr&"<tr>"
			TempStr=TempStr&"<td>名次</td>"
			TempStr=TempStr&"<td>帐号</td>"
			TempStr=TempStr&"<td style='TEXT-ALIGN: right;'>投稿数</td>"
			TempStr=TempStr&"</tr>"

			ForTotal = UBound(TempArray,2)

			For i=0 To ForTotal
				If i>=10 Then Exit For
				TempStr=TempStr&"<tr>"
				TempStr=TempStr&"<td>"&i+1&"</td>"
				TempStr=TempStr&"<td>"&TempArray(1,i)&"</td>"
				TempStr=TempStr&"<td style='TEXT-ALIGN: right;'>"&TempArray(2,i)&"</td>"
				TempStr=TempStr&"</tr>"
			Next
			TempStr=TempStr&"</table>"
		Else
			TempStr=""
		End If
		
		Load_MemberTopPost=TempStr
	End Function

	Public Function PageList (PageCount,iCurrentPage,Url)
		Dim PageRoot				'页列表头
		Dim PageFoot				'页列表尾
		Dim OutStr
		Dim i						'输出字符串
		
		If CLng(iCurrentPage)<=0 Then iCurrentPage=1
		If CLng(iCurrentPage)>CLng(PageCount) Then iCurrentPage=PageCount
		
		If iCurrentPage-4<=1 Then 
			PageRoot=1
		Else
			PageRoot=iCurrentPage-4
		End If	
		If iCurrentPage+4>=PageCount Then 
			PageFoot=PageCount
		Else
			PageFoot=iCurrentPage+4
		End If
		
		OutStr="<div id=""pageNav""><span class=""total"">" & Replace(SysMsg(12), "$1", PageCount) & "</span>&nbsp;"
		
		If iCurrentPage > 1 Then 
			OutStr=OutStr&"<a href="""
			OutStr=OutStr&Replace(Url, "$page", "1")
			OutStr=OutStr&""" title=""" & SysMsg(13) & """ class=""first"">&laquo;</a>&nbsp;"
			OutStr=OutStr&"<a href="""
			OutStr=OutStr&Replace(Url, "$page", iCurrentPage-1)
			OutStr=OutStr&""" title=""" & SysMsg(14) & """ class=""list"">&lt;</a>&nbsp;"
		End If
		
		For i=PageRoot To PageFoot
			If i=Cint(iCurrentPage) Then
				OutStr=OutStr&"<span class=""current"">"&i&"</span>&nbsp;"
			Else
				OutStr=OutStr&"<a href="""
				OutStr=OutStr&Replace(Url, "$page", Cstr(i))
				OutStr=OutStr&""" class=""list"">"&i&"</a>&nbsp;"
			End If
			If i=PageCount Then Exit For
		Next

		If CInt(iCurrentPage) <> CInt(PageCount) Then
			OutStr=OutStr&"<a href="""
			OutStr=OutStr&Replace(Url, "$page", iCurrentPage+1)
			OutStr=OutStr&""" title=""" & SysMsg(15) & """ class=""list"">&gt;</a>&nbsp;"
			OutStr=OutStr&"<a href="""
			OutStr=OutStr&Replace(Url, "$page", PageCount)
			OutStr=OutStr&""" title=""" & SysMsg(16) & """ class=""last"">&raquo;</a>&nbsp;"
		End If
		
		If PageCount > 1 Then
			OutStr=OutStr&"&nbsp;<input type=""text"" value="""&iCurrentPage&""" onmouseover=""this.focus();this.select();"" id=""PGNumber"" style=""width: 30px;"" onKeypress=""if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;"" />&nbsp;<input type=""button"" value=""" & SysMsg(17) & """ onclick=""if ($('PGNumber').value>0 && $('PGNumber').value<="&PageCount&"){var Url = '" & Url & "';window.location=Url.replace('$page', $('PGNumber').value);}"" />"
		End If

		OutStr = OutStr & "</div>"

		PageList=OutStr
	End Function
End Class
%>