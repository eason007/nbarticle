<!--#Include File="tag_friend.asp"-->
<!--#Include File="tag_placard.asp"-->
<!--#Include File="tag_review.asp"-->
<!--#Include File="tag_vote.asp"-->
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
'= 最后日期：2008-02-25
'====================================================================

Class cls_Template
	Public Title,Nav

	Private PageArray(4)
	Private i
	
	'*****************************
	'对象类初始化过程
	'*****************************
	Private Sub Class_Initialize()

	End Sub

	Public Sub Close_Obj()
		Erase PageArray
	End Sub
	
	'***********************************************
	'加载模版过程
	'输入参数：
	'	1、模版id
	'***********************************************
	Public Function Load_Template(TemplateId, TemplateType)
		FoundErr=False

		Dim Temp
		Temp=EA_DBO.Get_Template_Info(TemplateId, TemplateType)
		If IsArray(Temp) Then 
			'PageArray(0)=EA_DBO.Get_Template_Info(TemplateId, )(0,0)			'template name
			PageArray(1)=EA_DBO.Get_Template_Info(0, 1)(0,0)			'template css
			PageArray(2)=EA_DBO.Get_Template_Info(0, 2)(0,0)			'template head
			PageArray(3)=EA_DBO.Get_Template_Info(0, 3)(0,0)			'template foot
			Load_Template=Temp(0,0)
		Else
			FoundErr=True
		End If
		
		If Err.Number<>0 Or FoundErr Then 
			ErrMsg="读取模版["&Fields&"]时发生错误"&FoundErr
			Call EA_Pub.ShowErrMsg(0,0)
		End If
	End Function

	Public Sub ReplaceTag(sTag,sReplaceValue,ByRef sPageContent)
		If InStr(sPageContent,"{$"&sTag&"$}")>0 Then
			sPageContent=Replace(sPageContent,"{$"&sTag&"$}",sReplaceValue&"")
		End If
	End Sub
	
	Public Function Replace_PublicTag(ByRef PageContent)
		PageContent = PageContent & ""

		ReplaceTag "Head",PageArray(2),PageContent
		ReplaceTag "Foot",PageArray(3),PageContent

		ReplaceTag "PageTitle",Title,PageContent
		ReplaceTag "PageNav",Nav,PageContent

		ReplaceTag "SiteName",EA_Pub.SysInfo(0),PageContent
		ReplaceTag "SiteUrl",EA_Pub.SysInfo(11),PageContent
		ReplaceTag "SiteEMail",EA_Pub.SysInfo(12),PageContent
		ReplaceTag "SystemVersion",SysVersion,PageContent
		ReplaceTag "SkinName",PageArray(0),PageContent

		EA_Pub.SysInfo(17) = Replace(EA_Pub.SysInfo(17) & "", "<", "&lt;")
		EA_Pub.SysInfo(17) = Replace(EA_Pub.SysInfo(17) & "", ">", "&gt;")
		EA_Pub.SysInfo(17) = Replace(EA_Pub.SysInfo(17) & "", "&", "&amp;")

		EA_Pub.SysInfo(16) = Replace(EA_Pub.SysInfo(16) & "", "<", "&lt;")
		EA_Pub.SysInfo(16) = Replace(EA_Pub.SysInfo(16) & "", ">", "&gt;")
		EA_Pub.SysInfo(16) = Replace(EA_Pub.SysInfo(16) & "", "&", "&amp;")
	
		ReplaceTag "PageCSS",PageArray(1),PageContent

		ReplaceTag "PageDesc",EA_Pub.SysInfo(17),PageContent
		ReplaceTag "PageKeyword",Replace(EA_Pub.SysInfo(16),"|",","),PageContent

		PageContent=Replace(PageContent,"</title>","</title>"&Chr(13)&Chr(10)&"<meta name=""generator"" content=""NB文章系统(NBArticle) "&SysVersion&""" />",1,-1,0)

		ReplaceTag "SystemPath",SystemFolder,PageContent
		
		If InStr(PageContent,"{$SitePlacard")>0 Then Call Find_TemplateTag("SitePlacard",PageContent)

		If InStr(PageContent,"{$SiteVote")>0 Then Call Find_TemplateTags("SiteVote",PageContent)
		If InStr(PageContent,"{$GetArticleList")>0 Then Call Find_TemplateTags("GetArticleList",PageContent)
		If InStr(PageContent,"{$AdSense")>0 Then Call Find_TemplateTags("AdSense",PageContent)
		If InStr(PageContent,"{$Friend")>0 Then Call Find_TemplateTags("Friend",PageContent)
		
		Replace_PublicTag=PageContent
	End Function
	
	Public Function Find_TemplateTag(KeyStr,ByRef PageStr)
		Dim TempStr,PageLen
		Dim CurrentTag,StartTag,EndTag
		Dim ParameterArray,ReplaceStr,ReplaceLen

		CurrentTag=0
		StartTag=1
		PageLen=Len(PageStr)

		CurrentTag=InStr(StartTag,PageStr,"{$"&KeyStr&"(")

		If CurrentTag<>0 Then 
			StartTag=CurrentTag
			EndTag=InStr(StartTag,PageStr,")$}")

			If EndTag <> 0 Then
				TempStr=Mid(PageStr,StartTag+3+Len(KeyStr),EndTag-(StartTag+3+Len(KeyStr)))

				ParameterArray=Split(TempStr,",")
				
				Select Case KeyStr
				Case "SitePlacard"
					ReplaceStr=Load_Placard(ParameterArray)
				Case "NewReview"
					ReplaceStr=Load_NewReview(ParameterArray)
				End Select

				ReplaceLen=Len(ReplaceStr)

				PageStr=Left(PageStr,StartTag-1)&ReplaceStr&Right(PageStr,PageLen-(EndTag+2))
			End If
		End If
		
		Find_TemplateTag=PageStr
	End Function

	Public Function Find_TemplateTags(KeyStr,ByRef PageStr)
		Dim TempStr,PageLen
		Dim CurrentTag,StartTag,EndTag
		Dim ParameterArray,ReplaceStr,ReplaceLen

		CurrentTag=-1
		StartTag=1
		PageLen=Len(PageStr)

		Do While CurrentTag<>0
			CurrentTag=InStr(StartTag,PageStr,"{$"&KeyStr&"(")

			If CurrentTag<>0 Then 
				StartTag=CurrentTag
				EndTag=InStr(StartTag,PageStr,")$}")

				If EndTag <> 0 Then
					TempStr=Mid(PageStr,StartTag+Len(KeyStr)+3,EndTag-(StartTag+Len(KeyStr)+3))

					ParameterArray=Split(TempStr,",")
					
					Select Case KeyStr
					Case "GetArticleList"
						ReplaceStr=Get_ArticleList(ParameterArray)
					Case "Friend"
						ReplaceStr=Load_Friend(ParameterArray)
					Case "AdSense"
						ReplaceStr=Load_AdSense(ParameterArray)
					Case "SiteVote"
						ReplaceStr=Load_Vote(ParameterArray)
					End Select

					ReplaceLen=Len(ReplaceStr)

					PageStr=Left(PageStr,StartTag-1)&ReplaceStr&Right(PageStr,PageLen-(EndTag+2))

					StartTag=StartTag+ReplaceLen

					PageLen=Len(PageStr)
				End If
			End If
		Loop
		
		Find_TemplateTags=PageStr
	End Function

	Public Function Find_TemplateTagValues(KeyStr,ByRef PageStr)
		Dim TempStr,PageLen
		Dim CurrentTag,StartTag,EndTag
		Dim ParameterArray

		CurrentTag=0
		StartTag=1
		PageLen=Len(PageStr)

		CurrentTag=InStr(StartTag,PageStr,"{$"&KeyStr&"(")

		If CurrentTag<>0 Then
			StartTag=CurrentTag
			EndTag=InStr(StartTag,PageStr,")$}")

			If EndTag <> 0 Then
				TempStr=Mid(PageStr,StartTag+3+Len(KeyStr),EndTag-(StartTag+3+Len(KeyStr)))

				ParameterArray=Split(TempStr,",")
			End If
		End If
		
		Find_TemplateTagValues=ParameterArray
	End Function

	Public Sub Find_TemplateTagByInput(KeyStr,ReplaceStr,ByRef PageStr)
		Dim PageLen
		Dim CurrentTag,StartTag,EndTag
		Dim ReplaceLen

		CurrentTag=0
		StartTag=1
		PageLen=Len(PageStr)

		CurrentTag=InStr(StartTag,PageStr,"{$"&KeyStr&"(")
		If CurrentTag<>0 Then 
			StartTag=CurrentTag
			EndTag=InStr(StartTag,PageStr,")$}")

			ReplaceLen=Len(ReplaceStr)

			PageStr=Left(PageStr,StartTag-1)&ReplaceStr&Right(PageStr,PageLen-(EndTag+2))
		End If
	End Sub
	
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
	
	Private Function Load_AdSense(Parameter)
		Dim TempStr,Temp
		
		Temp=EA_DBO.Get_AdSense(Parameter(0))
		If IsArray(Temp) Then TempStr=Temp(1,0)
		
		Load_AdSense=TempStr
	End Function

	Private Function Get_ArticleList(Parameter)
		Dim TempArray

		If UBound(Parameter)=12 Then
			TempArray=EA_DBO.Get_Article_List(Parameter(2),Parameter(0),Parameter(1),Parameter(12))
			If IsArray(TempArray) Then 
				If Parameter(6)="0" Then 
					Get_ArticleList=Text_List(TempArray,CInt(Parameter(4)),CInt(Parameter(5)),CInt(Parameter(3)),CInt(Parameter(7)),CInt(Parameter(8)),CInt(Parameter(9)),CInt(Parameter(10)),CInt(Parameter(11)),CInt(Parameter(2))-1)
				Else
					Get_ArticleList=Img_List(TempArray,CInt(Parameter(4)),CInt(Parameter(5)),CInt(Parameter(3)),CInt(Parameter(7)),CInt(Parameter(8)),CInt(Parameter(9)),CInt(Parameter(10)),CInt(Parameter(2))-1)
				End If
			Else
				Get_ArticleList="&nbsp;·暂无"
			End If
		Else
			Get_ArticleList="&nbsp;调用参数不足，现有 "&UBound(Parameter)+1&" 个，需 13 个"
		End If
	End Function
	
	Public Function Text_List(DataArray,IsShowSort,IsShowDate,TitleLen,RowNum,IsShowNewTag,IsNewTarget,IsShowAuthor,IsShowFileType,ListTotal)
		Dim TempStr
		Dim IsCrlf
		Dim RowSize
		
		If IsArray(DataArray) And RowNum>0 Then 
			If ListTotal>UBound(DataArray,2) Then ListTotal=UBound(DataArray,2)

			IsCrlf = 0
			RowSize = 100/RowNum

			TempStr="<table>"
			TempStr=TempStr&"<tr>"
			For i=0 To ListTotal
				If IsCrlf = 1 Then TempStr=TempStr&"<tr>":IsCrlf = 0

				TempStr=TempStr&"<td style=""width: " & RowSize & "%;"">"
				
				If IsShowSort=1 Then TempStr=TempStr&"[<a href="""&EA_Pub.Cov_ColumnPath(DataArray(1,i),EA_Pub.SysInfo(18))&""" rel=""external"" class=""link-Column"">"&DataArray(2,i)&"</a>]&nbsp;"
				
				If IsShowFileType=1 Then TempStr=TempStr&EA_Pub.Chk_ArticleType(DataArray(6,i),DataArray(7,i))&"&nbsp;"
				
				TempStr=TempStr&"<a href="""&EA_Pub.Cov_ArticlePath(DataArray(0,i),DataArray(5,i),EA_Pub.SysInfo(18))&""""

				If IsNewTarget=1 Then TempStr=TempStr&" rel=""external"""

				TempStr=TempStr&" title=""" & EA_Pub.Base_HTMLFilter(DataArray(3,i)) & """>"
				DataArray(3,i)=EA_Pub.Base_HTMLFilter(DataArray(3,i))
				DataArray(3,i)=EA_Pub.Cut_Title(DataArray(3,i),TitleLen)
				TempStr=TempStr&EA_Pub.Add_ArticleColor(DataArray(4,i),DataArray(3,i))
				TempStr=TempStr&"</a>"

				If IsShowNewTag=1 Then TempStr=TempStr&EA_Pub.Chk_ArticleTime(DataArray(5,i))

				If IsShowAuthor=1 Then 
					If Len(DataArray(9,i))>0 Then TempStr=TempStr&"&nbsp;[<span class=""link-Author"">"&DataArray(9,i)&"</span>]"
				End If

				TempStr=TempStr&"</td>"
				If IsShowDate=1 Then 
					TempStr=TempStr&"<td style=""TEXT-ALIGN: right;"">"
					TempStr=TempStr&"<span class=""link-Date"">["&Month(DataArray(5,i))&"."&Day(DataArray(5,i))&"]</span>"
					TempStr=TempStr&"</td>"
				End If

				If (i+1) Mod RowNum=0 Then TempStr=TempStr&"</tr>":IsCrlf = 1
			Next
			If (i-1) Mod RowNum<>0 And IsCrlf = 0 Then TempStr=TempStr&"</tr>"
			TempStr=TempStr&"</table>"
		ElseIf RowNum<=0 Then
			TempStr="列数设置为0，请修改"
		Else 
			TempStr="·暂无"
		End If
	
		Text_List=TempStr
	End Function
	
	Public Function Img_List(DataArray,IsShowSort,IsShowDate,TitleLen,RowNum,IsShowNewTag,IsNewTarget,IsShowAuthor,ListTotal)
		Dim TempStr
		
		If IsArray(DataArray) And RowNum>0 Then 
			If ListTotal>UBound(DataArray,2) Then ListTotal=UBound(DataArray,2)

			TempStr="<table>"
			TempStr=TempStr&"<tr>"
			For i=0 To ListTotal
				TempStr=TempStr&"<td><table>"
				TempStr=TempStr&"<tr><td>"
				TempStr=TempStr&"<a href="""&EA_Pub.Cov_ArticlePath(DataArray(0,i),DataArray(5,i),EA_Pub.SysInfo(18))&""""
				
				If IsNewTarget=1 Then TempStr=TempStr&" rel=""external"""

				TempStr=TempStr&"><img src="""&DataArray(8,i)&""" alt="""&DataArray(3,i)&""" class=""midImg"" /></a></td></tr>"

				If TitleLen > 1 Then
					TempStr=TempStr&"<tr><td>"
					
					If IsShowSort=1 Then TempStr=TempStr&"&nbsp;[<a href="""&EA_Pub.Cov_ColumnPath(DataArray(1,i),EA_Pub.SysInfo(18))&""" rel=""external"" class=""link-Column"">"&DataArray(2,i)&"</a>]"

					TempStr=TempStr&"&nbsp;<a href="""&EA_Pub.Cov_ArticlePath(DataArray(0,i),DataArray(5,i),EA_Pub.SysInfo(18))&""""
					
					If IsNewTarget=1 Then TempStr=TempStr&" rel=""external"""
					
					TempStr=TempStr&" title="""&EA_Pub.Base_HTMLFilter(DataArray(3,i))&""">"
					DataArray(3,i)=EA_Pub.Base_HTMLFilter(DataArray(3,i))
					DataArray(3,i)=EA_Pub.Cut_Title(DataArray(3,i),TitleLen)
					TempStr=TempStr&EA_Pub.Add_ArticleColor(DataArray(4,i),DataArray(3,i))
					TempStr=TempStr&"</a>"

					If IsShowNewTag=1 Then TempStr=TempStr&EA_Pub.Chk_ArticleTime(DataArray(5,i))

					If IsShowAuthor=1 Then 
						If Len(DataArray(9,i))>0 Then TempStr=TempStr&"&nbsp;[<span class=""link-Author"">"&DataArray(9,i)&"</span>]"
					End If
					
					If IsShowDate=1 Then TempStr=TempStr&"&nbsp;<span class=""link-Date"">"&Month(DataArray(5,i))&"/"&Day(DataArray(5,i))&"</span>"
					
					TempStr=TempStr&"</td></tr>"
				End If
				
				TempStr=TempStr&"</table></td>"
				
				If (i+1) Mod RowNum=0 Then TempStr=TempStr&"</tr>"
			Next
			If i Mod RowNum<>0 Then TempStr=TempStr&"</tr>"
			TempStr=TempStr&"</table>"
		ElseIf RowNum<=0 Then
			TempStr="列数设置为0，请修改"
		Else 
			TempStr="·暂无"
		End If
		
		Img_List=TempStr
	End Function

	Public Function PageList (PageCount,iCurrentPage,FieldName,FieldValue)
		Dim Url
		Dim PageRoot				'页列表头
		Dim PageFoot				'页列表尾
		Dim OutStr
		Dim i						'输出字符串
		
		Url=URLStr(FieldName,FieldValue)
		
		If CLng(iCurrentPage)<=0 Then 
			iCurrentPage=1
		ElseIf CLng(iCurrentPage)>CLng(PageCount) Then
			iCurrentPage=PageCount
		End if
		
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
		
		OutStr="<div id=""pageList""><span class=""total"">共 "&PageCount&" 页</span>&nbsp;"
		
		If iCurrentPage > 1 Then 
			OutStr=OutStr&"<a href=""?page=1"
			OutStr=OutStr&Url
			OutStr=OutStr&""" title=""首页"" class=""first"">&laquo;</a>&nbsp;"
			OutStr=OutStr&"<a href=""?page="&iCurrentPage-1
			OutStr=OutStr&Url
			OutStr=OutStr&""" title=""上页"" class=""list"">&lt;</a>&nbsp;"
		End If
		
		For i=PageRoot To PageFoot
			If i=Cint(iCurrentPage) Then
				OutStr=OutStr&"<span class=""current"">"&i&"</span>&nbsp;"
			Else
				OutStr=OutStr&"<a href=""?page="&Cstr(i)
				OutStr=OutStr&Url
				OutStr=OutStr&""" class=""list"">"&i&"</a>&nbsp;"
			End If
			If i=PageCount Then Exit For
		Next

		If CInt(iCurrentPage) <> CInt(PageCount) Then
			OutStr=OutStr&"<a href=""?page="&iCurrentPage+1
			OutStr=OutStr&Url
			OutStr=OutStr&""" title=""下页"" class=""list"">&gt;</a>&nbsp;"
			OutStr=OutStr&"<a href=""?page="&PageCount
			OutStr=OutStr&Url
			OutStr=OutStr&""" title=""尾页"" class=""last"">&raquo;</a>&nbsp;"
		End If
		
		If PageCount > 0 Then
			OutStr=OutStr&"&nbsp;<input type=""text"" value="""&iCurrentPage&""" onmouseover=""this.focus();this.select();"" id=""PGNumber"" style=""width: 30px;"" onKeypress=""if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;"" />&nbsp;<input type=""button"" value=""GO"" onclick=""if ($('PGNumber').value>0 && $('PGNumber').value<="&PageCount&"){window.location='?page='+$('PGNumber').value+'"&Url&"'}"" />"
		End If

		OutStr = OutStr & "</div>"

		PageList=OutStr
	End Function

	Private Function URLStr(FieldName,FieldValue)
		If IsArray(FieldName) And IsArray(FieldValue) Then 
			Dim i
			Dim ForTotal

			ForTotal = Ubound(FieldName)

			For i = 0 To ForTotal
				URLStr=URLStr&"&amp;"&Cstr(FieldName(i))&"="&Cstr(FieldValue(i))
			Next
		End If
	End Function
End Class
%>