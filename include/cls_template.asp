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
'= 最后日期：2007-10-17
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
	'	2、页面名称
	'***********************************************
	Public Function Load_Template(TemplateId,Fields)
		FoundErr=False

		Dim Temp
		Temp=EA_DBO.Get_Template_Info(TemplateId,Fields)
		If IsArray(Temp) Then 
			PageArray(0)=Temp(0,0)			'template name
			PageArray(1)=Temp(1,0)			'template css
			PageArray(2)=Temp(2,0)			'template head
			PageArray(3)=Temp(3,0)			'template foot
			Load_Template=Temp(4,0)
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
				Case "ColumnNav"
					ReplaceStr=Load_ColumnList(ParameterArray)
				Case "DisList"
					ReplaceStr=Load_DisArticle(ParameterArray)
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
	
	Public Function Load_NewReview(Parameter)
		Dim TempStr,TempArray
		Dim i
		Dim ForTotal
		
		TempArray=EA_DBO.Get_Review_NewList(Parameter(0),Parameter(1))
		If IsArray(TempArray) Then 
			TempStr="<table>"
			ForTotal = UBound(TempArray,2)

			For i=0 To ForTotal
				TempStr=TempStr&"<tr>"
				TempStr=TempStr&"<td style=""TEXT-ALIGN: left;""><a href=""review.asp?articleid="&TempArray(0,i)&""">"&EA_Pub.Un_Full_HTMLFilter(TempArray(1,i))&"</a></td>"
				TempStr=TempStr&"<td style=""TEXT-ALIGN: center;"">"&TempArray(2,i)&"</td>"
				TempStr=TempStr&"<td style=""COLOR: #800000;TEXT-ALIGN: center;"">"&TempArray(3,i)&"</font></td>"
				TempStr=TempStr&"</tr>"
			Next
			TempStr=TempStr&"</table>"
		Else
			TempStr=""
		End If
		
		Load_NewReview=TempStr
	End Function
	
	Private Function Load_AdSense(Parameter)
		Dim TempStr,Temp
		
		Temp=EA_DBO.Get_AdSense(Parameter(0))
		If IsArray(Temp) Then TempStr=Temp(1,0)
		
		Load_AdSense=TempStr
	End Function
	
	Private Function Load_Placard(Parameter)
		Dim TempArray,TempStr
		Dim i
		Dim CutStr
		Dim ForTotal
		
		If Parameter(1)=0 Then 
			CutStr="&nbsp;&nbsp;&nbsp;"
		Else
			CutStr="<br />"
		End If
		
		If CLng(Parameter(0))=0 Then Parameter(0)=10
		
		TempArray=EA_DBO.Get_PlacardTopList(Parameter(0))
		If IsArray(TempArray) Then 
			ForTotal = UBound(TempArray,2)

			For i=0 To ForTotal
				TempStr=TempStr&"<img src="""&SystemFolder&"images/public/gb.gif"" alt="""" />&nbsp;<a href=""#"" onclick=""javascript:window.open('"&SystemFolder&"viewplacard.asp?postid="&TempArray(0,i)&"','','scrollbars=yes,height=350,width=550')"">"&TempArray(1,i)&"</a>&nbsp;<font color=""#999999"">("&FormatDateTime(TempArray(2,i),2)&")</font>"&CutStr
			Next
		Else
			TempStr=TempStr&"<font color=""#800000"">欢迎光临"&EA_Pub.SysInfo(0)&"。</font>"
		End If
		
		Load_Placard=TempStr
	End Function
	
	Private Function Load_Vote(Parameter)
		Dim Id,Title,VoteText,Mtype
		Dim Result,Content,j
		Dim TempArray
		Dim ForTotal
		
		TempArray=EA_DBO.Get_Vote_Info(Parameter(0))
		If IsArray(TempArray) Then 
			If TempArray(5,0)=0 Then
				Id=TempArray(0,0)
				Title=TempArray(1,0)
				VoteText=TempArray(2,0)
				Mtype=TempArray(4,0)
				
				Content=split(VoteText,"|")
				
				Result=Result&"<form action="""&SystemFolder&"vote.asp"" method=""post"" id=""vote_"&Id&"""><table><tr><td>"&Title&"</td></tr>"
		
				IF Mtype=False Then
					ForTotal = Ubound(Content)

					For i=0 To ForTotal
						Result=Result&"<tr><td>"
						Result=Result&"<input type=""radio"" name=""vote"" value="""&i&""" />&nbsp;&nbsp;&nbsp;"&Content(i)
						Result=Result&"</td></tr>"
					Next
				End if

				If Mtype=True Then
					ForTotal = Ubound(Content)

					For i=0 To ForTotal
						Result=Result&"<tr><td>"
						Result=Result&"<input type=""checkbox"" name=""vote"" value="""&i&""" />&nbsp;&nbsp;&nbsp;"&Content(i)
						Result=Result&"</td></tr>"
					Next
				End if
		
				Result=Result&"<tr><td><input name=""votetype"" id=""votetype"" type=""hidden"" value="""&CInt(Mtype)&""" /><input name=""voteid"" id=""voteid"" type=""hidden"" value="""&Id&""" /><input type=""button"" name=""submit"" value=""投票"" onclick=""window.open(submit_vote("&Id&"),'_blank','scrollbars=yes,width=645,height=380')"" />&nbsp;<input type=""button"" name=""view"" value=""查看"" onclick=""window.open('vote.asp?VoteId="&ID&"','_blank','scrollbars=yes,width=645,height=380')"" /></td></tr>"
				Result=Result&"</table></form>"
			End If
		End If
		
		Load_Vote=Result
	End Function

	Private Function Load_Friend(Parameter)
		Dim FriendList,WidthPercent
		Dim TempStr,i
		Dim ForTotal
		
		WidthPercent=100/CLng(Parameter(3))

		FriendList=EA_DBO.Get_Friend_List(Parameter(1),Parameter(0),Parameter(2))
		
		If IsArray(FriendList) Then 
			TempStr="<table>"
			TempStr=TempStr&"<tr>"
			ForTotal = UBound(FriendList,2)

			For i=0 To ForTotal
				If Parameter(2)="1" Then 
					TempStr=TempStr&"<td style""WIDTH: "&WidthPercent&"%;""><a href="""&FriendList(1,i)&""" rel=""external""><img src="""&FriendList(2,i)&""" alt="""&FriendList(0,i)&""" /></a></td>"
				Else
					TempStr=TempStr&"<td style""WIDTH: "&WidthPercent&"%;""><a href="""&FriendList(1,i)&""" rel=""external"" title="""&FriendList(0,i)&""">"&FriendList(0,i)&"</a></td>"
				End If

				If (i+1) Mod CLng(Parameter(3))=0 Then TempStr=TempStr&"</tr><tr>"
			Next
			If i Mod CLng(Parameter(3))=0 Then TempStr=TempStr&"</tr>"

			TempStr=TempStr&"</table>"
		Else
			TempStr = ""
		End If

		Load_Friend=TempStr
	End Function
	
	Private Function Load_DisArticle(Parameter)
		Dim TempStr,TempArray,i,IsCrlf
		Dim ForTotal

		TempArray=EA_DBO.Get_DisColumn(Parameter(1),Parameter(0))
		
		If IsArray(TempArray) Then 
			TempStr="<table>"
			TempStr=TempStr&"<tr>"

			IsCrlf = 0
			ForTotal = UBound(TempArray,2)

			For i=0 To ForTotal
				TempStr=TempStr&"<td><a href=""" & EA_Pub.Cov_ColumnPath(TempArray(0,i),EA_Pub.SysInfo(18)) & """>"&TempArray(1,i)&"</a></td>"
				TempStr=TempStr&"<td>"
				Select Case Parameter(2)
				Case "1"
					TempStr=TempStr&"&nbsp;文章总数："&TempArray(2,i)
				Case "2"
					TempStr=TempStr&"&nbsp;今日更新："&TempArray(3,i)
				End Select
				TempStr=TempStr&"</td>"
				
				If (i+1) Mod CLng(Parameter(3))=0 Then 
					TempStr=TempStr&"</tr>"

					If (i+1) <= UBound(TempArray,2) Then TempStr=TempStr&"<tr>":IsCrlf = 1
				End If
			Next
			If i Mod CLng(Parameter(3))=0 And IsCrlf = 0 Then TempStr=TempStr&"</tr>"
			
			TempStr=TempStr&"</table>"
		Else
			TempStr = ""
		End If
		
		Load_DisArticle=TempStr
	End Function
	
	Private Function Load_ColumnList(Parameter)
		Dim TempStr,ds8,i,j,MainId,URL,URL2,IsCrlf
		Dim ForTotal

		ds8=EA_DBO.Get_Column_List()
		
		If IsArray(ds8) Then
			TempStr="<table>"
			ForTotal = Ubound(ds8,2)

			For i=0 To ForTotal
				URL=EA_Pub.Cov_ColumnPath(ds8(0,i),EA_Pub.SysInfo(18))
				URL2=EA_Pub.Cov_ColumnPath(MainId,EA_Pub.SysInfo(18))

				If Parameter(0)="0" Then 
					If Len(ds8(2,i))=4 Then
						TempStr=TempStr&"<tr>"
						TempStr=TempStr&"<td>&nbsp;*&nbsp;<a href="""&URL&""" title="""&ds8(3,i)&"""><strong>"&ds8(1,i)&"</strong></a></td>"
						TempStr=TempStr&"</tr>"

						If i+1<=UBound(ds8,2) Then
							If Len(ds8(2,i+1))>4 Then
								TempStr=TempStr&"<tr>"
								TempStr=TempStr&"<td>&nbsp;</td>"
								TempStr=TempStr&"<td><table>"
								TempStr=TempStr&"<tr style=""HEIGHT: 20px;"">"
							End If
						End If
						j=1

						IsCrlf = 0
					Else
						TempStr=TempStr&"<td><a href="""&URL&""" title="""&ds8(3,i)&""">"
						TempStr=TempStr&ds8(1,i)&"</a></td>"
						j=j+1

						If j>=CInt(Parameter(1))+1 Then TempStr=TempStr&"</tr><tr><td colspan="""&Parameter(1)&""" style=""height: 1px;BACKGROUND: #000;""></td></tr>":j=1:IsCrlf = 1

						If i+1<=UBound(ds8,2) Then
							If Len(ds8(2,i+1))=4 Then
								If IsCrlf = 0 Then TempStr=TempStr&"</tr>"

								TempStr=TempStr&"</table></td></tr>"
								j=0
							ElseIf IsCrlf = 1 Then
								TempStr=TempStr&"<tr>"
								IsCrlf = 0
							End if
						ElseIf i+1>UBound(ds8,2) Then 
							If IsCrlf = 0 Then TempStr=TempStr&"</tr>"

							TempStr=TempStr&"</table></td></tr>"
						End If
					End If
				Else
					If Len(ds8(2,i))=4 Then
						MainId=ds8(0,i)
						TempStr=TempStr&"<tr>"
						TempStr=TempStr&"<td style=""WIDTH: 25%;"">&nbsp;*&nbsp;<a href="""&URL&""" title="""&ds8(3,i)&"""><strong>"&ds8(1,i)&"</strong></a></td>"
						j=1
						
						If i+1<=UBound(ds8,2) Then
							If Len(ds8(2,i+1))=4 Then 
								TempStr=TempStr&"</tr>"
							Else
								TempStr=TempStr&"<td>"
							End If
						Else
							TempStr=TempStr&"<td></td></tr>"
						End If
					Else
						If j<>0 Then 
							TempStr=TempStr&"<a href="""&URL&""" title="""&ds8(3,i)&""">"
							TempStr=TempStr&ds8(1,i)&"</a>&nbsp;|&nbsp;"
							j=j+1
						
							If j-1>=CInt(Parameter(1)) Then 
								TempStr=TempStr&"<a href="""&URL2&""">更多&gt;&gt;</a></td></tr>"
								j=0
							Else
								If (i+1) > UBound(ds8,2) Then 
									TempStr=TempStr&"</td></tr>"
								Else
									If Len(ds8(2,i+1))=4 Then TempStr=TempStr&"</td></tr>"
								End If
							End If
						End If
					End If
				End If			
			Next
			
			TempStr=TempStr&"</table>"
		Else
			TempStr = ""
		End If
		
		Load_ColumnList=TempStr
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

				TempStr=TempStr&"><img src="""&DataArray(8,i)&""" alt="""&DataArray(3,i)&""" width=""90"" class=""midImg"" /></a></td></tr>"

				If TitleLen > 1 Then
					TempStr=TempStr&"<tr><td style=""HEIGHT: 22px;"">"
					
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
			OutStr=OutStr&"&nbsp;<input type=""text"" value="""&iCurrentPage&""" onmouseover=""this.focus();this.select()"" id=""PGNumber"" style=""width: 30px;"" onKeypress=""if (event.keyCode < 45 || event.keyCode > 57) event.returnValue = false;"" />&nbsp;<input type=""button"" value=""GO"" onclick=""if (document.getElementById('PGNumber').value>0 && document.getElementById('PGNumber').value<="&PageCount&"){window.location='?page='+document.getElementById('PGNumber').value+'"&Url&"'}"" />"
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