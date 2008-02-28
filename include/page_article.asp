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
'= 最后日期：2008-02-26
'====================================================================

Class page_Article
	Public PageIndex(), PageStr()


	Public Function Make (ID, Info, Page)
		Dim FirstArticle,NextArticle
		Dim i,TempStr

		PageContent=EA_Temp.Load_Template(ArticleInfo(24,0),"view")

		EA_Temp.Title=ArticleInfo(3,0)&" - "&ArticleInfo(2,0)&" - "&EA_Pub.SysInfo(0)
		EA_Temp.Nav="<a href="""&SystemFolder&"""><b>"&EA_Pub.SysInfo(0)&"</b></a>"&EA_Pub.Get_NavByColumnCode(ArticleInfo(1,0))&" -=> 正文"

		ArticleInfo(5,0)=EA_Pub.Cov_InsideLink(ArticleInfo(5,0),ArticleInfo(0,0))

		If Not IsView Then 
			TempStr=ArticleInfo(4,0)
			TempStr=TempStr&"<br><br><b>您当前的权限不允许查看该文章，请先 [<a href='"&SystemFolder&"member/login.asp' rel=""external"">登陆</a>] 或 [<a href='"&SystemFolder&"member/register.asp' rel=""external"">注册</a>]。</b>"
		Else
			Call RegExpTest("\[NextPage([^\]])*\]", ArticleInfo(5,0))

			If UBound(PageIndex) = 1 Then
				TempStr="<div id=""article"">"&ArticleInfo(5,0)&"</div>"
			Else
				TempStr = Mid(ArticleInfo(5,0), PageIndex(Page - 1) + Len(PageStr(Page - 1)) + 1, PageIndex(Page) - PageIndex(Page - 1) - Len(PageStr(Page - 1)))
				TempStr = "<div id=""article"">" & TempStr & "</div>"
				TempStr = TempStr & "<div style='TEXT-ALIGN: center;margin-bottom: 5px;'>" & PageNav(UBound(PageIndex), Page) & "</div>"
			End If
		End If

		PageContent=Replace(PageContent,"{$ColumnId$}",ArticleInfo(0,0))
		PageContent=Replace(PageContent,"{$ArticleId$}",ArticleId)
		PageContent=Replace(PageContent,"{$ArticleTitle$}",EA_Pub.Add_ArticleColor(ArticleInfo(17,0),ArticleInfo(3,0)))
		PageContent=Replace(PageContent,"{$ArticlePostTime$}",ArticleInfo(13,0))
		PageContent=Replace(PageContent,"{$ArticleText$}",TempStr)
		PageContent=Replace(PageContent,"{$ArticleSummary$}",ArticleInfo(4,0))

		PageContent=Replace(PageContent,"{$ArticleAuthor$}","<a href='"&SystemFolder&"florilegium.asp?a_name="&ArticleInfo(8,0)&"&amp;a_id="&ArticleInfo(7,0)&"' rel=""external"">"&ArticleInfo(8,0)&"</a>")

		If Len(ArticleInfo(16,0))>0 Then
			PageContent=Replace(PageContent,"{$ArticleFrom$}","<a href='"&ArticleInfo(16,0)&"' rel=""external"">"&ArticleInfo(15,0)&"</a>")
		Else
			PageContent=Replace(PageContent,"{$ArticleFrom$}","本站")
		End If

		If InStr(PageContent, "{$FirstArticle$}") > 0 Then
			FirstArticle=EA_DBO.Get_Article_FirstArticle(ArticleInfo(0,0),ArticleInfo(25,0),ArticleId)

			If IsArray(FirstArticle) Then
				PageContent=Replace(PageContent,"{$FirstArticle$}","<a href='"&EA_Pub.Cov_ArticlePath(FirstArticle(0,0),FirstArticle(3,0),EA_Pub.SysInfo(18))&"' rel=""external"">"&EA_Pub.Add_ArticleColor(FirstArticle(2,0),FirstArticle(1,0))&"</a>")
			Else
				PageContent=Replace(PageContent,"{$FirstArticle$}","<span style=""color: #800000;"">已到尽头</span>")
			End If
		End If

		If InStr(PageContent, "{$NextArticle$}") > 0 Then
			NextArticle=EA_DBO.Get_Article_NextArticle(ArticleInfo(0,0),ArticleInfo(25,0),ArticleId)

			If IsArray(NextArticle) Then
				PageContent=Replace(PageContent,"{$NextArticle$}","<a href='"&EA_Pub.Cov_ArticlePath(NextArticle(0,0),NextArticle(3,0),EA_Pub.SysInfo(18))&"' rel=""external"">"&EA_Pub.Add_ArticleColor(NextArticle(2,0),NextArticle(1,0))&"</a>")
			Else
				PageContent=Replace(PageContent,"{$NextArticle$}","<span style=""color: #800000;"">已到尽头</span>")
			End If
		End If

		PageContent=Replace(PageContent,"{$ArticleViewTotal$}","<script type=""text/javascript"" src="""&SystemFolder&"articleinfo.asp?action=viewtotal&amp;articleid="&ArticleId&"""></script>")
		PageContent=Replace(PageContent,"{$ArticleCommentTotal$}","<script type=""text/javascript"" src="""&SystemFolder&"articleinfo.asp?action=commenttotal&amp;articleid="&ArticleId&"""></script>")

		EA_Pub.SysInfo(16)=ArticleInfo(12,0)&","&EA_Pub.SysInfo(16)
		EA_Pub.SysInfo(17)=ArticleInfo(4,0)

		Call CorrList(ArticleInfo(12,0),ArticleInfo(0,0))
		Call TagList(ArticleInfo(12,0))

		PageContent=EA_Temp.Replace_PublicTag(PageContent)

		Make = PageContent
	End Function


	Function TagList (Keyword)
		Dim TempArray,i
		Dim ForTotal
		Dim OutStr

		If Len(Trim(Keyword)) > 0 And Not IsNull(Keyword) Then
			TempArray= Split(Keyword,",")

			ForTotal = UBound(TempArray)

			For i=0 To ForTotal
				If Len(Trim(TempArray(i))) > 0 And Not IsNull(TempArray(i)) Then OutStr = OutStr & "<a href='" & SystemFolder & "search.asp?action=query&amp;field=1&amp;keyword=" & EA_Pub.c(Trim(TempArray(i))) & "' rel='external'>" & Trim(TempArray(i)) & "</a>&nbsp;"
			Next
		End If

		Call EA_Temp.ReplaceTag("TagList",OutStr,PageContent)
	End Function

	Function CorrList(Keyword,ColumnId)
		Dim ConfigParameterArray

		ConfigParameterArray=EA_Temp.Find_TemplateTagValues("CorrList",PageContent)

		If IsArray(ConfigParameterArray) Then 
			If UBound(ConfigParameterArray) < 8 Then 
				ReDim Preserve ConfigParameterArray(8)
				ConfigParameterArray(8) = "5"
			End If

			If Keyword <> "" Then
				Dim TempArray,i,TempStr,SearchKeyWord
				Dim ForTotal
				
				TempArray= Split(Keyword,",")
				ForTotal = UBound(TempArray)

				For i=0 To ForTotal
					Select Case iDataBaseType
					Case 0
						SearchKeyWord=SearchKeyWord&"InStr(','+keyword+',',',"&TempArray(i)&",')>0 or "
					Case 1
						SearchKeyWord=SearchKeyWord&" CharIndex(',"&TempArray(i)&",',','+keyword+',')>0 or "
					End Select
				Next

				TempArray=EA_DBO.Get_Article_CorrList(SearchKeyWord,ArticleId,ColumnId,CInt(ConfigParameterArray(8)))
			End If

			TempStr=EA_Temp.Text_List(TempArray,CInt(ConfigParameterArray(0)),CInt(ConfigParameterArray(1)),CInt(ConfigParameterArray(2)),CInt(ConfigParameterArray(3)),CInt(ConfigParameterArray(4)),CInt(ConfigParameterArray(5)),CInt(ConfigParameterArray(6)),CInt(ConfigParameterArray(7)),CInt(ConfigParameterArray(8)))

			Call EA_Temp.Find_TemplateTagByInput("CorrList",TempStr,PageContent)
		End If 
	End Function

	Function RegExpTest(patrn, strng) 
		Dim regEx, Match, Matches			' 建立变量。 
		Dim i

		Set regEx = New RegExp				' 建立正则表达式。 

		regEx.Pattern = patrn				' 设置模式。 
		regEx.IgnoreCase = True				' 设置是否区分字符大小写。 
		regEx.Global = True					' 设置全局可用性。 
		Set Matches = regEx.Execute(strng)	' 执行搜索。 

		ReDim PageIndex(Matches.Count + 1)
		ReDim PageStr(Matches.Count + 1)

		i = 1
		
		PageIndex(0) = 0

		For Each Match in Matches			' 遍历匹配集合。 
			PageIndex(i) = Match.FirstIndex
			PageStr(i)	 = Match.Value

			i = i + 1
		Next

		PageIndex(i) = Len(strng)
	End Function

	Function PageNav (iCount, iCurrentPage)
		Dim i
		Dim OutStr

		For i = 1 To iCount
			If i = iCurrentPage Then 
				OutStr = OutStr & "<span style='color: red;'>[" & i & "]</span>&nbsp;"
			ElseIf i = 1 Then
				OutStr = OutStr & "<a href='?articleid=" & ArticleId & "'>[" & i & "]</a>&nbsp;"
			Else
				OutStr = OutStr & "<a href='?articleid=" & ArticleId & "&amp;page=" & i & "'>[" & i & "]</a>&nbsp;"
			End If
		Next

		PageNav = OutStr
	End Function
End Class
%>