<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2004 - 2008 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：cls_article.asp
'= 摘    要：内容页类文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-03-17
'====================================================================

Class page_Article
	Public PageIndex(), PageStr()
	Public ID, Info
	

	Public Function Make (iID, ByRef aInfo, Page, IsView)
		ID	 = iID
		Info = aInfo

		Dim FirstArticle, NextArticle
		Dim i, TempStr
		Dim PageContent

		PageContent  = EA_Temp.Load_Template(Info(24, 0), 5)

		EA_Temp.Title= Info(3, 0) & " - " & Info(2, 0) & " - " & EA_Pub.SysInfo(0)
		EA_Temp.Nav	 = "<a href=""" & SystemFolder & """>" & EA_Pub.SysInfo(0) & "</a>" & EA_Pub.Get_NavByColumnCode(Info(1, 0), 0) & " - <a href=""" & EA_Pub.Cov_ArticlePath(ID, Info(13, 0), EA_Pub.SysInfo(18)) & """><strong>" & Info(3, 0) & "</strong></a>"

		EA_Pub.SysInfo(16) = Info(12, 0) & "," & EA_Pub.SysInfo(16)
		EA_Pub.SysInfo(17) = Info(4, 0)

		If Not IsView Then 
			TempStr = "<strong>" & SysMsg(11) & "</strong>"
		Else
			Call CutContent("\[NextPage([^\]])*\]", Info(5, 0))

			If UBound(PageIndex) = 1 Then
				Call Cov_InsideLink(Info(5, 0), Info(0, 0))

				TempStr = "<div id=""article"">" & Info(5, 0) & "</div>"
			Else
				TempStr = Mid(Info(5, 0), PageIndex(Page - 1) + Len(PageStr(Page - 1)) + 1, PageIndex(Page) - PageIndex(Page - 1) - Len(PageStr(Page - 1)))

				Call Cov_InsideLink(TempStr, Info(0, 0))

				TempStr = "<div id=""article"">" & TempStr & "</div>"
				TempStr = TempStr & "<div style='TEXT-ALIGN: center;margin-bottom: 5px;'>" & PageNav(UBound(PageIndex), Page) & "</div>"
			End If
		End If

		EA_Temp.SetVariable "Article.ColumnID", Info(0, 0), PageContent
		EA_Temp.SetVariable "Article.ID", ID, PageContent
		EA_Temp.SetVariable "Article.Url", EA_Pub.Cov_ArticlePath(ID, Info(13, 0), EA_Pub.SysInfo(18)), PageContent
		EA_Temp.SetVariable "Article.Title", EA_Pub.Add_ArticleColor(Info(17, 0),Info(3, 0)), PageContent
		EA_Temp.SetVariable "Article.Date", FormatDateTime(Info(13, 0), 2), PageContent
		EA_Temp.SetVariable "Article.Time", FormatDateTime(Info(13, 0), 4), PageContent
		EA_Temp.SetVariable "Article.Author", Info(8, 0), PageContent
		EA_Temp.SetVariable "Article.Source", "<a href='" & Info(16,0) & "'>" & Info(15,0) & "</a>", PageContent
		EA_Temp.SetVariable "Article.Summary", Info(4, 0), PageContent
		EA_Temp.SetVariable "Article.Content", TempStr, PageContent
		EA_Temp.SetVariable "Article.Tag", TagList(Info(12, 0)), PageContent

		EA_Temp.SetVariable "Article.ViewTotal", "<script type=""text/javascript"" src=""" & SystemFolder & "action.asp?action=viewtotal&amp;articleid=" & ID & """></script>", PageContent
		EA_Temp.SetVariable "Article.CommentTotal", "<script type=""text/javascript"" src=""" & SystemFolder & "action.asp?action=commenttotal&amp;articleid=" & ID & """></script>", PageContent

		If EA_Temp.ChkTag("Article.FirstTopic", PageContent) Then
			FirstArticle = EA_DBO.Get_Article_FirstArticle(Info(0, 0), Info(25, 0), ID)

			If IsArray(FirstArticle) Then
				EA_Temp.SetVariable "Article.FirstTopic", "<a href='" & EA_Pub.Cov_ArticlePath(FirstArticle(0, 0), FirstArticle(3, 0), EA_Pub.SysInfo(18)) & "'>" & EA_Pub.Add_ArticleColor(FirstArticle(2, 0),FirstArticle(1, 0)) & "</a>", PageContent
			End If
		End If

		If EA_Temp.ChkTag("Article.NextTopic", PageContent) Then
			NextArticle  =EA_DBO.Get_Article_NextArticle(Info(0, 0), Info(25, 0), ID)

			If IsArray(NextArticle) Then
				EA_Temp.SetVariable "Article.NextTopic", "<a href='" & EA_Pub.Cov_ArticlePath(NextArticle(0, 0), NextArticle(3, 0), EA_Pub.SysInfo(18)) & "'>" & EA_Pub.Add_ArticleColor(NextArticle(2, 0), NextArticle(1, 0)) & "</a>", PageContent
			End If
		End If

		If EA_Temp.ChkTag("Article.RelatedList", PageContent) Then Call CorrList(Info(12, 0), Info(0, 0), PageContent)

		PageContent = EA_Temp.Replace_PublicTag(PageContent)

		Make = PageContent
	End Function

	Private Function PageNav (iCount, iCurrentPage)
		Dim i
		Dim OutStr
		Dim Url
		Dim re

		If EA_Pub.SysInfo(18) <> "0" Then
			Url = EA_Pub.Cov_ArticlePath(ID, Info(13, 0), EA_Pub.SysInfo(18)) & "&page=#1"
		Else
			Set re = New RegExp

			re.IgnoreCase	= True
			re.Global		= True
			re.Pattern		= "\/(\w+)_(\d+).(\w+)"

			Url = EA_Pub.Cov_ArticlePath(ID, Info(13, 0), EA_Pub.SysInfo(18))
			Url = re.Replace(Url, "/$1_$2_#1.$3")
		End If

		For i = 1 To iCount
			If i = iCurrentPage Then 
				OutStr = OutStr & "<span style='color: red;'>[" & i & "]</span>&nbsp;"
			ElseIf i = 1 Then
				OutStr = OutStr & "<a href='" & Replace(Url, "_#1", "") & "'>[" & i & "]</a>&nbsp;"
			Else
				OutStr = OutStr & "<a href='" & Replace(Url, "#1", i) & "'>[" & i & "]</a>&nbsp;"
			End If
		Next

		PageNav = OutStr
	End Function

	Private Function TagList (ByRef Keyword)
		Dim TempArray, i
		Dim ForTotal
		Dim OutStr

		If Len(Trim(Keyword)) > 0 And Not IsNull(Keyword) Then
			TempArray= Split(Keyword, ",")

			ForTotal = UBound(TempArray)

			For i = 0 To ForTotal
				If Len(Trim(TempArray(i))) > 0 And Not IsNull(TempArray(i)) Then OutStr = OutStr & "<a href='" & SystemFolder & "search.asp?action=query&amp;field=1&amp;keyword=" & Trim(TempArray(i)) & "'>" & Trim(TempArray(i)) & "</a>&nbsp;"
			Next
		End If

		TagList = OutStr
	End Function

	Private Sub CorrList(Keyword, ColumnId, ByRef PageContent)
		If Len(Keyword) = 0 Or IsNull(Keyword) Then Exit Sub

		Dim Block, Parameter
		Dim List
		Dim Temp, ForTotal, i, TempArray
		Dim SearchKeyWord

		Block = EA_Temp.GetBlock("Article.RelatedList", PageContent)
		If Block = "" Then Exit Sub

		Parameter = EA_Temp.GetParameter("Parameter", Block)
		If Not IsArray(Parameter) Then EA_Temp.CloseBlock "Article.RelatedList", PageContent: Exit Sub

		If Keyword = "" Then EA_Temp.CloseBlock "Article.RelatedList", PageContent: Exit Sub

		TempArray= Split(Keyword, ",")
		ForTotal = UBound(TempArray)

		For i = 0 To ForTotal
			Select Case iDataBaseType
			Case 0
				SearchKeyWord = SearchKeyWord & "InStr(','+keyword+',','," & TempArray(i) & ",')>0 OR "
			Case 1
				SearchKeyWord = SearchKeyWord & " CharIndex('," & TempArray(i) & ",',','+keyword+',')>0 OR "
			End Select
		Next

		List = EA_DBO.Get_Article_CorrList(SearchKeyWord, ID, ColumnId, CInt(Parameter(0)), CInt(Parameter(2)))
		If Not IsArray(List) Then EA_Temp.CloseBlock "Article.RelatedList", PageContent: Exit Sub
		
		ForTotal = UBound(List, 2)

		For i = 0 To ForTotal
			Temp = Block

			List(3, i) = EA_Pub.Base_HTMLFilter(List(3, i))
			List(3, i) = EA_Pub.Cut_Title(List(3, i), Parameter(1))
			
			EA_Temp.SetVariable "Url", EA_Pub.Cov_ArticlePath(List(0, i), List(5, i), EA_Pub.SysInfo(18)), Temp
			EA_Temp.SetVariable "Title", EA_Pub.Add_ArticleColor(List(4, i), List(3, i)), Temp
			EA_Temp.SetVariable "Date", FormatDateTime(List(5, i), 2), Temp
			EA_Temp.SetVariable "Time", FormatDateTime(List(5, i), 4), Temp
			EA_Temp.SetVariable "Icon", EA_Pub.Chk_ArticleType(List(6, i), List(7, i)), Temp
			EA_Temp.SetVariable "Summary", List(10, i), Temp
			EA_Temp.SetVariable "Author", List(9, i), Temp

			EA_Temp.SetBlock "Article.RelatedList", Temp, PageContent
		Next

		EA_Temp.CloseBlock "Article.RelatedList", PageContent
	End Sub

	Public Sub CutContent(patrn, strng) 
		Dim regEx, Match, Matches			' 建立变量。 
		Dim i

		Set regEx = New RegExp				' 建立正则表达式。 

		regEx.Pattern	 = patrn			' 设置模式。 
		regEx.IgnoreCase = True				' 设置是否区分字符大小写。 
		regEx.Global	 = True				' 设置全局可用性。 

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

		Set regEx	= Nothing
		Set Matches = Nothing
	End Sub

	'**************************************************
	'替换文章正文中的内部连接函数
	'输入参数：
	'	1、文章内容
	'	2、文章地址[栏目id]
	'**************************************************
	Private Sub Cov_InsideLink(ByRef StrContent, ColumnId)
		Dim i
		Dim TempArray
		Dim WordIndex
		Dim ForTotal
		
		TempArray = EA_DBO.Get_InsideLink_ByColumn(ColumnId)
		If IsArray(TempArray) Then 
			ForTotal = UBound(TempArray, 2)

			For i = 0 To ForTotal
				StrContent = Replace(StrContent, TempArray(0, i),"<a href=""" & TempArray(1, i)&""">" & TempArray(0, i) & "</a>")
			Next
		End If
	End Sub
End Class
%>