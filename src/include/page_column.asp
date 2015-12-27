<%
'====================================================================
'= Team Elite - EliteCMS
'= Copyright (c) 2004 - 2008 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：page_column.asp
'= 摘    要：列表类文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-03-14
'====================================================================

Class page_Column
	Public PageContent
	Private Info, ID

	Public Sub Make (iID, aInfo, iPageNum)
		Info = aInfo
		ID = iID

		PageContent = EA_Temp.Load_Template(Info(9, 0), 4)

		EA_Pub.SysInfo(16) = Info(0, 0) & "," & EA_Pub.SysInfo(16)
		If Len(Info(2,0)) Then EA_Pub.SysInfo(17) = Info(2, 0)

		EA_Temp.Title= Info(0, 0) & " - " & EA_Pub.SysInfo(0)
		EA_Temp.Nav	 = "<a href=""" & SystemFolder & """>" & EA_Pub.SysInfo(0) & "</a>" & EA_Pub.Get_NavByColumnCode(Info(1, 0), 1)

		EA_Temp.SetVariable "List.ID", ID, PageContent
		EA_Temp.SetVariable "List.Name", Info(0, 0), PageContent
		EA_Temp.SetVariable "List.Description", Info(2, 0), PageContent
		EA_Temp.SetVariable "List.TopicTotal", Info(3, 0), PageContent

		If EA_Temp.ChkTag("List.Topic", PageContent) Then ListTopic iPageNum

		EA_Temp.Replace_PublicTag PageContent
	End Sub

	Private Sub ListTopic (PageNum)
		Dim Url
		Dim ArticleList
		Dim PageCount, PageSize
		Dim Temp, ListBlock
		Dim ForTotal
		Dim ArticleUrlType
		Dim i

		PageSize	= Info(17, 0)
		PageCount	= EA_Pub.Stat_Page_Total(PageSize, Info(3, 0))
		If CLng(PageNum) > PageCount And PageCount > 0 Then PageNum = PageCount

		'load article list
		If Info(3, 0) > 0 Then ArticleList = EA_DBO.Get_Article_ByColumnId(ID, PageNum, PageSize)

		If Info(12, 0) > 0 Or Info(13, 0) = 1 Then 
			ArticleUrlType = 1
		Else
			ArticleUrlType = EA_Pub.SysInfo(18)
		End If

		ListBlock = EA_Temp.GetBlock("List.Topic", PageContent)

		If IsArray(ArticleList) Then
			ForTotal = UBound(ArticleList, 2)

			For i = 0 To ForTotal
				Temp = ListBlock
		  
				EA_Temp.SetVariable "Url", EA_Pub.Cov_ArticlePath(ArticleList(0, i), ArticleList(3, i), ArticleUrlType), Temp
				EA_Temp.SetVariable "Title", EA_Pub.Add_ArticleColor(ArticleList(1, i), EA_Pub.Base_HTMLFilter(ArticleList(2, i))), Temp
				EA_Temp.SetVariable "SubUrl", ArticleList(15, i), Temp
				EA_Temp.SetVariable "SubTitle", ArticleList(14, i), Temp
				EA_Temp.SetVariable "Date", FormatDateTime(ArticleList(3, i), 2), Temp
				EA_Temp.SetVariable "Time", FormatDateTime(ArticleList(3, i), 4), Temp
				EA_Temp.SetVariable "CommentNum", ArticleList(4, i), Temp
				EA_Temp.SetVariable "Summary", ArticleList(5, i), Temp
				EA_Temp.SetVariable "LastComment", ArticleList(6, i), Temp
				EA_Temp.SetVariable "ViewNum", ArticleList(7, i), Temp
				EA_Temp.SetVariable "Icon", EA_Pub.Chk_ArticleType(ArticleList(8, i), ArticleList(10, i)), Temp
				EA_Temp.SetVariable "Img", ArticleList(9, i), Temp
				EA_Temp.SetVariable "Author", ArticleList(11, i), Temp
				EA_Temp.SetVariable "Tag", TagList(ArticleList(13, i)), Temp

				EA_Temp.SetBlock "List.Topic", Temp, PageContent
			Next
		End If

		EA_Temp.CloseBlock "List.Topic", PageContent

		If EA_Temp.ChkTag("List.PageNav", PageContent) Then 
			If ArticleUrlType = 1 Then
				Url = "list.asp?classid=" & ID & "&page=$page"
			Else
				Url = Replace(EA_Pub.Cov_ColumnPath(ID, EA_Pub.SysInfo(18)), "_1", "_$page")
			End If

			EA_Temp.SetVariable "List.PageNav", EA_Temp.PageList(PageCount, PageNum, Url), PageContent
		End If
	End Sub

	Private Function TagList (Keyword)
		Dim OutStr

		If Len(Keyword) > 0 Then
			Dim TempArray, i
			Dim ForTotal

			TempArray= Split(Keyword, ",")

			ForTotal = UBound(TempArray)

			For i = 0 To ForTotal
				If Len(TempArray(i)) > 0 Then OutStr = OutStr & "<a href='" & SystemFolder & "search.asp?action=query&amp;field=1&amp;keyword=" & server.urlencode(Trim(TempArray(i))) & "'>" & Trim(TempArray(i)) & "</a>&nbsp;"
			Next
		End If

		TagList = OutStr
	End Function
End Class
%>