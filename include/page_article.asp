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

Class page_Article
	Public Function Make (ID, Info)
		Dim ArticleList
		Dim FieldName(0),FieldValue(0)
		Dim i
		Dim PageNum,PageCount,PageSize
		Dim Template
		Dim Temp,ListBlock
		Dim ForTotal
		Dim PageContent

		Set Template = New cls_NEW_TEMPLATE

		FieldName(0)	= "classid"
		FieldValue(0)	= ID

		PageNum		= EA_Pub.SafeRequest(3, "page", 0, 1, 0)
		PageSize	= Info(17, 0)
		PageCount	= EA_Pub.Stat_Page_Total(PageSize, Info(3, 0))
		If CLng(PageNum) > PageCount And PageCount > 0 Then PageNum = PageCount

		'load article list
		If Info(3, 0) > 0 Then ArticleList = EA_DBO.Get_Article_ByColumnId(ID, PageNum, PageSize)

		EA_Pub.SysInfo(16) = Info(0, 0) & "," & EA_Pub.SysInfo(16)
		If Len(Info(2,0)) Then EA_Pub.SysInfo(17) = Info(2, 0)

		PageContent = EA_Temp.Load_Template(Info(9, 0), 4)
		ListBlock	= Template.GetBlock("list", PageContent)

		If IsArray(ArticleList) Then
			ForTotal = UBound(ArticleList, 2)

			For i  =0 To ForTotal
				Temp = ListBlock
		  
				Template.SetVariable "Url", EA_Pub.Cov_ArticlePath(ArticleList(0, i), ArticleList(3, i), EA_Pub.SysInfo(18)), Temp
				Template.SetVariable "Title", EA_Pub.Add_ArticleColor(ArticleList(1, i), EA_Pub.Base_HTMLFilter(ArticleList(2, i))), Temp
				Template.SetVariable "Date", ArticleList(3, i), Temp
				Template.SetVariable "CommentNum", ArticleList(4, i), Temp
				Template.SetVariable "Summary", ArticleList(5, i), Temp
				Template.SetVariable "LastComment", ArticleList(6, i), Temp
				Template.SetVariable "ViewNum", ArticleList(7, i), Temp
				Template.SetVariable "Icon", EA_Pub.Chk_ArticleType(ArticleList(8, i),ArticleList(10, i)), Temp
				Template.SetVariable "Img", ArticleList(9, i), Temp
				Template.SetVariable "Author", "<a href='" & SystemFolder & "florilegium.asp?a_name=" & ArticleList(11, i) & "&a_id=" & ArticleList(12, i) & "' rel=""external"">" & ArticleList(11, i) & "</a>", Temp
				Template.SetVariable "Tag", TagList(ArticleList(13, i)), Temp

				Template.SetBlock "list", Temp, PageContent
			Next

			Template.CloseBlock "list", PageContent
		End If

		EA_Temp.Title	= Info(0, 0) & " - " & EA_Pub.SysInfo(0)
		EA_Temp.Nav		= "<a href=""./""><b>" & EA_Pub.SysInfo(0) & "</b></a>" & EA_Pub.Get_NavByColumnCode(Info(1, 0))

		PageContent		= Replace(PageContent, "{$ColumnId$}", ID)
		PageContent		= Replace(PageContent, "{$ColumnName$}", Info(0, 0))
		PageContent		= Replace(PageContent, "{$Info$}", Info(2, 0))
		PageContent		= Replace(PageContent, "{$ColumnTopicTotal$}", Info(3, 0))
		PageContent		= Replace(PageContent, "{$ColumnMangerTotal$}", Info(4, 0))
		PageContent		= Replace(PageContent, "{$ColumnPageNumNav$}", EA_Temp.PageList(PageCount, PageNum, FieldName, FieldValue))

		EA_Temp.Find_TemplateTagByInput "ChildColumnNav", ChildColumnNav(Info(1, 0)), PageContent

		PageContent = EA_Temp.Replace_PublicTag(PageContent)

		Make = PageContent
	End Function


	
End Class
%>