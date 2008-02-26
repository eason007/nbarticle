<!--#include file="_cls_teamplate.asp"-->
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

Class page_Column
	Private PageContent
	Private Template
	Private Info, ID

	Public Function Make (iID, ByVal aInfo)
		Info = aInfo
		ID = iID

		Set Template = New cls_NEW_TEMPLATE

		PageContent  = EA_Temp.Load_Template(Info(9, 0), 4)

		If Template.ChkBlock("list", PageContent) Then MakeArticleList()
		If Template.ChkBlock("placard", PageContent) Then MakePlacardList()

		EA_Pub.SysInfo(16) = Info(0, 0) & "," & EA_Pub.SysInfo(16)
		If Len(Info(2,0)) Then EA_Pub.SysInfo(17) = Info(2, 0)

		EA_Temp.Title	= Info(0, 0) & " - " & EA_Pub.SysInfo(0)
		EA_Temp.Nav		= "<a href=""./""><b>" & EA_Pub.SysInfo(0) & "</b></a>" & EA_Pub.Get_NavByColumnCode(Info(1, 0))

		PageContent		= Replace(PageContent, "{$ColumnId$}", ID)
		PageContent		= Replace(PageContent, "{$ColumnName$}", Info(0, 0))
		PageContent		= Replace(PageContent, "{$Info$}", Info(2, 0))
		PageContent		= Replace(PageContent, "{$ColumnTopicTotal$}", Info(3, 0))
		PageContent		= Replace(PageContent, "{$ColumnMangerTotal$}", Info(4, 0))

		EA_Temp.Find_TemplateTagByInput "ChildColumnNav", ChildColumnNav(Info(1, 0)), PageContent

		PageContent = EA_Temp.Replace_PublicTag(PageContent)

		Make = PageContent
	End Function

	Private Function MakeArticleList ()
		Dim FieldName(0),FieldValue(0)
		Dim ArticleList
		Dim PageNum,PageCount,PageSize
		Dim Temp,ListBlock
		Dim ForTotal
		Dim ArticleUrlType
		Dim i

		FieldName(0)	= "classid"
		FieldValue(0)	= ID

		PageNum		= EA_Pub.SafeRequest(3, "page", 0, 1, 0)
		PageSize	= Info(17, 0)
		PageCount	= EA_Pub.Stat_Page_Total(PageSize, Info(3, 0))
		If CLng(PageNum) > PageCount And PageCount > 0 Then PageNum = PageCount

		'load article list
		If Info(3, 0) > 0 Then ArticleList = EA_DBO.Get_Article_ByColumnId(ID, PageNum, PageSize)

		If Info(12, 0) > 0 Or Info(13, 0) = 1 Then 
			ArticleUrlType = 0
		Else
			ArticleUrlType = 1
		End If

		ListBlock	= Template.GetBlock("list", PageContent)

		If IsArray(ArticleList) Then
			ForTotal = UBound(ArticleList, 2)

			For i  =0 To ForTotal
				Temp = ListBlock
		  
				Template.SetVariable "Url", EA_Pub.Cov_ArticlePath(ArticleList(0, i), ArticleList(3, i), ArticleUrlType), Temp
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

		PageContent		= Replace(PageContent, "{$ColumnPageNumNav$}", EA_Temp.PageList(PageCount, PageNum, FieldName, FieldValue))
	End Function

	Private Function TagList (Keyword)
		Dim TempArray,i
		Dim ForTotal
		Dim OutStr

		If Len(Keyword) > 0 Then
			TempArray= Split(Keyword,",")

			ForTotal = UBound(TempArray)

			For i=0 To ForTotal
				If Len(TempArray(i)) > 0 Then OutStr = OutStr & "<a href='" & SystemFolder & "search.asp?action=query&field=1&keyword=" & server.urlencode(Trim(TempArray(i))) & "' rel='external'>" & Trim(TempArray(i)) & "</a>&nbsp;"
			Next
		End If

		TagList = OutStr
	End Function

	Private Function ChildColumnNav(ColumnCode)
		Dim ChilColumnConfig
		Dim Temp,OutStr,Column,j
		Dim TempArray, ForTotal, i, StepLen
		Dim ChildColumnList

		TempArray=EA_DBO.Get_Column_Nav(ColumnCode)
		If IsArray(TempArray) Then 
			ForTotal = UBound(TempArray,2)
			For i=0 To ForTotal
				StepLen=(Len(TempArray(1,i))/4)*2-2
				If Len(TempArray(1,i)) = Len(ColumnCode)+4 And ColumnCode = Left(TempArray(1,i),Len(ColumnCode)) Then ChildColumnList = ChildColumnList & TempArray(0,i) & "," & TempArray(2,i) & "|"
			Next
		End If

		Temp = Split(ChildColumnList,"|")

		ChilColumnConfig = EA_Temp.Find_TemplateTagValues("ChildColumnNav",PageContent)
		If Not IsArray(ChilColumnConfig) Then Exit Function

		j = 1
		ForTotal = UBound(Temp)-1

		For i=0 To ForTotal
			Column = Split(Temp(i),",")

			OutStr = OutStr & "<a href="""&EA_Pub.Cov_ColumnPath(Column(0),EA_Pub.SysInfo(18))&""">"&Column(1)
			OutStr = OutStr & "</a>&nbsp;"

			If j = CLng(ChilColumnConfig(1)) Then Exit For
			j = j + 1
			If (i+1) Mod ChilColumnConfig(0) = 0 And (i+1) <= (UBound(Temp)-1) Then OutStr = OutStr & "<br>"
		Next

		ChildColumnNav = OutStr
	End Function
End Class
%>