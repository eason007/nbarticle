<!--#Include File="conn.asp" -->
<!--#Include File="include/inc.asp"-->
<!--#Include File="include/page_article.asp"-->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：Article.asp
'= 摘    要：文章显示文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-03-05
'====================================================================

Dim ArticleId, ArticleInfo
Dim Page

ArticleId	= EA_Pub.SafeRequest(3, "articleid", 0, 0, 0)
Page		= EA_Pub.SafeRequest(3, "page", 0, 1, 0)

'get article
ArticleInfo = EA_DBO.Get_Article_Info(ArticleId, 1)
If Not IsArray(ArticleInfo) Then Call EA_Pub.ShowErrMsg(9, 1)
If Not ArticleInfo(20, 0) Or ArticleInfo(21, 0) Then Call EA_Pub.ShowErrMsg(9, 1)

Dim PageContent
Dim MakeHtml

MakeHtml = False

If EA_Pub.SysInfo(18) = "0" Then
	Dim sHTMLFilePath

	If ArticleInfo(22, 0) <= 0 And ArticleInfo(23, 0) = 0 Then
		sHTMLFilePath = EA_Pub.Cov_ArticlePath(ArticleId, ArticleInfo(13, 0), "0")
		
		If ArticleInfo(10, 0) Then
			PageContent = "<meta http-equiv=""refresh"" content=""0;URL=" & ArticleInfo(11, 0) & """>"
			
			Call EA_Pub.Save_HtmlFile(sHTMLFilePath, PageContent)
		End If
		
		If Not EA_Pub.Chk_IsExistsHtmlFile(sHTMLFilePath) Then 
			MakeHtml = True
		Else
			Call EA_DBO.Set_Article_ViewNum_UpDate(ArticleId)

			Call EA_Pub.Close_Obj
			Set EA_Pub=Nothing
			
			Response.Redirect sHTMLFilePath
			Response.End 
		End If
	End If
End If

Dim IsView

If ArticleInfo(22, 0) > 0 Or ArticleInfo(23, 0) <> 0 Then 
	If Not EA_Pub.IsMember Then 
		IsView = False
	Else
		If CDbl(EA_Pub.Mem_GroupSetting(2)) >= CDbl(ArticleInfo(22, 0)) Then 
			If ArticleInfo(23, 0) Then 
				If EA_Pub.Mem_GroupSetting(3) = "1" Then 
					IsView = True
				Else
					IsView = False
				End If
			Else
				IsView = True
			End If
		Else
			IsView = False
		End If
	End If
Else
	IsView = True
End If

If ArticleInfo(10, 0) And IsView Then
	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing

	Response.Redirect ArticleInfo(11, 0)
	Response.End
End If

Dim clsArticle

Set clsArticle = New page_Article

PageContent = clsArticle.Make(ArticleId, ArticleInfo, Page, IsView)

If MakeHtml Then
	Call EA_Pub.Save_HtmlFile(sHTMLFilePath,PageContent)
	
	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing

	Response.Redirect sHTMLFilePath
	Response.End 
Else
	Response.Write PageContent

	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing
End If
%>
