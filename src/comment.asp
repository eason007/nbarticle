<!--#Include File="include/inc.asp"-->
<!--#Include File="include/page_comment.asp"-->
<%
'====================================================================
'= Team Elite - EliteCMS
'= Copyright (c) 2004 - 2008 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：Comment.asp
'= 摘    要：评论文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-04-17
'====================================================================

Dim ArticleId, ArticleInfo
Dim Page
Dim PageContent

ArticleId	= EA_Pub.SafeRequest(3, "articleid", 0, 0, 3)
ArticleInfo	= EA_DBO.Get_Article_Info_Single(ArticleId)
Page		= EA_Pub.SafeRequest(3, "page", 0, 1, 0)
If Not IsArray(ArticleInfo) Then ErrMsg = SysMsg(34):Call EA_Pub.ShowErrMsg(0, 0)
If Not ArticleInfo(20, 0) Or ArticleInfo(21, 0) Then ErrMsg = SysMsg(34):Call EA_Pub.ShowErrMsg(0, 0)


Dim clsComment

Set clsComment = New page_Comment

PageContent = clsComment.Make(ArticleId, ArticleInfo, Page)

Response.Write PageContent

Call EA_Pub.Close_Obj
Set EA_Pub=Nothing
%>