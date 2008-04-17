<!--#Include File="include/inc.asp"-->
<%
'====================================================================
'= Team Elite - EliteCMS
'= Copyright (c) 2004 - 2008 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：Print.asp
'= 摘    要：文章打印版本文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-04-17
'====================================================================

Dim ArticleId, ArticleInfo
Dim FirstArticle,NextArticle
Dim i,TempStr,TempArray
Dim IsView
Dim PageContent

ArticleId=EA_Pub.SafeRequest(3,"articleid",0,0,0)

'load article info
ArticleInfo = EA_DBO.Get_Article_Info(ArticleId, 1)
If Not IsArray(ArticleInfo) Then Call EA_Pub.ShowErrMsg(2, 0)
If Not ArticleInfo(20, 0) Or ArticleInfo(21, 0) Then Call EA_Pub.ShowErrMsg(2, 0)

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

If Not IsView Then 
	ArticleInfo(5,0)="<br><br><b>您当前的权限不允许查看该文章，请先 [<a href='member/login.asp' target='_blank'>登陆</a>] 或 [<a href='member/register.asp' target='_blank'>注册</a>]。</b>"
End If

PageContent	= EA_Temp.Load_Template(0, 7)

EA_Temp.Title= ArticleInfo(3, 0) & " - " & ArticleInfo(2, 0) & " - " & EA_Pub.SysInfo(0)
EA_Temp.Nav	 = "<a href=""" & SystemFolder & """>" & EA_Pub.SysInfo(0) & "</a>" & EA_Pub.Get_NavByColumnCode(ArticleInfo(1, 0), 0) & " - <a href=""" & EA_Pub.Cov_ArticlePath(ArticleId, ArticleInfo(13, 0), EA_Pub.SysInfo(18)) & """><strong>" & ArticleInfo(3, 0) & "</strong></a>"

EA_Temp.SetVariable "Article.ColumnID", ArticleInfo(0, 0), PageContent
EA_Temp.SetVariable "Article.ID", ArticleId, PageContent
EA_Temp.SetVariable "Article.Url", EA_Pub.Cov_ArticlePath(ArticleId, ArticleInfo(13, 0), EA_Pub.SysInfo(18)), PageContent
EA_Temp.SetVariable "Article.Title", EA_Pub.Add_ArticleColor(ArticleInfo(17, 0),ArticleInfo(3, 0)), PageContent
EA_Temp.SetVariable "Article.Date", FormatDateTime(ArticleInfo(13, 0), 2), PageContent
EA_Temp.SetVariable "Article.Time", FormatDateTime(ArticleInfo(13, 0), 4), PageContent
EA_Temp.SetVariable "Article.Author", ArticleInfo(8, 0), PageContent
EA_Temp.SetVariable "Article.Source", ArticleInfo(15,0), PageContent
EA_Temp.SetVariable "Article.Summary", ArticleInfo(4, 0), PageContent
EA_Temp.SetVariable "Article.Content", ArticleInfo(5,0), PageContent

PageContent	= EA_Temp.Replace_PublicTag(PageContent)

Response.Write PageContent

Call EA_Pub.Close_Obj
Set EA_Pub = Nothing
%>
