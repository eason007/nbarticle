<!--#Include File="include/inc.asp"-->
<%
'====================================================================
'= Team Elite - EliteCMS
'= Copyright (c) 2004 - 2008 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：Review.asp
'= 摘    要：评论文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-03-19
'====================================================================

Dim ArticleId
ArticleId=EA_Pub.SafeRequest(3,"articleid",0,0,3)

EA_Temp.Title=EA_Pub.SysInfo(0)&" - 评论系统"
EA_Temp.Nav="<a href=""./""><b>"&EA_Pub.SysInfo(0)&"</b></a> - <a href=""article.asp?articleid="&ArticleId&""" target=""_blank"">文章正文</a> - 评论列表"
%>