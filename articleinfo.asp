<!--#Include File="conn.asp" -->
<!--#Include File="include/inc.asp"-->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：ArticleInfo.asp
'= 摘    要：文章信息文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2006-03-22
'====================================================================

Response.Buffer = True 
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache"

Call EA_Pub.Chk_Post

Dim ArticleId
Dim Action
Dim ArticleInfo
ArticleId=EA_Pub.SafeRequest(3,"articleid",0,0,3)
Action=Request.QueryString ("action")

ArticleInfo=EA_DBO.Get_Article_Info(ArticleId,0)
If IsArray(ArticleInfo) Then
	Select Case LCase(Action)
	Case "viewtotal"
		Response.Write "document.write ("""&ArticleInfo(6,0)&""");"
		Call EA_DBO.Set_Article_ViewNum_UpDate(ArticleId)
	Case "commenttotal"
		Response.Write "document.write ("""&ArticleInfo(9,0)&""");"

		Call EA_DBO.Set_Article_CommentNum_UpDate(ArticleId)
	End Select
End If

EA_Pub.Close_Obj
Set EA_Pub=Nothing
%>