<!--#Include File="include/inc.asp"-->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2004 - 2008 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：App_Link.asp
'= 摘    要：申请友情连接文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-03-12
'====================================================================

Dim Action
Action = Request.QueryString ("action")

Select Case LCase(Action)
Case "save_link"
	Call SetLink
Case "viewtotal", "commenttotal"
	Call GetArticleInfo()
End Select

Sub SetLink()
	Call EA_Pub.Chk_Post

	If EA_Pub.Chk_PostTime(30, "s", Session("lastpost")) Then
		ErrMsg = SysMsg(0)
		Call EA_Pub.ShowErrMsg(0, 2)
	End If
		
	Dim LinkName, LinkImg, LinkUrl, LinkInfo, ColumnId, Style

	LinkName= EA_Pub.SafeRequest(2, "name", 1, "", 1)
	LinkImg = EA_Pub.SafeRequest(2, "logo", 1, "", 1)
	LinkUrl = EA_Pub.SafeRequest(2, "url", 1, "", 1)
	LinkInfo= EA_Pub.SafeRequest(2, "info", 1, "", 1)
	ColumnId= EA_Pub.SafeRequest(2, "column", 0, 0, 0)
	Style	= EA_Pub.SafeRequest(2, "style", 0, 0, 0)
	FoundErr= False
	
	If LinkName = "" Or Len(LinkName) > 50 Then FoundErr = True
	If Len(LinkImg) > 150 Then FoundErr = True
	If LinkUrl = "" Or Len(LinkUrl) > 150 Then FoundErr = True
	
	If Not FoundErr Then
		Call EA_DBO.Set_FriendList_Insert(LinkName, LinkImg, LinkUrl, LinkInfo, ColumnId, Style, 0, 0)
		
		ErrMsg = SysMsg(1)
		Session("lastpost") = Now()
	Else
		ErrMsg = SysMsg(2)
	End If

	Call EA_Pub.ShowErrMsg(0, 2)
End Sub

Sub GetArticleInfo()
	Response.Buffer			 = True 
	Response.ExpiresAbsolute = Now() - 1 
	Response.Expires		 = 0 
	Response.CacheControl	 = "no-cache"

	Call EA_Pub.Chk_Post

	Dim ArticleId
	Dim ArticleInfo

	ArticleId	= EA_Pub.SafeRequest(3, "articleid", 0, 0, 3)
	ArticleInfo	= EA_DBO.Get_Article_Info(ArticleId,0)

	If IsArray(ArticleInfo) Then
		Select Case LCase(Action)
		Case "viewtotal"
			Response.Write "document.write (""" & ArticleInfo(6, 0) & """);"
			Call EA_DBO.Set_Article_ViewNum_UpDate(ArticleId)
		Case "commenttotal"
			Response.Write "document.write (""" & ArticleInfo(9, 0) & """);"
			Call EA_DBO.Set_Article_CommentNum_UpDate(ArticleId)
		End Select
	End If

	Call EA_Pub.Close_Obj
	Set EA_Pub = Nothing
End Sub


Sub Save_Review
	Call EA_Pub.Chk_Post
	
	Dim RUserId, RUserName, RContent, RState, R_IsGhost
	Dim IP
	Dim ArticleId
	
	ArticleId	= EA_Pub.SafeRequest(3, "articleid", 0, 0, 3)
	RContent	= EA_Pub.BadWords_Filter(EA_Pub.SafeRequest(2, "review", 1, "", 2))
	IP			= EA_Pub.Get_UserIp
	
	
	If EA_Pub.IsMember Then 
		If EA_Pub.Mem_GroupSetting(8) = "0" Then 
			ErrMsg	= SysMsg(3)
			FoundErr= True
		End If
		If EA_Pub.Mem_GroupSetting(7) > "0" Then 
			If EA_Pub.Chk_PostTime(CLng(EA_Pub.Mem_GroupSetting(7)), "n", EA_Pub.Mem_Info(2)) Then 
				ErrMsg	= Replace(SysMsg(4), "$1", EA_Pub.Mem_GroupSetting(7))
				FoundErr=True
			End If
		End If
		If EA_Pub.Mem_GroupSetting(9) = "0" Then 
			RState = 1
		Else
			RState = 0
		End If

		RUserId	 = EA_Pub.Mem_Info(0)
		RUserName= EA_Pub.Mem_Info(1)
	Else
		If EA_Pub.SysInfo(19) = "0" Then 
			ErrMsg	= SysMsg(5)
			FoundErr= True
		Else
			If EA_Pub.SysInfo(20) = "0" Then 
				RState = 1
			Else
				RState = 0
			End If
		End If
		RUserId	 = 0
		RUserName= EA_Pub.SafeRequest(2, "name", 1, "", 2)
	End If
	
	If Len(RContent)<5 Then
		ErrMsg = SysMsg(6)
		FoundErr=True
	End If
	If Len(RContent)>250 Then
		ErrMsg = SysMsg(7)
		FoundErr=True
	End If
	If EA_Pub.Chk_PostTime(30,"s",Session("lastpost")) Then
		ErrMsg = SysMsg(8)
		FoundErr=True
	End If

	If Not FoundErr Then Call EA_DBO.Set_Review_Insert(ArticleId,RUserId,RUserName,RContent,Ip,RState)

	Application.Lock 
	Application(sCacheName&"IsFlush")=1
	Application.UnLock 
	
	Call EA_Pub.ShowErrMsg(0,2)
End Sub
%>