<%
'====================================================================
'= Team Elite - EliteCMS
'= Copyright (c) 2004 - 2008 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：page_comment.asp
'= 摘    要：评论页文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-04-17
'====================================================================

Class page_Comment
	Public PageContent
	Private Info, ID

	Public Function Make (iID, vInfo, iPage)
		ID	 = iID
		Info = vInfo

		PageContent = EA_Temp.Load_Template(0, 7)

		EA_Temp.Title	= EA_Pub.SysInfo(0) & " - " & SysMsg(54)
		EA_Temp.Nav		= "<a href=""" & SystemFolder & """>" & EA_Pub.SysInfo(0) & "</a>" & EA_Pub.Get_NavByColumnCode(Info(1, 0), 0) & " - <a href=""" & EA_Pub.Cov_ArticlePath(ID, Info(13, 0), EA_Pub.SysInfo(18)) & """>" & Info(3, 0) & "</a> - <strong>" & SysMsg(54) & "</strong>"

		If EA_Temp.ChkBlock("Comment.Topic", PageContent) Then Call CommentTopic(iPage)

		PageContent = EA_Temp.Replace_PublicTag(PageContent)

		Make = PageContent
	End Function

	Private Sub CommentTopic (iPage)
		Dim CommentList
		Dim ListBlock, ForTotal, i, Temp

		CommentList = EA_DBO.Get_CommentList(ID, iPage)

		ListBlock = EA_Temp.GetBlock("Comment.Topic", PageContent)

		If IsArray(CommentList) Then
			ForTotal = UBound(CommentList, 2)

			For i = 0 To ForTotal
				Temp = ListBlock
		  
				EA_Temp.SetVariable "Content", CommentList(0, i), Temp
				EA_Temp.SetVariable "UserName", CommentList(1, i), Temp
				EA_Temp.SetVariable "Date", CommentList(2, i), Temp

				EA_Temp.SetBlock "Comment.Topic", Temp, PageContent
			Next
		End If

		EA_Temp.CloseBlock "Comment.Topic", PageContent
	End Sub
End Class
%>