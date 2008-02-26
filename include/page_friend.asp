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

Class page_Friend
	Private Function MakePlacardList ()
		Dim RCount
		Dim PlacardArray
		Dim ForTotal
		Dim Temp,ListBlock


			TopicList=EA_DBO.Get_FriendList_All()

			If IsArray(TopicList) Then
				ForTotal = UBound(PlacardArray,2)

				For i=0 To ForTotal
					Temp = ListBlock
			  
					Template.SetVariable "Url", TopicList(2,i), Temp
					Template.SetVariable "Img", TopicList(4,i), Temp
					Template.SetVariable "Title", TopicList(5, i), Temp
					Template.SetVariable "Info", TopicList(3, i), Temp

					Template.SetBlock "placard", Temp, PageContent
				Next

				Template.CloseBlock "placard", PageContent
			End If
	End Function
End Class
%>