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
'= 最后日期：2008-02-28
'====================================================================

Sub MakeFriendList (ByRef objTemplate, ByRef sPageContent, iColumnID)
	Dim TopicList
	Dim ForTotal
	Dim Temp, ListBlock
	Dim i

	TopicList = EA_DBO.Get_Friend_List(50, 0, 0)

	ListBlock = objTemplate.GetBlock("txt-link", sPageContent)

	If IsArray(TopicList)  Then
		ForTotal = UBound(TopicList, 2)

		For i = 0 To ForTotal
			Temp = ListBlock
	  
			objTemplate.SetVariable "Url", TopicList(2,i), Temp
			objTemplate.SetVariable "Img", TopicList(2,i), Temp
			objTemplate.SetVariable "Title", TopicList(0, i), Temp
			objTemplate.SetVariable "Info", TopicList(3, i), Temp

			objTemplate.SetBlock "txt-link", Temp, PageContent
		Next

		objTemplate.CloseBlock "txt-link", sPageContent
	End If


	TopicList = EA_DBO.Get_Friend_List(50, 0, 1)

	ListBlock = objTemplate.GetBlock("img-link", sPageContent)

	If IsArray(TopicList) Then
		ForTotal = UBound(TopicList, 2)

		For i = 0 To ForTotal
			Temp = ListBlock
	  
			objTemplate.SetVariable "Url", TopicList(2,i), Temp
			objTemplate.SetVariable "Img", TopicList(2,i), Temp
			objTemplate.SetVariable "Title", TopicList(0, i), Temp
			objTemplate.SetVariable "Info", TopicList(3, i), Temp

			objTemplate.SetBlock "img-link", Temp, PageContent
		Next

		objTemplate.CloseBlock "img-link", sPageContent
	End If
End Sub
%>