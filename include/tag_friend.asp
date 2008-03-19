<%
'====================================================================
'= Team Elite - EliteCMS
'= Copyright (c) 2004 - 2008 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：tag_friend.asp
'= 摘    要：friend模版标签文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-03-02
'====================================================================

Sub MakeFriend(ByRef PageContent)
	If EA_Temp.ChkBlock("Friend.List", PageContent) Then
		FriendList PageContent
	End If
End Sub

Sub FriendList (ByRef PageContent)
	Dim Block, Parameter
	Dim List
	Dim Temp, ForTotal, i

	Do 
		Block = EA_Temp.GetBlock("Friend.List", PageContent)
		If Block = "" Then Exit Do

		Parameter = EA_Temp.GetParameter("Parameter", Block)
		If Not IsArray(Parameter) Then EA_Temp.CloseBlock "Friend.List", PageContent: Exit Do

		List = EA_DBO.Get_Friend_List(Parameter(0), Parameter(1), Parameter(2))
		If Not IsArray(List) Then EA_Temp.CloseBlock "Friend.List", PageContent: Exit Do
		
		ForTotal = UBound(List, 2)

		For i = 0 To ForTotal
			Temp = Block
	  
			EA_Temp.SetVariable "Name", List(0, i), Temp
			EA_Temp.SetVariable "Url", List(1, i), Temp
			EA_Temp.SetVariable "Img", List(2, i), Temp
			EA_Temp.SetVariable "Info", List(3, i), Temp

			EA_Temp.SetBlock "Friend.List", Temp, PageContent
		Next

		EA_Temp.CloseBlock "Friend.List", PageContent
	Loop While 1 = 1
End Sub
%>