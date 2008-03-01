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

		Parameter = EA_Temp.GetBlockParameter(Block)
		List = EA_DBO.Get_Friend_List(Parameter(1), Parameter(0), Parameter(2))
		
		ForTotal = UBound(List, 2)

		For i = 0 To ForTotal
			Temp = Block
	  
			EA_Temp.SetVariable "Title", List(1, i), Temp
			EA_Temp.SetVariable "Url", EA_Pub.Cov_ColumnPath(List(0, i), EA_Pub.SysInfo(18)), Temp
			EA_Temp.SetVariable "Info", List(4, i), Temp
			EA_Temp.SetVariable "Total", List(2, i), Temp

			EA_Temp.SetBlock "Friend.List", Temp, PageContent
		Next

		EA_Temp.CloseBlock "Friend.List", PageContent
	Loop While 1 = 1
End Sub
%>