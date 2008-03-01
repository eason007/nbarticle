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
'= 最后日期：2008-02-24
'====================================================================

Sub MakeComment(ByRef PageContent)
	If EA_Temp.ChkBlock("Comment.List", PageContent) Then
		CommentList PageContent
	End If
End Sub

Sub CommentList (ByRef PageContent)
	Dim Block, Parameter
	Dim List
	Dim Temp, ForTotal, i

	Do 
		Block = EA_Temp.GetBlock("Comment.List", PageContent)
		If Block = "" Then Exit Do

		Parameter = EA_Temp.GetBlockParameter(Block)
		If Not IsArray(Parameter) Then EA_Temp.CloseBlock "Comment.List", PageContent: Exit Do

		List = EA_DBO.Get_Review_List(Parameter(0), Parameter(1), Parameter(2))
		If Not IsArray(List) Then EA_Temp.CloseBlock "Comment.List", PageContent: Exit Do
		
		ForTotal = UBound(List, 2)

		For i = 0 To ForTotal
			Temp = Block
	  
			EA_Temp.SetVariable "UserName", List(2, i), Temp
			EA_Temp.SetVariable "Content", List(1, i), Temp
			EA_Temp.SetVariable "Date", List(3, i), Temp
			EA_Temp.SetVariable "ArticleUrl", EA_Pub.Cov_ArticlePath(List(0, i), List(4, i), EA_Pub.SysInfo(18)), Temp
			EA_Temp.SetVariable "ArticleTitle", List(5, i), Temp

			EA_Temp.SetBlock "Comment.List", Temp, PageContent
		Next

		EA_Temp.CloseBlock "Comment.List", PageContent
	Loop While 1 = 1
End Sub
%>