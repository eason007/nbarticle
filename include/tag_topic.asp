<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2004 - 2008 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：tag_topic.asp
'= 摘    要：topic模版标签文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-03-02
'====================================================================

Sub MakeTopic(ByRef PageContent)
	If EA_Temp.ChkBlock("Topic.List", PageContent) Then
		TopicList PageContent
	End If
End Sub

Sub TopicList (ByRef PageContent)
	Dim Block, Parameter
	Dim List
	Dim Temp, ForTotal, i

	Do 
		Block = EA_Temp.GetBlock("Topic.List", PageContent)
		If Block = "" Then Exit Do

		Parameter = EA_Temp.GetParameter("Parameter", Block)
		If Not IsArray(Parameter) Then EA_Temp.CloseBlock "Topic.List", PageContent: Exit Do

		List = EA_DBO.Get_Article_List(Parameter(0), Parameter(1), Parameter(2), Parameter(4))
		If Not IsArray(List) Then EA_Temp.CloseBlock "Topic.List", PageContent: Exit Do
		
		ForTotal = UBound(List, 2)

		For i = 0 To ForTotal
			Temp = Block
	  
			List(3, i) = EA_Pub.Base_HTMLFilter(List(3, i))
			List(3, i) = EA_Pub.Cut_Title(List(3, i), Parameter(3))

			If Len(List(12, i)) > 0 Then List(11, i) = "<a href='" & List(12, i) & "'>" & List(11, i) & "</a>"

			EA_Temp.SetVariable "Title", EA_Pub.Add_ArticleColor(List(4, i), List(3, i)), Temp
			EA_Temp.SetVariable "Url", EA_Pub.Cov_ArticlePath(List(0, i), List(5, i), EA_Pub.SysInfo(18)), Temp
			EA_Temp.SetVariable "SubTitle", List(11, i), Temp
			EA_Temp.SetVariable "Img", List(8, i), Temp
			EA_Temp.SetVariable "Date", FormatDateTime(List(5, i), 2), Temp
			EA_Temp.SetVariable "Time", FormatDateTime(List(5, i), 4), Temp
			EA_Temp.SetVariable "Author", List(9, i), Temp
			EA_Temp.SetVariable "Icon", EA_Pub.Chk_ArticleType(List(6, i), List(7, i)), Temp
			EA_Temp.SetVariable "ColumnName", List(2, i), Temp
			EA_Temp.SetVariable "ColumnUrl", EA_Pub.Cov_ColumnPath(List(1, i), EA_Pub.SysInfo(18)), Temp

			EA_Temp.SetBlock "Topic.List", Temp, PageContent
		Next

		EA_Temp.CloseBlock "Topic.List", PageContent
	Loop While 1 = 1
End Sub
%>