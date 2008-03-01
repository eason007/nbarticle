<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：tag_column.asp
'= 摘    要：column模版标签文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-03-01
'====================================================================

Sub MakeColumn(ByRef PageContent)
	If EA_Temp.ChkBlock("Column.List", PageContent) Then
		ColumnList PageContent
	End If
End Sub

Sub ColumnList (ByRef PageContent)
	Dim Block, Parameter
	Dim List
	Dim Temp, ForTotal, i

	Do 
		Block = EA_Temp.GetBlock("Column.List", PageContent)
		If Block = "" Then Exit Do

		Parameter = EA_Temp.GetParameter("Parameter", Block)
		If Not IsArray(Parameter) Then EA_Temp.CloseBlock "Column.List", PageContent: Exit Do

		If CInt(Parameter(0)) = 0 Then
			List = EA_DBO.Get_Column_ChildList("")
		Else
			Temp = EA_DBO.Get_Column_Info(Parameter(0))
			List = EA_DBO.Get_Column_ChildList(Temp(1, 0))
		End If
		If Not IsArray(List) Then EA_Temp.CloseBlock "Column.List", PageContent: Exit Do
		
		ForTotal = UBound(List, 2)

		For i = 0 To ForTotal
			Temp = Block
	  
			EA_Temp.SetVariable "Title", List(1, i), Temp
			EA_Temp.SetVariable "Url", EA_Pub.Cov_ColumnPath(List(0, i), EA_Pub.SysInfo(18)), Temp
			EA_Temp.SetVariable "Info", List(4, i), Temp
			EA_Temp.SetVariable "Total", List(2, i), Temp

			EA_Temp.SetBlock "Column.List", Temp, PageContent
		Next

		EA_Temp.CloseBlock "Column.List", PageContent
	Loop While 1 = 1
End Sub
%>