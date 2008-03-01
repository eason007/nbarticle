<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：tag_placard.asp
'= 摘    要：placard模版标签文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-03-01
'====================================================================

Sub MakePlacard(ByRef PageContent)
	If EA_Temp.ChkBlock("Placard.List", PageContent) Then
		PlacardList PageContent
	End If
End Sub

Sub PlacardList (ByRef PageContent)
	Dim Block, Parameter
	Dim List
	Dim Temp, ForTotal, i

	Do 
		Block = EA_Temp.GetBlock("Placard.List", PageContent)
		If Block = "" Then Exit Do

		Parameter = EA_Temp.GetBlockParameter(Block)
		If Not IsArray(Parameter) Then EA_Temp.CloseBlock "Placard.List", PageContent: Exit Do

		If InStr(1, Parameter(0), "ID=") Then
			List = EA_DBO.Get_PlacardInfo(Replace(Parameter(0), "ID=", ""))
		Else
			List = EA_DBO.Get_PlacardTopList(Parameter(0))
		End If
		If Not IsArray(List) Then EA_Temp.CloseBlock "Placard.List", PageContent: Exit Do
		
		ForTotal = UBound(List, 2)

		For i = 0 To ForTotal
			Temp = Block
	  
			EA_Temp.SetVariable "Title", List(1, i), Temp
			EA_Temp.SetVariable "Content", List(4, i), Temp
			EA_Temp.SetVariable "AddTime", List(2, i), Temp
			EA_Temp.SetVariable "OverTime", List(3, i), Temp

			EA_Temp.SetBlock "Placard.List", Temp, PageContent
		Next

		EA_Temp.CloseBlock "Placard.List", PageContent
	Loop While 1 = 1
End Sub
%>