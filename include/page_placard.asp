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

Sub MakePlacardList (ByRef objTemplate, ByRef sPageContent)
	Dim RCount
	Dim PlacardArray
	Dim ForTotal
	Dim Temp, ListBlock, i

	RCount = EA_DBO.Get_PlacardStat()(0, 0)

	ListBlock = objTemplate.GetBlock("placard", sPageContent)

	If RCount > 0 Then 
		PlacardArray = EA_DBO.Get_PlacardList(1, 100)

		ForTotal = UBound(PlacardArray, 2)

		For i = 0 To ForTotal
			Temp = ListBlock
	  
			objTemplate.SetVariable "Title", PlacardArray(1, i), Temp
			objTemplate.SetVariable "Content", PlacardArray(4, i), Temp
			objTemplate.SetVariable "AddTime", PlacardArray(2, i), Temp
			objTemplate.SetVariable "OverTime", PlacardArray(3, i), Temp

			objTemplate.SetBlock "placard", Temp, sPageContent
		Next

		objTemplate.CloseBlock "placard", sPageContent
	End If
End Sub
%>