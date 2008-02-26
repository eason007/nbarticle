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

Class page_Placard
	Private Function MakePlacardList ()
		Dim RCount
		Dim PlacardArray
		Dim ForTotal
		Dim Temp,ListBlock

		RCount=EA_DBO.Get_PlacardStat()(0,0)

		ListBlock	= Template.GetBlock("placard", PageContent)

		If RCount>0 Then 
			PlacardArray=EA_DBO.Get_PlacardList(1, 100)
			ForTotal = UBound(PlacardArray,2)

			For i=0 To ForTotal
				Temp = ListBlock
		  
				Template.SetVariable "ID", "viewplacard.asp?postid="&PlacardArray(0,i), Temp
				Template.SetVariable "Title", PlacardArray(1,i), Temp
				Template.SetVariable "AddTime", PlacardArray(2, i), Temp
				Template.SetVariable "OverTime", PlacardArray(3, i), Temp

				Template.SetBlock "placard", Temp, PageContent
			Next

			Template.CloseBlock "placard", PageContent
		End If
	End Function
End Class
%>