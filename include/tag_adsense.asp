<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2004 - 2008 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：tag_adsense.asp
'= 摘    要：adsense模版标签文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-03-18
'====================================================================

Sub MakeAdSense(ByRef PageContent)
	If EA_Temp.ChkTag("AdSense.Single", PageContent) Then
		AdSenseSingle PageContent
	End If
End Sub

Sub AdSenseSingle (ByRef PageContent)
	Dim Parameter
	Dim List

	Do 
		Parameter = EA_Temp.GetParameter("AdSense.Single", PageContent)
		If Not IsArray(Parameter) Then Exit Do

		List = EA_DBO.Get_AdSense_Info(Parameter(0))
		If Not IsArray(List) Then 
			Exit Do
		Else
			Call EA_Temp.SetSingle(List(1, 0), PageContent)
		End If
	Loop While 1 = 1
End Sub
%>