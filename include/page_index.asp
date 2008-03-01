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
'= 最后日期：2008-02-27
'====================================================================

Class page_Index
	Public Function Make ()
		Dim PageContent

		PageContent = EA_Temp.Load_Template(0, 0)

		EA_Temp.Title	= EA_Pub.SysInfo(0) & " - 首页"
		EA_Temp.Nav		= "<a href=""./""><b>" & EA_Pub.SysInfo(0) & "</b></a> - 首页"

		

		If EA_Temp.ChkTag("MemberTopPost", PageContent) Then EA_Temp.SetVariable "MemberTopPost", EA_Temp.Load_MemberTopPost, PageContent

		If EA_Temp.ChkTag("NewReview", PageContent) Then EA_Temp.Find_TemplateTag "NewReview", PageContent

		PageContent = EA_Temp.Replace_PublicTag(PageContent)

		Make = PageContent
	End Function
End Class
%>