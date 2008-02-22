<!--#Include File="conn.asp" -->
<!--#Include File="include/inc.asp"-->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：Index.asp
'= 摘    要：首页文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2005-10-20
'====================================================================

Dim PageContent
Dim MakeHtml

MakeHtml=False

If EA_Pub.SysInfo(18)="0" Then
	If EA_Pub.SysInfo(26)="0" Then
		EA_Pub.SysInfo(18)="1"
	Else
		If Not EA_Pub.Chk_IsExistsHtmlFile("default.htm") Then 
			MakeHtml=True
		Else
			Call EA_Pub.Close_Obj
			Set EA_Pub=Nothing

			Response.Redirect "default.htm"
			Response.End 
		End If
	End If
End If

PageContent=EA_Temp.Load_Template(0, 0)

EA_Temp.Title=EA_Pub.SysInfo(0)&" - 首页"
EA_Temp.Nav="<a href=""./""><b>"&EA_Pub.SysInfo(0)&"</b></a> - 首页"

EA_Temp.ReplaceTag "SiteColumnTotal",EA_Pub.SysStat(0),PageContent
EA_Temp.ReplaceTag "SiteTopicTotal",EA_Pub.SysStat(1),PageContent
EA_Temp.ReplaceTag "SiteUserTotal",EA_Pub.SysStat(3),PageContent
EA_Temp.ReplaceTag "SiteMangerTopicTotal",EA_Pub.SysStat(2),PageContent
EA_Temp.ReplaceTag "SiteReviewTotal",EA_Pub.SysStat(4),PageContent

EA_Temp.ReplaceTag "MemberTopPost",EA_Temp.Load_MemberTopPost,PageContent

Call EA_Temp.Find_TemplateTag("ColumnNav",PageContent)
Call EA_Temp.Find_TemplateTag("DisList",PageContent)
Call EA_Temp.Find_TemplateTag("NewReview",PageContent)

PageContent=EA_Temp.Replace_PublicTag(PageContent)

If MakeHtml Then 
	Call EA_Pub.Save_HtmlFile("default.htm",PageContent)
	
	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing

	Response.Redirect "default.htm"
	Response.End 
Else
	Response.Write PageContent
End If

Call EA_Pub.Close_Obj
Set EA_Pub=Nothing
%>