<!--#Include File="comm/inc.asp" -->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：/Manager/Admin_MakeIndex.asp
'= 摘    要：后台-HTML首页生成文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2006-07-06
'====================================================================

Call EA_Manager.Chk_IsMaster

If Not EA_Manager.Chk_Power(Admin_Power,"41") Then 
	ErrMsg=str_Comm_NotAccess
	Call EA_Manager.Error(1)
End If

Dim Action
Action=Request.Form("action")

Select Case LCase(Action)
Case "make"
	Call MakeIndex
Case Else
	Call Main
End Select
Call EA_Pub.Close_Obj
Set EA_Pub=Nothing

Sub Main()
	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_MakeIndex_Help",str_MakeIndex_Help)

	Call EA_M_XML.AppElements("Language_MakeIndex_Info",str_MakeIndex_Info)

	Call EA_M_XML.AppElements("btnSubmit",str_MakeIndex_StartNow)

	Page = EA_M_XML.make("","",0)
	Call EA_M_XML.Out(Page)
End Sub

Sub MakeIndex()
	Set EA_Temp=New cls_Template

	PageContent=EA_Temp.Load_Template(0,"index")

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

	Call EA_Pub.Save_HtmlFile("../default.htm",PageContent)

	Response.Write "0"
	Response.End
End Sub
%>