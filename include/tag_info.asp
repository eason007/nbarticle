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
'= 最后日期：2008-03-01
'====================================================================

Sub MakeInfo(ByRef PageContent)
	EA_Temp.SetVariable "Info.ColumnTotal", EA_Pub.SysStat(0), PageContent
	EA_Temp.SetVariable "Info.TopicTotal", EA_Pub.SysStat(1), PageContent
	EA_Temp.SetVariable "Info.UserTotal", EA_Pub.SysStat(3), PageContent
	EA_Temp.SetVariable "Info.CommentTotal", EA_Pub.SysStat(4), PageContent
	EA_Temp.SetVariable "Info.SiteName", EA_Pub.SysInfo(0), PageContent
	EA_Temp.SetVariable "Info.SiteUrl", EA_Pub.SysInfo(11), PageContent
	EA_Temp.SetVariable "Info.SiteMail", EA_Pub.SysInfo(12), PageContent
	EA_Temp.SetVariable "Info.Version", SysVersion, PageContent
	EA_Temp.SetVariable "Info.ThemeName", EA_Temp.PageArray(0), PageContent
End Sub
%>