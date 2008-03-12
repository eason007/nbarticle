<!--#Include File="comm/inc.asp" -->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：/Manager/Admin_Main.asp
'= 摘    要：后台-控制台文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-02-18
'====================================================================

Call EA_Manager.Chk_IsMaster

Dim theInstalledObjects(4)
theInstalledObjects(0) = "Scripting.FileSystemObject"
theInstalledObjects(1) = "adodb.connection"
theInstalledObjects(2) = "JMail.SMTPMail"
theInstalledObjects(3) = "CDONTS.NewMail"

Dim RegUser,TopicNum,ColumnNum,MangerTopicNum,ReviewNum
Dim UserGroup_Array,GroupList
Dim i

RegUser=0
TopicNum=0
ColumnNum=0
MangerTopicNum=0
ReviewNum=0

Sql="select reguser,topicnum,ColumnNum,MangerTopicNum,ReviewNum from [nb_system]"
set rs=conn.execute(sql)
If Not rs.eof And Not rs.bof Then 
	RegUser=rs(0)
	TopicNum=rs(1)
	ColumnNum=rs(2)
	MangerTopicNum=rs(3)
	ReviewNum=rs(4)
End If

Rs.Close
Set Rs=Nothing


Call EA_M_XML.AppElements("Language_SystemInformation",str_SystemInformation)
Call EA_M_XML.AppElements("Language_SystemStat",str_SystemStat)
Call EA_M_XML.AppElements("Language_Column",str_Column)
Call EA_M_XML.AppElements("Language_Article",str_Article)
Call EA_M_XML.AppElements("Language_AuditArticle",str_AuditArticle)
Call EA_M_XML.AppElements("Language_RegUser",str_RegUser)
Call EA_M_XML.AppElements("Language_Review",str_Review)
Call EA_M_XML.AppElements("Language_SystemOwner",str_SystemOwner)
Call EA_M_XML.AppElements("Language_SystemVersion",str_SystemVersion)
Call EA_M_XML.AppElements("Language_ServerWindow",str_ServerWindow)
Call EA_M_XML.AppElements("Language_ScripEngine",str_ScripEngine)
Call EA_M_XML.AppElements("Language_SiteFolderPath",str_SiteFolderPath)
Call EA_M_XML.AppElements("Language_FSOEnable",str_FSOEnable)
Call EA_M_XML.AppElements("Language_ADOEnable",str_ADOEnable)
Call EA_M_XML.AppElements("Language_JMailEnable",str_JMailEnable)
Call EA_M_XML.AppElements("Language_CDONTSEnable",str_CDONTSEnable)
Call EA_M_XML.AppElements("Language_MoreSystemInformation",str_MoreSystemInformation)
Call EA_M_XML.AppElements("Language_SystemManagerShortcut",str_SystemManagerShortcut)
Call EA_M_XML.AppElements("Language_QuickSearchArticle",str_QuickSearchArticle)
Call EA_M_XML.AppElements("Language_SearchNow",str_SearchNow)
Call EA_M_XML.AppElements("Language_QuickSearchUser",str_QuickSearchUser)
Call EA_M_XML.AppElements("Language_UserGroup",str_UserGroup)
Call EA_M_XML.AppElements("Language_FunctionShortcut",str_FunctionShortcut)
Call EA_M_XML.AppElements("Language_ColumnAdmin",str_ColumnAdmin)
Call EA_M_XML.AppElements("Language_ArticleAdmin",str_ArticleAdmin)
Call EA_M_XML.AppElements("Language_ReLoadCache",str_ReLoadCache)
Call EA_M_XML.AppElements("Language_ProductInformation",str_ProductInformation)
Call EA_M_XML.AppElements("Language_ProductCopyright",str_ProductCopyright)
Call EA_M_XML.AppElements("Language_AboutMe",str_AboutMe)
Call EA_M_XML.AppElements("Language_ProductSales",str_ProductSales)
Call EA_M_XML.AppElements("Language_UseGuide",str_UseGuide)
Call EA_M_XML.AppElements("Language_Thruway",str_Thruway)
Call EA_M_XML.AppElements("ColumnTotal",ColumnNum)
Call EA_M_XML.AppElements("TopicTotal",TopicNum)
Call EA_M_XML.AppElements("MangerTopicTotal",MangerTopicNum)
Call EA_M_XML.AppElements("UserTotal",RegUser)
Call EA_M_XML.AppElements("CommentTotal",ReviewNum)
Call EA_M_XML.AppElements("SystemOwner",EA_Pub.SysInfo(0))
Call EA_M_XML.AppElements("SystemVersion",SysVersion)
Call EA_M_XML.AppElements("Request_OS",Request.ServerVariables("OS"))
Call EA_M_XML.AppElements("Request_Local",Request.ServerVariables("LOCAL_ADDR"))
Call EA_M_XML.AppElements("Request_Path",Request.ServerVariables("APPL_PHYSICAL_PATH"))
Call EA_M_XML.AppElements("ScriptEngine",ScriptEngine & " "& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion)

If iDataBaseType=0 Then
	Call EA_M_XML.AppElements("DatabaseVersion","Access")
Else
	Call EA_M_XML.AppElements("DatabaseVersion","MS SQL Server 2000")
End If
If Not EA_Manager.IsObjInstalled(theInstalledObjects(0)) Then
	Call EA_M_XML.AppElements("FSOEnable",str_IconDisabled)
Else
	Call EA_M_XML.AppElements("FSOEnable",str_IconEnabled)
End If
If Not EA_Manager.IsObjInstalled(theInstalledObjects(1)) Then
	Call EA_M_XML.AppElements("ADOEnable",str_IconDisabled)
Else
	Call EA_M_XML.AppElements("ADOEnable",str_IconEnabled)
End If
If Not EA_Manager.IsObjInstalled(theInstalledObjects(2)) Then
	Call EA_M_XML.AppElements("JMailEnable",str_IconDisabled)
Else
	Call EA_M_XML.AppElements("JMailEnable",str_IconEnabled)
End If
If Not EA_Manager.IsObjInstalled(theInstalledObjects(3)) Then
	Call EA_M_XML.AppElements("CDONTSEnable",str_IconDisabled)
Else
	Call EA_M_XML.AppElements("CDONTSEnable",str_IconEnabled)
End If

Page = EA_M_XML.make("","",0)

Call EA_M_XML.Out(Page)

Call EA_Pub.Close_Obj
Set EA_Pub=Nothing
%>