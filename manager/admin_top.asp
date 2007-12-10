<!--#Include File="../conn.asp" -->
<!--#Include File="comm/inc.asp" -->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：/Manager/Admin_Top.asp
'= 摘    要：后台-头部控制菜单文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2006-05-22
'====================================================================

Call EA_Manager.Chk_IsMaster

Call EA_M_XML.AppElements("Language_HelpCenter",str_HelpCenter)
Call EA_M_XML.AppElements("Language_SiteIndexLink",str_SiteIndexLink)
Call EA_M_XML.AppElements("Language_ControlPanel",str_ControlPanel)
Call EA_M_XML.AppElements("Language_Exit",str_Exit)
Call EA_M_XML.AppElements("SysVersion",SysVersion)

Page = EA_M_XML.make("","",0)

Call EA_M_XML.Out(Page)

Call EA_Pub.Close_Obj
Set EA_Pub=Nothing
%>
