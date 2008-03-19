<!--#Include File="../init.asp"-->
<!--#Include File="../language/zh-cn.asp"-->
<!--#Include File="cls_public.asp"-->
<!--#Include File="cls_template.asp"-->
<!--#Include File="cls_ini.asp"-->
<!--#Include File="cls_dboperation.asp"-->
<%
'====================================================================
'= Team Elite - EliteCMS
'= Copyright (c) 2004 - 2008 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：Inc.asp
'= 摘    要：头文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-03-12
'====================================================================

Dim EA_Pub, EA_Temp, EA_DBO
Dim ErrMsg

Set EA_DBO = New cls_DBOperation
Set EA_Pub = New cls_Public

If EA_Pub.SysInfo(1) = "0" Then Call EA_Pub.ShowErrMsg(EA_Pub.SysInfo(2), 1)
If EA_Pub.SysInfo(3) = "1" And EA_Pub.Chk_SystemTimer(EA_Pub.SysInfo(4)) Then Call EA_Pub.ShowErrMsg(37, 0)

Set EA_Temp = New cls_Template
%>