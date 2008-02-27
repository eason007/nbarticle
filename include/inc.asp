<!--#Include File="cls_public.asp"-->
<!--#Include File="cls_template.asp"-->
<!--#Include File="cls_ini.asp"-->
<!--#Include File="cls_dboperation.asp"-->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：Inc.asp
'= 摘    要：头文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-02-27
'====================================================================

Dim EA_Pub, EA_Temp, EA_DBO
Dim ErrMsg

Set EA_DBO = New cls_DBOperation
Set EA_Pub = New cls_Public

If EA_Pub.SysInfo(1) = "0" Then
	ErrMsg = EA_Pub.SysInfo(2)
	Call EA_Pub.ShowErrMsg(0, 0)
End If

If EA_Pub.SysInfo(3) = "1" Then 
	If EA_Pub.Chk_SystemTimer(EA_Pub.SysInfo(4)) Then Call EA_Pub.ShowErrMsg(0, 0)
End If

Set EA_Temp = New cls_Template
%>