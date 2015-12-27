<!--#Include File="../../init.asp"-->
<!--#Include File="../../include/cls_public.asp"-->
<!--#Include File="../../include/cls_ini.asp"-->
<!--#Include File="../../include/cls_dboperation.asp"-->
<!--#Include File="../../include/cls_template.asp"-->
<!--#Include File="../../language/zh-cn.asp"-->
<!--#include file="cls_manager.asp"-->
<!--#include file="cls_manager_db.asp"-->
<!--#include file="cls_xml.asp"-->
<!--#include file="../language_files/comm.asp"-->
<%
'====================================================================
'= Team Elite - EliteCMS
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：/Manager/Comm/Inc.asp
'= 摘    要：后台-头文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-03-12
'====================================================================
Session.Timeout = 30

Dim Admin_Power,Column_Power
Dim Login_Key
Dim EA_Pub,EA_Manager,EA_Temp,EA_DBO,EA_M_DBO,EA_M_XML
Dim PageContent
Dim SQL, Rs, Conn
Dim Page, ErrMsg

Set Rs=Server.CreateObject("adodb.recordSet")

Set EA_DBO=New cls_DBOperation
Set EA_Pub=New cls_Public
Set EA_Manager=New cls_Manager
Set EA_M_DBO=New Cls_Manager_DBOperation
Set EA_M_XML=New Cls_XML
Set EA_Temp = New cls_Template
%>