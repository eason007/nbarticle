<%@ LANGUAGE = VBScript CodePage = 65001%>
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2004 - 2008 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 文件名称：Coon.asp
'= 摘    要：数据库连接文件
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-03-02
'====================================================================
Option Explicit

Response.Charset= "UTF-8"
Response.Buffer	= True

Const iDataBaseType	= 0						'定义数据库类别，0为Access，1为SQL数据库，2=SQL Pro
Const BowelVersion	= "ECMS-301-F-080226"
Const SysVersion	= "EliteCMS Ver 3.01 Beta1"

Dim StarTime, EndTime
Dim FoundErr

StarTime = Timer()
FoundErr = False
%>
<!--#Include File="connection.asp"-->