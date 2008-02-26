<%@ LANGUAGE = VBScript CodePage = 65001%>
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 文件名称：Coon.asp
'= 摘    要：数据库连接文件
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2007-03-10
'====================================================================
Option Explicit

Response.Charset= "UTF-8"
Response.Buffer	= True

Const iDataBaseType	= 0						'定义数据库类别，0为Access，1为SQL数据库，2=SQL Pro
Const BowelVersion	= "EAS-300-F-070310"
Const SysVersion	= "EliteArticle System Version 3.00 Beta2"

Dim StarTime,EndTime
Dim FoundErr

StarTime = Timer()
FoundErr = False
%>
<!--#include file="connection.asp"-->