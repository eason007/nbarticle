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
Response.Charset="UTF-8"
Response.Buffer=True

Const iDataBaseType=0						'定义数据库类别，0为Access，1为SQL数据库，2=SQL Pro
Const bIsShowRunTime=1
Const BowelVersion="EAS-300-F-070310"
Const SysVersion="EliteArticle System Version 3.00 Beta2"

Dim Conn
Dim StarTime,EndTime
Dim FoundErr,ErrMsg

StarTime=Timer()
FoundErr=False
ErrMsg=""

%>
<!--#include file="connection.asp"-->
<%
Function ConnectionDatabase
	On Error Resume Next

	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open ConnStr
	If Err Then	
		Response.Clear
		CloseDataBase
		Response.Write "对不起，数据连接错误！如果第一次使用，请先运行setup.asp进行系统配置。"
		Err.Clear
		Response.End
	End If
End Function

Function CloseDataBase
	If IsObject(EA_Temp) Then 
		EA_Temp.Close_Obj
		Set EA_Temp=Nothing
	End If
	If IsObject(Rs) Then
		If Rs.State=1 Then Rs.Close
		Set Rs=Nothing
	End If

	If Conn.State=1 Then Conn.Close
	Set Conn = Nothing
End Function
%>
