<!--#Include File="conn.asp" -->
<!--#Include File="include/inc.asp"-->
<!--#Include File="include/page_column.asp"-->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：List.asp
'= 摘    要：列表页文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-02-27
'====================================================================

Dim PageContent
Dim ColumnId, ColumnInfo

'get column info
ColumnId	= EA_Pub.SafeRequest(3, "classid", 0, 0, 0)
ColumnInfo	= EA_DBO.Get_Column_Info(ColumnId)
If Not IsArray(ColumnInfo) Then Call EA_Pub.ShowErrMsg(9, 1)

'redirect the url
If ColumnInfo(6, 0) Then 
	Call EA_Pub.Close_Obj
	Set EA_Pub = Nothing

	Response.Redirect ColumnInfo(7, 0)
	Response.End 
End If

'check system state
If EA_Pub.SysInfo(18) = "0" Then
	PageContent = EA_Pub.Cov_ColumnPath(Request("classid"), EA_Pub.SysInfo(18))

	Call EA_Pub.Close_Obj
	Set EA_Pub = Nothing

	Response.Redirect PageContent
	Response.End 
End If

Dim clsColumn
Set clsColumn = New page_Column

PageContent = clsColumn.Make(ColumnId, ColumnInfo)

Response.Write PageContent

Call EA_Pub.Close_Obj
Set EA_Pub = Nothing
%>