<!--#Include File="include/inc.asp"-->
<!--#Include File="include/page_column.asp"-->
<%
'====================================================================
'= Team Elite - EliteCMS
'= Copyright (c) 2004 - 2008 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：List.asp
'= 摘    要：列表页文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-03-18
'====================================================================

Dim PageContent
Dim ColumnId, ColumnInfo, PageNum

'get column info
ColumnId	= EA_Pub.SafeRequest(3, "classid", 0, 0, 0)
ColumnInfo	= EA_DBO.Get_Column_Info(ColumnId)
PageNum		= EA_Pub.SafeRequest(3, "page", 0, 1, 0)
If Not IsArray(ColumnInfo) Then Call EA_Pub.ShowErrMsg(35, 0)

'redirect the url
If ColumnInfo(6, 0) Then 
	Call EA_Pub.Close_Obj
	Set EA_Pub = Nothing

	Response.Redirect ColumnInfo(7, 0)
	Response.End 
End If

'check system state
If EA_Pub.SysInfo(18) = "0" Then
	PageContent = EA_Pub.Cov_ColumnPath(ColumnId, EA_Pub.SysInfo(18))

	Call EA_Pub.Close_Obj
	Set EA_Pub = Nothing

	Response.Redirect PageContent
	Response.End 
End If

Dim clsColumn
Set clsColumn = New page_Column

Call clsColumn.Make(ColumnId, ColumnInfo, PageNum)

Response.Write clsColumn.PageContent

Call EA_Pub.Close_Obj
Set EA_Pub = Nothing
%>