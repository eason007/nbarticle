<!--#Include File="include/inc.asp"-->
<!--#Include File="include/page_index.asp"-->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：Index.asp
'= 摘    要：首页文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-03-12
'====================================================================

Dim MakeHtml

MakeHtml = False

If EA_Pub.SysInfo(18) = "0" Or EA_Pub.SysInfo(26) = "1" Then
	If Not EA_Pub.Chk_IsExistsHtmlFile("default.htm") Then 
		MakeHtml = True
	Else
		Call EA_Pub.Close_Obj
		Set EA_Pub = Nothing

		Response.Redirect "default.htm"
		Response.End 
	End If
End If

Dim PageContent
Dim clsIndex

Set clsIndex = New page_Index

PageContent	 = clsIndex.Make()

If MakeHtml Then 
	Call EA_Pub.Save_HtmlFile("default.htm", PageContent)
	
	Call EA_Pub.Close_Obj
	Set EA_Pub = Nothing

	Response.Redirect "default.htm"
	Response.End 
Else
	Response.Write PageContent
End If

Call EA_Pub.Close_Obj
Set EA_Pub = Nothing
%>