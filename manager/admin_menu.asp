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
'= 文件名称：/Manager/Admin_Left.asp
'= 摘    要：后台-左边控制菜单文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-02-16
'====================================================================

Call EA_Manager.Chk_IsMaster

Dim i,j

Call EA_Pub.Close_Obj
Set EA_Pub=Nothing

Response.Write "var mainMenu = new Array();" & Chr(10) & Chr(13)
Response.Write "var subMenu = '';" & Chr(10) & Chr(13)

For i=0 To Ubound(str_LeftMenu)-1
	Response.Write "subMenu = '&nbsp;';" & Chr(10) & Chr(13)

	For j=1 To Ubound(str_LeftMenu,2)
		If IsEmpty(str_LeftMenu(i,j)) Then Exit For
		Response.Write "subMenu += '[" & str_LeftMenu(i,j) & "]&nbsp;&nbsp;&nbsp;';" & Chr(10) & Chr(13)
	Next

	Response.Write "mainMenu[" & i & "] = new menuFormat ('" & str_LeftMenu(i,0) & "', subMenu);" & Chr(10) & Chr(13)
Next
%>
