<!--#Include File="comm/inc.asp" -->
<!--#Include File="../include/page_index.asp"-->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2004 - 2008 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：/Manager/Admin_MakeIndex.asp
'= 摘    要：后台-HTML首页生成文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2006-07-06
'====================================================================

Call EA_Manager.Chk_IsMaster

If Not EA_Manager.Chk_Power(Admin_Power,"41") Then 
	ErrMsg=str_Comm_NotAccess
	Call EA_Manager.Error(1)
End If

Dim Action
Action=Request.Form("action")

Select Case LCase(Action)
Case "make"
	Call MakeIndex
Case Else
	Call Main
End Select
Call EA_Pub.Close_Obj
Set EA_Pub=Nothing

Sub Main()
	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_MakeIndex_Help",str_MakeIndex_Help)
	Call EA_M_XML.AppElements("Language_MakeIndex_Info",str_MakeIndex_Info)
	Call EA_M_XML.AppElements("btnSubmit",str_MakeIndex_StartNow)

	Page = EA_M_XML.make("","",0)
	Call EA_M_XML.Out(Page)
End Sub

Sub MakeIndex()
	Dim PageContent
	Dim clsIndex

	Set clsIndex = New page_Index

	PageContent	 = clsIndex.Make()

	Call EA_Pub.Save_HtmlFile("../default.htm",PageContent)

	Response.Write "0"
	Response.End
End Sub
%>