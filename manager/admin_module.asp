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
'= 文件名称：/Manager/Admin_Module.asp
'= 摘    要：后台-管理文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-02-21
'====================================================================

Call EA_Manager.Chk_IsMaster

If Not EA_Manager.Chk_Power(Admin_Power,"44") Then 
	ErrMsg=str_Comm_NotAccess
	Call EA_Manager.Error(1)
End If

Dim Atcion
Atcion=Request ("action")

Select Case LCase(Atcion)
Case "add"
	Call Edit
Case "save"
	Call Save
Case "del"
	Call DelTheme
Case "default"
	Call SaveDefaultTheme
Case Else
	Call Main
End Select
Call EA_Pub.Close_Obj
Set EA_Pub=Nothing

Sub Main
	Dim i
	Dim TopicList
	Dim ListName(6),ListValue()
	Dim ForTotal,ThemeId

	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_Theme_ModuleInfo",str_Theme_ModuleInfo)

	Call EA_M_XML.AppElements("Language_Theme_ModuleEdit",str_Theme_ModuleAdd)
	Call EA_M_XML.AppElements("Language_Theme_ModuleName",str_Theme_ModuleName)
	Call EA_M_XML.AppElements("Language_Theme_ModuleDesc",str_Theme_ModuleDesc)
	Call EA_M_XML.AppElements("Language_Theme_ModuleType",str_Theme_ModuleType)
	Call EA_M_XML.AppElements("Language_Theme_ModuleName1",str_Theme_ModuleName)
	Call EA_M_XML.AppElements("Language_Theme_ModuleDesc1",str_Theme_ModuleDesc)
	Call EA_M_XML.AppElements("Language_Theme_ModuleType1",str_Theme_ModuleType)
	Call EA_M_XML.AppElements("Language_Theme_ModuleCode",str_Theme_ModuleCode)

	Call EA_M_XML.AppElements("Comm_Add_Operation",str_Comm_Add_Operation)
	Call EA_M_XML.AppElements("Comm_Del_Operation",str_Comm_Del_Operation)
	Call EA_M_XML.AppElements("btnReturn",str_Comm_Return_Button)
	Call EA_M_XML.AppElements("btnSubmit",str_Comm_Save_Button)
	Call EA_M_XML.AppElements("btnReset",str_Comm_Reset_Button)

	Call EA_M_XML.AppElements("Language_Comm_Bar_Operation",str_Comm_Bar_Operation)

	ThemeId=EA_Pub.SafeRequest(2,"ID",0,0,0)

	TopicList=EA_M_DBO.Get_Module_List(ThemeId)

	ListName(0) = "checkbox"
	ListName(1) = "ID"
	ListName(2) = "IDs"
	ListName(3) = "Name"
	ListName(4) = "Desc"
	ListName(5) = "Type"
	ListName(6) = "action"
	ForTotal = Ubound(TopicList,2)
	
	For i=0 To ForTotal
		ReDim Preserve ListValue(6,i)

		ListValue(0,i) = "checkbox"
		ListValue(1,i) = TopicList(0,i)
		ListValue(2,i) = TopicList(0,i)
		ListValue(3,i) = TopicList(1,i)
		ListValue(4,i) = TopicList(2,i)
		Select Case TopicList(5,i)
		Case 0
			ListValue(5,i) = str_Theme_ModuleHome
		End Select
		ListValue(6,i) = "action"
	Next

	Page = EA_M_XML.make(ListName,ListValue,Ubound(TopicList,2)+1)

	Call EA_M_XML.Out(Page)
End Sub

Sub Edit
	Dim ModuleID
	Dim ModuleName, ModuleDesc, ModuleType, ModuleCode
	Dim i, TempStr, Tmp
	
	ModuleID=EA_Pub.SafeRequest(2,"ID",0,0,0)
	Call EA_M_XML.AppInfo("ID",ModuleID)
	
	If ModuleID > 0 Then
		TempStr=EA_M_DBO.Get_Module_Info(ModuleID)

		Call EA_M_XML.AppInfo("Title",TempStr(1,0))
		Call EA_M_XML.AppInfo("Desc",TempStr(2,0))
		Call EA_M_XML.AppInfo("Code",TempStr(3,0))

		Tmp = "(build-select)," & TempStr(4,0) & " " & str_Theme_ModuleHome & ",0 " & str_Theme_ModuleCss & ",1 " & str_Theme_ModuleHead & ",2 " & str_Theme_ModuleFoot & ",3 " & str_Theme_ModulePage & ",4 " & str_Theme_ModuleContent & ",5"
		Call EA_M_XML.AppInfo("Typer",Tmp)

		Call EA_M_XML.AppElements("Language_Theme_ModuleEdit",str_Theme_ModuleEdit)
	Else
		Call EA_M_XML.AppElements("Language_Theme_ModuleEdit",str_Theme_ModuleAdd)

		Tmp = "(build-select),0 " & str_Theme_ModuleHome & ",0 " & str_Theme_ModuleCss & ",1 " & str_Theme_ModuleHead & ",2 " & str_Theme_ModuleFoot & ",3 " & str_Theme_ModulePage & ",4 " & str_Theme_ModuleContent & ",5"
		Call EA_M_XML.AppInfo("Typer",Tmp)
	End If

	Page = EA_M_XML.make("","",0)
	Call EA_M_XML.Out(Page)
End Sub

Sub Save
	If rs.state=1 Then rs.close

	Dim TemplateContent,TemplateTag
	Dim TemplateID
	
	TemplateID		= EA_Pub.SafeRequest(2,"ID",0,0,0)
	TemplateContent = EA_Pub.SafeRequest(2,"TemplateContent",1,"",-1)
	TemplateTag		= EA_Pub.SafeRequest(2,"TemplateTag",1,"",-1)

	If TemplateTag = "" Then Response.Write "-1":Response.End
	
	If TemplateID<>0 Then
		rs.open "select * from NB_Template where id="&TemplateID,conn,2,3
	Else
		rs.open "select * from NB_Template",conn,2,3
		rs.addnew
	End If
	
	If TemplateTag = "Name" Then
		rs("Temp_" & TemplateTag)=TemplateContent
	Else
		rs("Page_" & TemplateTag)=TemplateContent

		If LCase(TemplateTag) = "css" Then
			Dim re
			Set re=new RegExp
			re.IgnoreCase =true
			re.Global=True

			re.Pattern="<style([^>]*)>"
			TemplateContent=re.Replace(TemplateContent,"")

			TemplateContent=Replace(TemplateContent,"</style>","")

			TemplateContent=Replace(TemplateContent,"{$SystemPath$}",SystemFolder)

			Call EA_Pub.Save_HtmlFile("../common/css/style_" & TemplateID & ".css",TemplateContent)
		End If
	End If
	rs.update
		
	Rs.Close:Set Rs=Nothing
	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing
	
	Response.Write "1"
	Response.End
End Sub

Sub SaveDefaultTheme
	Dim DefId
	DefId=EA_Pub.SafeRequest(2,"ID",0,0,0)
	
	If DefId > 0 Then EA_M_DBO.Set_DefaultTheme DefId
	
	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing
	
	Response.Write "1"
	Response.End
End Sub

Sub DelTheme
	Dim IDs
	Dim i,Tmp
	Dim TempStr
	Dim IsDel
	Dim ForTotal

	IDs = Split(Request("ID"), ",")
	ForTotal = UBound(IDs)

	For i = 0 To ForTotal
		IsDel=False

		Tmp = EA_Pub.SafeRequest(5,IDs(i),0,0,0)

		TempStr=EA_M_DBO.Get_Theme_Info(Tmp)
		If IsArray(TempStr) Then 
			If TempStr(2,0)=0 Then 
				IsDel=True
			Else
				IsDel=False
			End If
		End If

		If IsDel Then EA_M_DBO.Set_Theme_Delete Tmp
	Next
	
	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing
	
	Response.Write "1"
	Response.End
End Sub 
%>