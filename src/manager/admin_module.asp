<!--#Include File="comm/inc.asp" -->
<%
'====================================================================
'= Team Elite - EliteCMS
'= Copyright (c) 2004 - 2008 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：/Manager/Admin_Module.asp
'= 摘    要：后台-模块管理文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-04-17
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
	Call Del
Case Else
	Call Main
End Select
Call EA_Pub.Close_Obj
Set EA_Pub=Nothing

Sub Main
	Dim i, Count
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

	Count=EA_M_DBO.Get_Module_Total(ThemeId)(0,0)
	If Count>0 Then 
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
			Case 1
				ListValue(5,i) = str_Theme_ModuleCss
			Case 2
				ListValue(5,i) = str_Theme_ModuleHead
			Case 3
				ListValue(5,i) = str_Theme_ModuleFoot
			Case 4
				ListValue(5,i) = str_Theme_ModulePage
			Case 5
				ListValue(5,i) = str_Theme_ModuleContent
			Case 6
				ListValue(5,i) = str_Theme_ModuleSearch
			Case 7
				ListValue(5,i) = str_Theme_ModuleComment
			End Select
			ListValue(6,i) = "action"
		Next

		Page = EA_M_XML.make(ListName,ListValue,Ubound(TopicList,2)+1)
	Else
		Page = EA_M_XML.make("","",0)
	End If

	Call EA_M_XML.Out(Page)
End Sub

Sub Edit
	Dim ModuleID
	Dim ModuleName, ModuleDesc, ModuleType, ModuleCode
	Dim i, TempStr, Tmp
	
	ModuleID=EA_Pub.SafeRequest(1,"ID",0,0,0)
	Call EA_M_XML.AppInfo("ID",ModuleID)
	
	If ModuleID > 0 Then
		TempStr=EA_M_DBO.Get_Module_Info(ModuleID)

		Call EA_M_XML.AppInfo("Title",TempStr(1,0))
		Call EA_M_XML.AppInfo("Desc",TempStr(2,0))
		Call EA_M_XML.AppInfo("Code",TempStr(3,0))

		Tmp = "(build-select)," & TempStr(4,0) & " " & str_Theme_ModuleHome & ",0 " & str_Theme_ModuleCss & ",1 " & str_Theme_ModuleHead & ",2 " & str_Theme_ModuleFoot & ",3 " & str_Theme_ModulePage & ",4 " & str_Theme_ModuleContent & ",5 " & str_Theme_ModuleSearch & ",6 " & str_Theme_ModuleComment & ",7"
		Call EA_M_XML.AppInfo("Typer",Tmp)

		Call EA_M_XML.AppElements("Language_Theme_ModuleEdit",str_Theme_ModuleEdit)
	Else
		Call EA_M_XML.AppElements("Language_Theme_ModuleEdit",str_Theme_ModuleAdd)

		Tmp = "(build-select),0 " & str_Theme_ModuleHome & ",0 " & str_Theme_ModuleCss & ",1 " & str_Theme_ModuleHead & ",2 " & str_Theme_ModuleFoot & ",3 " & str_Theme_ModulePage & ",4 " & str_Theme_ModuleContent & ",5 " & str_Theme_ModuleSearch & ",6 " & str_Theme_ModuleComment & ",7"
		Call EA_M_XML.AppInfo("Typer",Tmp)
	End If

	Page = EA_M_XML.make("","",0)
	Call EA_M_XML.Out(Page)
End Sub

Sub Save
	Dim Title,Desc, Code, Typer, ThemesID
	Dim ModuleID
	
	Title		= EA_Pub.SafeRequest(2,"Title",1,"",-1)
	Desc		= EA_Pub.SafeRequest(2,"Desc",1,"",-1)
	Code		= EA_Pub.SafeRequest(2,"Code",1,"",-1)
	Typer		= EA_Pub.SafeRequest(2,"Typer",0,0,0)
	ThemesID	= EA_Pub.SafeRequest(2,"ThemesID",0,0,0)
	ModuleID	= EA_Pub.SafeRequest(2,"ID",0,0,0)

	If Title = "" Then Response.Write "-1":Response.End
	
	If ModuleID<>0 Then
		Sql="Select * From [NB_Module] Where [Id]="&ModuleID
		rs.Open Sql,Conn,2,2
	Else
		rs.Open "NB_Module",Conn,2,2
		rs.AddNew
	End If

	rs("Title")=Title
	rs("Desc")=Desc
	rs("Code")=Code
	rs("Type")=Typer
	rs("ThemesID")=ThemesID
	rs.update
	Rs.Close:Set Rs=Nothing

	If Typer = 1 Then
		Dim Tmp

		If ModuleID = 0 Then
			Tmp = EA_M_DBO.Get_ModuleID(ThemesID, Title)
			If IsArray(Tmp) Then ModuleID = Tmp(0, 0)
		End If

		Dim re
		Set re=new RegExp
		re.IgnoreCase =true
		re.Global=True

		re.Pattern="<style([^>]*)>"
		Code=re.Replace(Code,"")

		Code=Replace(Code,"</style>","")
		Code=Replace(Code,"{$SystemPath$}",SystemFolder)

		Call EA_Pub.Save_HtmlFile("../themes/css/style_" & ThemesID & "-" & ModuleID & ".css",Code)
	End If
	
	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing

	Response.Write "1"
	Response.End
End Sub

Sub Del
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

		EA_M_DBO.Set_Module_Delete Tmp
	Next
	
	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing
	
	Response.Write "1"
	Response.End
End Sub 
%>