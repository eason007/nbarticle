
<!--#Include File="comm/inc.asp" -->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2004 - 2008 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：/Manager/Admin_Theme.asp
'= 摘    要：后台-风格管理文件
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
	Dim ListName(5),ListValue()
	Dim ForTotal

	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_Theme_Info",str_Theme_Info)
	Call EA_M_XML.AppElements("Language_Theme_Edit",str_Theme_Add)

	Call EA_M_XML.AppElements("Language_Theme_Name",str_Theme_Name)
	Call EA_M_XML.AppElements("Language_Theme_Default",str_Theme_Default)
	Call EA_M_XML.AppElements("Language_Theme_Name1",str_Theme_Name)
	Call EA_M_XML.AppElements("Language_Theme_Default1",str_Theme_Default)

	Call EA_M_XML.AppElements("Comm_Add_Operation",str_Comm_Add_Operation)
	Call EA_M_XML.AppElements("Comm_Del_Operation",str_Comm_Del_Operation)
	Call EA_M_XML.AppElements("btnSubmit",str_Comm_Save_Button)
	Call EA_M_XML.AppElements("btnReset",str_Comm_Reset_Button)

	Call EA_M_XML.AppElements("Language_Comm_Yes",str_Comm_Yes)
	Call EA_M_XML.AppElements("Language_Comm_No",str_Comm_No)

	Call EA_M_XML.AppElements("Language_Comm_Bar_Operation",str_Comm_Bar_Operation)

	TopicList=EA_M_DBO.Get_Theme_List()

	ListName(0) = "checkbox"
	ListName(1) = "ID"
	ListName(2) = "IDs"
	ListName(3) = "Name"
	ListName(4) = "IsDefault"
	ListName(5) = "action"
	ForTotal = Ubound(TopicList,2)
	
	For i=0 To ForTotal
		ReDim Preserve ListValue(5,i)

		ListValue(0,i) = "checkbox"
		ListValue(1,i) = TopicList(0,i)
		ListValue(2,i) = TopicList(0,i)
		ListValue(3,i) = TopicList(1,i)
		If TopicList(2,i) Then
			ListValue(4,i) = "<strong>√</strong>"
		Else
			ListValue(4,i) = "<font color=""red""><strong>×</strong></font>"
		End If
		ListValue(5,i) = "action"
	Next

	Page = EA_M_XML.make(ListName,ListValue,Ubound(TopicList,2)+1)

	Call EA_M_XML.Out(Page)
End Sub

Sub Edit
	Dim TempStr,ThemeId
	Dim i
	
	ThemeId=EA_Pub.SafeRequest(2,"ID",0,0,0)
	Call EA_M_XML.AppInfo("ID",ThemeId)
	Call EA_M_XML.AppElements("Language_Theme_Edit",str_Theme_Edit)
	
	If ThemeId > 0 Then
		TempStr=EA_M_DBO.Get_Theme_Info(ThemeId)

		Call EA_M_XML.AppInfo("title",TempStr(1,0))
		Call EA_M_XML.AppInfo("isdefault",TempStr(2,0))
	End If

	Page = EA_M_XML.make("","",0)
	Call EA_M_XML.Out(Page)
End Sub

Sub Save
	Dim ThemeID, Title, IsDefault

	FoundErr = False
	
	ThemeID		= EA_Pub.SafeRequest(2,"ID",0,0,0)
	Title		= EA_Pub.SafeRequest(2,"Title",1,"",0)
	IsDefault	= EA_Pub.SafeRequest(2,"isdefault",0,0,0)
	
	If Title="" Or Len(Title)>50 Then FoundErr = True
	
	If FoundErr Then
		Response.Write "-1"
	Else
		If ThemeID<>0 Then
			Sql="Select * From [NB_Themes] Where [Id]="&ThemeID
			rs.Open Sql,Conn,2,2
		Else
			rs.Open "NB_Themes",Conn,2,2
			rs.AddNew
		End If

		rs("Title")=Title
		rs("IsDefault")=IsDefault
		rs.update
		Rs.Close:Set Rs=Nothing
		
		Call EA_Pub.Close_Obj
		Set EA_Pub=Nothing
	
		Response.Write ThemeID
	End If

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