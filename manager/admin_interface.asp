<!--#Include File="comm/inc.asp" -->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2004 - 2008 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：/Manager/Admin_Interface.asp
'= 摘    要：后台-接口文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2006-07-27
'====================================================================

Call EA_Manager.Chk_IsMaster

If Not EA_Manager.Chk_Power(Admin_Power,"71") Then 
	ErrMsg=str_Comm_NotAccess
	Call EA_Manager.Error(1)
End If

Dim Action
Action=Request.Form("action")

Select Case LCase(Action)
Case "add"
	Call Add
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
	Dim Count,Page,i
	Dim TopicList
	Dim ListName(6),ListValue()
	Dim ForTotal
	
	Page=EA_Pub.SafeRequest(2,"nowPage",0,1,0)

	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_Interface_Help",str_Interface_Help)

	Call EA_M_XML.AppElements("Language_Interface_Title",str_Interface_Title)
	Call EA_M_XML.AppElements("Language_Interface_RemoteURL",str_Interface_RemoteURL)
	Call EA_M_XML.AppElements("Language_Interface_StructFile",str_Interface_StructFile)
	Call EA_M_XML.AppElements("Language_Interface_Type",str_Interface_Type)

	Call EA_M_XML.AppElements("Comm_Add_Operation",str_Comm_Add_Operation)
	Call EA_M_XML.AppElements("Comm_Del_Operation",str_Comm_Del_Operation)

	Call EA_M_XML.AppElements("Language_Comm_Bar_Operation",str_Comm_Bar_Operation)

	Count=EA_M_DBO.Get_Interface_Total()(0,0)
	If Count>0 Then 
		TopicList=EA_M_DBO.Get_Interface_List(Page,10)

		ListName(0) = "checkbox"
		ListName(1) = "ID"
		ListName(2) = "Title"
		ListName(3) = "RemoteURL"
		ListName(4) = "StructFile"
		ListName(5) = "Type"
		ListName(6) = "action"
		ForTotal = Ubound(TopicList,2)
	
	    For i=0 To ForTotal
			ReDim Preserve ListValue(6,i)

			ListValue(0,i) = "checkbox"
			ListValue(1,i) = TopicList(0,i)
			ListValue(2,i) = TopicList(1,i)
			ListValue(3,i) = TopicList(2,i)
			ListValue(4,i) = TopicList(3,i)
			Select Case TopicList(4,i)
			Case 1
				ListValue(5,i) = str_Interface_Type_UserRegister
			Case 2
				ListValue(5,i) = str_Interface_Type_UserChanngePassword
			Case 3
				ListValue(5,i) = str_Interface_Type_UserChanngeInfo
			Case 4
				ListValue(5,i) = str_Interface_Type_UserPostArticle
			Case 5
				ListValue(5,i) = str_Interface_Type_ManagerPostArticle
			End Select
			ListValue(6,i) = "action"
		Next

		Page = EA_M_XML.make(ListName,ListValue,Count)
	Else
		Page = EA_M_XML.make("","",0)
	End If

	Call EA_M_XML.Out(Page)
End Sub

Sub Add
	Dim PostId
	Dim Types
	Dim Temp,Tmp
	
	PostId=EA_Pub.SafeRequest(2,"ID",0,0,0)
	Call EA_M_XML.AppInfo("ID",PostId)
	
	Temp=EA_M_DBO.Get_Interface_Info(PostId)
	If IsArray(Temp) Then 
		Call EA_M_XML.AppInfo("Title",Temp(0,0))
		Call EA_M_XML.AppInfo("RemoteURL",Temp(1,0))
		Call EA_M_XML.AppInfo("StructFile",Temp(2,0))
		Call EA_M_XML.AppInfo("SKey",Temp(4,0))

		Types = Temp(3,0)
	End If

	Tmp = "(build-select)," & Types & " " & str_Comm_Select & ",0 " & str_Interface_Type_UserRegister & ",1 " & str_Interface_Type_UserChanngePassword & ",2 " & str_Interface_Type_UserChanngeInfo & ",3 " & str_Interface_Type_UserPostArticle & ",4 " & str_Interface_Type_ManagerPostArticle & ",5"
	Call EA_M_XML.AppInfo("Type",Tmp)

	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_Interface_Help",str_Interface_Help)
	Call EA_M_XML.AppElements("Language_Interface_Input",str_Interface_Input)

	Call EA_M_XML.AppElements("Language_Interface_Title",str_Interface_Title)
	Call EA_M_XML.AppElements("Language_Interface_RemoteURL",str_Interface_RemoteURL)
	Call EA_M_XML.AppElements("Language_Interface_StructFile",str_Interface_StructFile)
	Call EA_M_XML.AppElements("Language_Interface_Type",str_Interface_Type)
	Call EA_M_XML.AppElements("Language_Interface_SKey",str_Interface_SKey)

	Call EA_M_XML.AppElements("btnSubmit",str_Comm_Save_Button)
	Call EA_M_XML.AppElements("btnReturn",str_Comm_Return_Button)

	Page = EA_M_XML.make("","",0)
	Call EA_M_XML.Out(Page)
End Sub

Sub Save
	Dim Title,RemoteURL,StructFile,Types,SKey
	Dim PostId

	FoundErr = False
	
	PostId		= EA_Pub.SafeRequest(2,"ID",0,0,0)
	Title		= EA_Pub.SafeRequest(2,"Title",1,"",1)
	RemoteURL	= EA_Pub.SafeRequest(2,"RemoteURL",1,"",0)
	StructFile	= EA_Pub.SafeRequest(2,"StructFile",1,"",0)
	Types		= EA_Pub.SafeRequest(2,"Type",0,0,0)
	SKey		= EA_Pub.SafeRequest(2,"SKey",1,"",0)
	
	If Title="" Or Len(Title)>50 Then FoundErr = True
	If RemoteURL="" Or Len(RemoteURL)>150 Then FoundErr = True
	If StructFile="" Or Len(StructFile)>50 Then FoundErr = True
	If SKey="" Or Len(SKey)>50 Then FoundErr = True
	
	If FoundErr Then
		Response.Write "-1"
	Else
		If Rs.State=1 Then rs.Close
		If PostId<>0 Then
			Sql="Select * From [NB_Interface] Where [Id]="&PostId
			rs.Open Sql,Conn,2,2
		Else
			rs.Open "NB_Interface",Conn,2,2
			rs.AddNew
		End If
			rs("Title")=Title
			rs("RemoteURL")=RemoteURL
			rs("StructFile")=StructFile
			rs("Type")=Types
			rs("SKey")=SKey
			rs.update
		Rs.Close:Set Rs=Nothing
		
		Call EA_Pub.Close_Obj
		Set EA_Pub=Nothing
	
		Response.Write PostId
	End If

	Response.End
End Sub

Sub Del
	Dim IDs
	Dim i,Tmp
	Dim ForTotal

	IDs = Split(Request.Form("ID"), ",")
	ForTotal = UBound(IDs)
	
	For i = 0 To ForTotal
		Tmp = EA_Pub.SafeRequest(5,IDs(i),0,0,0)

		SQL="delete from [NB_Interface] where ID="&Tmp
		EA_M_DBO.DB_Execute(SQL)
	Next
	
	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing
	
	Response.Write "1"
	Response.End
End Sub
%>