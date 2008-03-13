<!--#Include File="comm/inc.asp" -->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：/Manager/Admin_AdSense.asp
'= 摘    要：后台-广告管理文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2006-07-27
'====================================================================

Call EA_Manager.Chk_IsMaster

If Not EA_Manager.Chk_Power(Admin_Power,"06") Then 
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
	Dim ListName(4),ListValue()
	Dim ForTotal
	
	Page=EA_Pub.SafeRequest(2,"nowPage",0,1,0)

	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_AdSense_Help",str_AdSense_Help)

	Call EA_M_XML.AppElements("Language_AdSense_Title",str_AdSense_Title)

	Call EA_M_XML.AppElements("Comm_Add_Operation",str_Comm_Add_Operation)
	Call EA_M_XML.AppElements("Comm_Del_Operation",str_Comm_Del_Operation)

	Call EA_M_XML.AppElements("Language_Comm_Bar_Operation",str_Comm_Bar_Operation)

	Count=EA_M_DBO.Get_AdSense_Total()(0,0)
	If Count>0 Then 
		TopicList=EA_M_DBO.Get_AdSense_List(Page,10)

		ListName(0) = "checkbox"
		ListName(1) = "ID"
		ListName(2) = "I_ID"
		ListName(3) = "Title"
		ListName(4) = "action"
		ForTotal = Ubound(TopicList,2)
	
	    For i=0 To ForTotal
			ReDim Preserve ListValue(4,i)

			ListValue(0,i) = "checkbox"
			ListValue(1,i) = TopicList(0,i)
			ListValue(2,i) = TopicList(0,i)
			ListValue(3,i) = TopicList(1,i)
			ListValue(4,i) = "action"
		Next

		Page = EA_M_XML.make(ListName,ListValue,Count)
	Else
		Page = EA_M_XML.make("","",0)
	End If

	Call EA_M_XML.Out(Page)
End Sub

Sub Add
	Dim PostId,Temp

	PostId=EA_Pub.SafeRequest(2,"ID",0,0,0)
	Call EA_M_XML.AppInfo("ID",PostId)
	
	Temp=EA_DBO.Get_AdSense_Info(PostId)
	If IsArray(Temp) Then 
		Call EA_M_XML.AppInfo("Title",Temp(0,0))
		Call EA_M_XML.AppInfo("Content",Temp(1,0))
	End If

	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_AdSense_Help",str_AdSense_Help)
	Call EA_M_XML.AppElements("Language_AdSense_Input_Template",str_AdSense_Input_Template)

	Call EA_M_XML.AppElements("Language_AdSense_Title",str_AdSense_Title)
	Call EA_M_XML.AppElements("Language_AdSense_Content",str_AdSense_Content)

	Call EA_M_XML.AppElements("btnSubmit",str_Comm_Save_Button)
	Call EA_M_XML.AppElements("btnReturn",str_Comm_Return_Button)

	Page = EA_M_XML.make("","",0)
	Call EA_M_XML.Out(Page)
End Sub

Sub Save
	Dim Title,Content
	Dim PostId

	FoundErr = False
	
	PostId	= EA_Pub.SafeRequest(2,"ID",0,0,0)
	Title	= EA_Pub.SafeRequest(2,"Title",1,"",0)
	Content	= EA_Pub.SafeRequest(2,"Content",1,"",-1)
	
	If Title="" Or Len(Title)>50 Then FoundErr = True
	
	If FoundErr Then
		Response.Write "-1"
	Else
		If PostId<>0 Then
			Sql="Select * From [NB_AdSense] Where [Id]="&PostId
			rs.Open Sql,Conn,2,2
		Else
			rs.Open "NB_AdSense",Conn,2,2
			rs.AddNew
		End If
			rs("Title")=Title
			rs("Content")=Content
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

		EA_M_DBO.Set_AdSense_Del Tmp
	Next
	
	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing
	
	Response.Write "1"
	Response.End
End Sub
%>