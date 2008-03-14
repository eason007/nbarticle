<!--#Include File="comm/inc.asp" -->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2004 - 2008 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：/Manager/Admin_Placard.asp
'= 摘    要：后台-站点公告文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2006-07-27
'====================================================================

Call EA_Manager.Chk_IsMaster

If Not EA_Manager.Chk_Power(Admin_Power,"03") Then 
	ErrMsg=str_Comm_NotAccess
	Call EA_Manager.Error(1)
End If

Dim Action
Dim ForTotal

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
	
	Page=EA_Pub.SafeRequest(2,"nowPage",0,1,0)

	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_Placard_Help",str_Placard_Help)

	Call EA_M_XML.AppElements("Language_Placard_Title",str_Placard_Title)
	Call EA_M_XML.AppElements("Language_Placard_AddTime",str_Placard_AddTime)
	Call EA_M_XML.AppElements("Language_Placard_OverTime",str_Placard_OverTime)

	Call EA_M_XML.AppElements("Comm_Add_Operation",str_Comm_Add_Operation)
	Call EA_M_XML.AppElements("Comm_Del_Operation",str_Comm_Del_Operation)

	Call EA_M_XML.AppElements("Language_Comm_Bar_Operation",str_Comm_Bar_Operation)

	Count=EA_M_DBO.Get_Placard_Total()(0,0)
	If Count>0 Then 
		TopicList=EA_M_DBO.Get_Placard_List(Page,10)

		ListName(0) = "checkbox"
		ListName(1) = "ID"
		ListName(2) = "P_ID"
		ListName(3) = "Title"
		ListName(4) = "AddTime"
		ListName(5) = "OverTime"
		ListName(6) = "action"
		ForTotal = Ubound(TopicList,2)

	    For i=0 To ForTotal
			ReDim Preserve ListValue(6,i)

			ListValue(0,i) = "checkbox"
			ListValue(1,i) = TopicList(0,i)
			ListValue(2,i) = TopicList(0,i)
			ListValue(3,i) = TopicList(1,i)
			ListValue(4,i) = TopicList(2,i)
			ListValue(5,i) = TopicList(3,i)
			ListValue(6,i) = "action"
		Next

		Page = EA_M_XML.make(ListName,ListValue,Count)
	Else
		Page = EA_M_XML.make("","",0)
	End If

	Call EA_M_XML.Out(Page)
End Sub

Sub Add
	Dim PostId,Info
	Dim i
	
	PostId=EA_Pub.SafeRequest(2,"ID",0,0,0)
	Call EA_M_XML.AppInfo("ID",PostId)
	
	Info=EA_DBO.Get_PlacardInfo(PostId)
	If IsArray(Info) Then 
		Call EA_M_XML.AppInfo("Title",Info(0,0))
		Call EA_M_XML.AppInfo("Content",Info(1,0))
		Call EA_M_XML.AppInfo("OverTime",FormatDateTime(Info(2,0),2))
	End If

	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_Placard_Help",str_Placard_Help)
	Call EA_M_XML.AppElements("Language_Placard_Input_Placard",str_Placard_Input_Placard)

	Call EA_M_XML.AppElements("Language_Placard_Title",str_Placard_Title)
	Call EA_M_XML.AppElements("Language_Placard_Content",str_Placard_Content)
	Call EA_M_XML.AppElements("Language_Placard_OverTime",str_Placard_OverTime)

	Call EA_M_XML.AppElements("btnSubmit",str_Comm_Save_Button)
	Call EA_M_XML.AppElements("btnReturn",str_Comm_Return_Button)

	Page = EA_M_XML.make("","",0)
	Call EA_M_XML.Out(Page)
End Sub

Sub Save
	Dim Title,Content,OverTime
	Dim PostId

	FoundErr = False
	
	PostId	= EA_Pub.SafeRequest(2,"ID",0,0,0)
	Title	= EA_Pub.SafeRequest(2,"Title",1,"",1)
	Content	= EA_Pub.SafeRequest(2,"Content",1,"",0)
	OverTime= EA_Pub.SafeRequest(2,"OverTime",2,Now()+7,0)
	
	If Title="" Or Len(Title)>150 Then FoundErr = True
	If Content="" Or Len(Content)>250 Then FoundErr = True
	
	If FoundErr Then
		Response.Write "-1"
	Else
		If Rs.State=1 Then rs.Close
		If PostId<>0 Then
			Sql="Select * From [NB_Placard] Where [Id]="&PostId
			rs.Open Sql,Conn,2,2
		Else
			rs.Open "NB_Placard",Conn,2,2
			rs.AddNew
		End If
			rs("Title")=Title
			rs("Content")=Content
			rs("OverTime")=OverTime
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

	IDs = Split(Request.Form("ID"), ",")
	ForTotal = UBound(IDs)

	For i = 0 To ForTotal
		Tmp = EA_Pub.SafeRequest(5,IDs(i),0,0,0)

		SQL="delete from [NB_Placard] where ID="&Tmp
		EA_M_DBO.DB_Execute(SQL)
	Next
	
	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing
	
	Response.Write "1"
	Response.End
End Sub
%>