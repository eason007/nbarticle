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
'= 文件名称：/Manager/Admin_InsideLink.asp
'= 摘    要：后台-站内连接管理文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2006-07-27
'====================================================================

Call EA_Manager.Chk_IsMaster

If Not EA_Manager.Chk_Power(Admin_Power,"05") Then 
	ErrMsg=str_Comm_NotAccess
	Call EA_Manager.Error(1)
End If

Dim Action
Dim ForTotal
Action=Request.Form("action")

Select Case LCase(Action)
Case "save"
	Call Save
Case "del"
	Call del
Case "add"
	Call Add
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
	Call EA_M_XML.AppElements("Language_InsideLink_Help",str_InsideLink_Help)

	Call EA_M_XML.AppElements("Language_InsideLink_LinkWord",str_InsideLink_LinkWord)
	Call EA_M_XML.AppElements("Language_InsideLink_LinkURL",str_InsideLink_LinkURL)
	Call EA_M_XML.AppElements("Language_InsideLink_Location",str_InsideLink_Location)

	Call EA_M_XML.AppElements("Comm_Add_Operation",str_Comm_Add_Operation)
	Call EA_M_XML.AppElements("Comm_Del_Operation",str_Comm_Del_Operation)

	Call EA_M_XML.AppElements("Language_Comm_Bar_Operation",str_Comm_Bar_Operation)

	Count=EA_M_DBO.Get_InsideLink_Total()(0,0)
	If Count>0 Then 
		TopicList=EA_M_DBO.Get_InsideLink_List(Page,10)

		ListName(0) = "checkbox"
		ListName(1) = "ID"
		ListName(2) = "I_ID"
		ListName(3) = "LinkWord"
		ListName(4) = "LinkURL"
		ListName(5) = "Location"
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
	Dim ColumnId
	Dim Level,Temp,Tmp,i
	
	PostId=EA_Pub.SafeRequest(2,"ID",0,0,0)
	Call EA_M_XML.AppInfo("ID",PostId)

	Info=EA_M_DBO.Get_InsideLink_Info(PostId)
	If IsArray(Info) Then 
		Call EA_M_XML.AppInfo("Word",Info(0,0))
		Call EA_M_XML.AppInfo("Link",Info(1,0))
	End If

	Temp=EA_DBO.Get_Column_List()
	Tmp = "(build-select)," & ColumnId & " " & str_InsideLink_All & ",0"
	If IsArray(Temp) Then
		ForTotal = UBound(Temp,2)

		For i=0 To ForTotal
			Level=(Len(Temp(2,i))/4-1)*3

			Tmp = Tmp & " ├" & String(Level,"-") & Replace(Temp(1,i)," ","") & "," & Temp(0,i)
		Next
	End If
	Call EA_M_XML.AppInfo("Column",Tmp)

	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_InsideLink_Help",str_InsideLink_Help)
	Call EA_M_XML.AppElements("Language_InsideLink_Input_Info",str_InsideLink_Input_Info)

	Call EA_M_XML.AppElements("Language_InsideLink_LinkWord",str_InsideLink_LinkWord)
	Call EA_M_XML.AppElements("Language_InsideLink_LinkURL",str_InsideLink_LinkURL)
	Call EA_M_XML.AppElements("Language_InsideLink_Location",str_InsideLink_Location)

	Call EA_M_XML.AppElements("btnSubmit",str_Comm_Save_Button)
	Call EA_M_XML.AppElements("btnReturn",str_Comm_Return_Button)

	Page = EA_M_XML.make("","",0)
	Call EA_M_XML.Out(Page)
End Sub

Sub Save
	Dim Word,LinkStr,ColumnId
	Dim PostId

	PostId	= EA_Pub.SafeRequest(2,"ID",0,0,0)
	Word	= EA_Pub.SafeRequest(2,"Word",1,"",1)
	LinkStr	= EA_Pub.SafeRequest(2,"Link",1,"",1)
	ColumnId= EA_Pub.SafeRequest(2,"Column",0,0,0)

	If Rs.State=1 Then rs.Close
	If PostId<>0 Then
		Sql="Select * From [NB_Link] Where [Id]="&PostId
		rs.Open Sql,Conn,2,2
	Else
		rs.Open "NB_Link",Conn,2,2
		rs.AddNew
	End If
		rs("Word")=Word
		rs("Link")=LinkStr
		rs("ColumnId")=ColumnId
		rs.update
	Rs.Close:Set Rs=Nothing
	
	Set Rs=Nothing
	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing
	
	Response.Write PostId

	Response.End
End Sub

Sub Del
	Dim IDs
	Dim i,Tmp

	IDs = Split(Request.Form("ID"), ",")
	ForTotal = UBound(IDs)

	For i = 0 To ForTotal
		Tmp = EA_Pub.SafeRequest(5,IDs(i),0,0,0)

		EA_M_DBO.Set_InsideLink_Delete Tmp
	Next
	
	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing
	
	Response.Write "1"
	Response.End
End Sub
%>