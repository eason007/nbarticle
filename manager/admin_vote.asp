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
'= 文件名称：/Manager/Admin_Vote.asp
'= 摘    要：后台-投票管理文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2006-07-27
'====================================================================

Call EA_Manager.Chk_IsMaster

If Not EA_Manager.Chk_Power(Admin_Power,"04") Then 
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
Case "lock"
	Call Locker
Case "unlock"
	Call UnLocker
Case Else
	Call Main
End Select
Call EA_Pub.Close_Obj
Set EA_Pub=Nothing

Sub Main
	Dim Count,Page,i
	Dim TopicList
	Dim ListName(7),ListValue()
	
	Page=EA_Pub.SafeRequest(2,"nowPage",0,1,0)

	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_Vote_Help",str_Vote_Help)

	Call EA_M_XML.AppElements("Language_Vote_Title",str_Vote_Title)
	Call EA_M_XML.AppElements("Language_Vote_VotedTotal",str_Vote_VotedTotal)
	Call EA_M_XML.AppElements("Language_Vote_Type",str_Vote_Type)
	Call EA_M_XML.AppElements("Language_Vote_State",str_Vote_State)

	Call EA_M_XML.AppElements("Comm_Add_Operation",str_Comm_Add_Operation)
	Call EA_M_XML.AppElements("Comm_Del_Operation",str_Comm_Del_Operation)

	Call EA_M_XML.AppElements("Language_Comm_Bar_Operation",str_Comm_Bar_Operation)

	Call EA_M_XML.AppElements("Comm_Disabled",str_Comm_Disabled)
	Call EA_M_XML.AppElements("Comm_Enabled",str_Comm_Enabled)

	Count=EA_M_DBO.Get_Vote_Stat()(0,0)
	If Count>0 Then 
		TopicList=EA_M_DBO.Get_Vote_List(Page,10)

		ListName(0) = "checkbox"
		ListName(1) = "ID"
		ListName(2) = "V_ID"
		ListName(3) = "Title"
		ListName(4) = "VotedTotal"
		ListName(5) = "Type"
		ListName(6) = "State"
		ListName(7) = "action"
		ForTotal = Ubound(TopicList,2)
	
	    For i=0 To ForTotal
			ReDim Preserve ListValue(7,i)

			ListValue(0,i) = "checkbox"
			ListValue(1,i) = TopicList(0,i)
			ListValue(2,i) = TopicList(0,i)
			ListValue(3,i) = TopicList(1,i)
			ListValue(4,i) = TopicList(2,i)
			ListValue(5,i) = TopicList(3,i)
			If TopicList(5,i) Then
				ListValue(6,i) = str_Comm_Disabled
			Else
				ListValue(6,i) = str_Comm_Enabled
			End If
			ListValue(7,i) = "action"
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
	
	Temp=EA_DBO.Get_Vote_Info(PostId)
	If IsArray(Temp) Then 
		Call EA_M_XML.AppInfo("Title",Temp(1,0))
		Call EA_M_XML.AppInfo("VoteText",Replace(Temp(2,0),"|",chr(13)&chr(10)))
		Call EA_M_XML.AppInfo("Typer",Abs(CInt(Temp(4,0))))
	End If

	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_Vote_Help",str_Vote_Help)
	Call EA_M_XML.AppElements("Language_Vote_Input_Vote",str_Vote_Input_Vote)

	Call EA_M_XML.AppElements("Language_Vote_Title",str_Vote_Title)
	Call EA_M_XML.AppElements("Language_Vote_Content",str_Vote_Content)
	Call EA_M_XML.AppElements("Language_Vote_Type",str_Vote_Type)
	Call EA_M_XML.AppElements("Language_Vote_Type_Radio",str_Vote_Type_Radio)
	Call EA_M_XML.AppElements("Language_Vote_Type_Check",str_Vote_Type_Check)

	Call EA_M_XML.AppElements("btnSubmit",str_Comm_Save_Button)
	Call EA_M_XML.AppElements("btnReturn",str_Comm_Return_Button)

	Page = EA_M_XML.make("","",0)
	Call EA_M_XML.Out(Page)
End Sub

Sub Save
	Dim Title,VoteType,VoteText,VoteNum
	Dim j,i
	Dim TempStr
	Dim PostId

	FoundErr = False
	
	PostId	= EA_Pub.SafeRequest(2,"ID",0,0,0)
	Title	= EA_Pub.SafeRequest(2,"Title",1,"",1)
	VoteType= EA_Pub.SafeRequest(2,"Typer",0,0,0)
	VoteText= EA_Pub.SafeRequest(2,"VoteText",1,"",0)
	
	If VoteType="" Or VoteText="" Then FoundErr = True

	If FoundErr Then
		Response.Write "-1"
	Else
		VoteText=Split(VoteText,Chr(13)&Chr(10))
		j=0
		ForTotal = UBound(VoteText)
		
		For i = 0 To ForTotal
			If Not (VoteText(i)="" Or VoteText(i)=" ") Then
				TempStr=TempStr&""&VoteText(i)&"|"
				j=j+1
			End If
		Next

		For i = 1 To j
			VoteNum=VoteNum&"0|"
		Next

		VoteNum=Left(VoteNum,Len(VoteNum)-1)
		VoteText=Left(TempStr,Len(TempStr)-1)
		
		If Rs.State=1 Then rs.Close
		If PostId<>0 Then
			Sql="Select * From [NB_Vote] Where [Id]="&PostId
			rs.Open Sql,Conn,2,2

			TempStr = Split(VoteText, "|")
			j = UBound(TempStr)

			TempStr = Split(rs("VoteNum"), "|")
			
			If j > UBound(TempStr) Then
				ForTotal = UBound(TempStr)

				For i = ForTotal To j
					rs("VoteNum") = rs("VoteNum") & "|0"
				Next
			End If
		Else
			rs.Open "NB_Vote",Conn,2,2
			rs.AddNew
			rs("VoteNum")=VoteNum
		End If
			rs("Title")=Title
			rs("VoteText")=VoteText
			rs("Type")=VoteType
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

		EA_M_DBO.Set_Vote_Delete Tmp
	Next
	
	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing
	
	Response.Write "1"
	Response.End
End Sub

Sub Locker
	Dim IDs
	Dim i,Tmp

	IDs = Split(Request.Form("ID"), ",")
	ForTotal = UBound(IDs)

	For i = 0 To ForTotal
		Tmp = EA_Pub.SafeRequest(5,IDs(i),0,0,0)

		EA_M_DBO.Set_Vote_State 1,Tmp
	Next
	
	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing
	
	Response.Write "1"
	Response.End
End Sub

Sub UnLocker
	Dim IDs
	Dim i,Tmp

	IDs = Split(Request.Form("ID"), ",")
	ForTotal = UBound(IDs)

	For i = 0 To ForTotal
		Tmp = EA_Pub.SafeRequest(5,IDs(i),0,0,0)

		EA_M_DBO.Set_Vote_State 0,Tmp
	Next
	
	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing
	
	Response.Write "1"
	Response.End
End Sub
%>