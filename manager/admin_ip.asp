<!--#Include File="comm/inc.asp" -->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2004 - 2008 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：/Manager/Admin_Ip.asp
'= 摘    要：后台-IP管理文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2006-07-27
'====================================================================

Call EA_Manager.Chk_IsMaster

If Not EA_Manager.Chk_Power(Admin_Power,"32") Then 
	ErrMsg=str_Comm_NotAccess
	Call EA_Manager.Error(1)
End If

Dim Action
Action=Request("action")

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
	Dim ListName(5),ListValue()
	Dim ForTotal
	
	Page=EA_Pub.SafeRequest(2,"nowPage",0,1,0)

	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_IP_Help",str_IP_Help)

	Call EA_M_XML.AppElements("Language_IP_IPHead",str_IP_IPHead)
	Call EA_M_XML.AppElements("Language_IP_IPFoot",str_IP_IPFoot)
	Call EA_M_XML.AppElements("Language_IP_OverTime",str_IP_OverTime)

	Call EA_M_XML.AppElements("Comm_Add_Operation",str_Comm_Add_Operation)
	Call EA_M_XML.AppElements("Comm_Del_Operation",str_Comm_Del_Operation)

	Call EA_M_XML.AppElements("Language_Comm_Bar_Operation",str_Comm_Bar_Operation)

	Count=EA_M_DBO.Get_Ip_Total()(0,0)
	If Count>0 Then 
		TopicList=EA_M_DBO.Get_IP_List(Page,10)

		ListName(0) = "checkbox"
		ListName(1) = "ID"
		ListName(2) = "IPHead"
		ListName(3) = "IPFoot"
		ListName(4) = "OverTime"
		ListName(5) = "action"
		ForTotal = Ubound(TopicList,2)
	
	    For i=0 To ForTotal
			ReDim Preserve ListValue(5,i)

			ListValue(0,i) = "checkbox"
			ListValue(1,i) = TopicList(0,i)
			ListValue(2,i) = EA_Manager.ShowIp(TopicList(1,i))
			ListValue(3,i) = EA_Manager.ShowIp(TopicList(2,i))
			ListValue(4,i) = TopicList(3,i)
			ListValue(5,i) = "action"
		Next

		Page = EA_M_XML.make(ListName,ListValue,Count)
	Else
		Page = EA_M_XML.make("","",0)
	End If

	Call EA_M_XML.Out(Page)
End Sub

Sub Add
	Dim PostId,Head,Foot
	Dim Temp
	
	PostId=EA_Pub.SafeRequest(1,"ID",0,0,0)
	Call EA_M_XML.AppInfo("ID",PostId)
	
	Temp=EA_M_DBO.Get_IP_Info(PostId)
	If IsArray(Temp) Then 
		Head=EA_Manager.SplitIp(Temp(1,0))
		Foot=EA_Manager.SplitIp(Temp(2,0))

		Call EA_M_XML.AppInfo("Head_A",Head(0))
		Call EA_M_XML.AppInfo("Head_B",Head(1))
		Call EA_M_XML.AppInfo("Head_C",Head(2))
		Call EA_M_XML.AppInfo("Head_D",Head(3))
		Call EA_M_XML.AppInfo("Foot_A",Foot(0))
		Call EA_M_XML.AppInfo("Foot_B",Foot(1))
		Call EA_M_XML.AppInfo("Foot_C",Foot(2))
		Call EA_M_XML.AppInfo("Foot_D",Foot(3))
		Call EA_M_XML.AppInfo("OverTime",Temp(3,0))
	End If

	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_IP_Help",str_IP_Help)
	Call EA_M_XML.AppElements("Language_IP_InputIp",str_IP_InputIp)

	Call EA_M_XML.AppElements("Language_IP_IPHead",str_IP_IPHead)
	Call EA_M_XML.AppElements("Language_IP_IPFoot",str_IP_IPFoot)
	Call EA_M_XML.AppElements("Language_IP_OverTime",str_IP_OverTime)

	Call EA_M_XML.AppElements("btnSubmit",str_Comm_Save_Button)
	Call EA_M_XML.AppElements("btnReturn",str_Comm_Return_Button)

	Page = EA_M_XML.make("","",0)
	Call EA_M_XML.Out(Page)
End Sub

Sub Save
	Dim Ip_Head,Ip_Foot,OverTime
	Dim PostId
	
	PostId	= EA_Pub.SafeRequest(2,"ID",0,0,0)
	Ip_Head	= EA_Pub.SafeRequest(2,"Head_A",0,0,0) & "." & EA_Pub.SafeRequest(2,"Head_B",0,0,0) & "." & EA_Pub.SafeRequest(2,"Head_C",0,0,0) & "." & EA_Pub.SafeRequest(2,"Head_D",0,0,0)
	Ip_Foot	= EA_Pub.SafeRequest(2,"Foot_A",0,0,0) & "." & EA_Pub.SafeRequest(2,"Foot_B",0,0,0) & "." & EA_Pub.SafeRequest(2,"Foot_C",0,0,0) & "." & EA_Pub.SafeRequest(2,"Foot_D",0,0,0)
	OverTime= EA_Pub.SafeRequest(2,"OverTime",2,Date()+3,0)

	Ip_Head	= EA_Pub.FormatIp(Ip_Head)
	Ip_Foot	= EA_Pub.FormatIp(Ip_Foot)

	If FoundErr Then
		Response.Write "-1"
	Else
		If Rs.State=1 Then rs.Close
		If PostId<>0 Then
			Sql="Select * From [NB_IP] Where [Id]="&PostId
			rs.Open Sql,Conn,2,2
		Else
			rs.Open "NB_IP",Conn,2,2
			rs.AddNew
		End If
			rs("Head_Ip")=Ip_Head
			rs("Foot_Ip")=Ip_Foot
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
	Dim ForTotal

	IDs = Split(Request.Form("ID"), ",")
	ForTotal = UBound(IDs)
	
	For i = 0 To ForTotal
		Tmp = EA_Pub.SafeRequest(5,IDs(i),0,0,0)

		EA_M_DBO.Set_IP_Delete Tmp
	Next
	
	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing
	
	Response.Write "1"
	Response.End
End Sub
%>