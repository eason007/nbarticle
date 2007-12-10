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
'= 文件名称：/Manager/Admin_UserGroup.asp
'= 摘    要：后台-会员组管理文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2007-01-02
'====================================================================

Call EA_Manager.Chk_IsMaster

If Not EA_Manager.Chk_Power(Admin_Power,"21") Then 
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
	Call del
Case Else
	Call Main
End Select
Call EA_Pub.Close_Obj
Set EA_Pub=Nothing

Sub Main
	Dim Count,i
	Dim TopicList
	Dim ListName(4),ListValue()
	Dim ForTotal

	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_Group_Help",str_Group_Help)

	Call EA_M_XML.AppElements("Language_Group_Name",str_Group_Name)
	Call EA_M_XML.AppElements("Language_Group_UserTotal",str_Group_UserTotal)

	Call EA_M_XML.AppElements("Comm_Add_Operation",str_Comm_Add_Operation)
	Call EA_M_XML.AppElements("Comm_Del_Operation",str_Comm_Del_Operation)

	Call EA_M_XML.AppElements("Language_Comm_Bar_Operation",str_Comm_Bar_Operation)

	TopicList=EA_M_DBO.Get_Group_List()
	If IsArray(TopicList) Then 
		ListName(0) = "checkbox"
		ListName(1) = "ID"
		ListName(2) = "Name"
		ListName(3) = "UserTotal"
		ListName(4) = "action"
		ForTotal = UBound(TopicList,2)

		For i=0 To ForTotal
			ReDim Preserve ListValue(4,i)

			ListValue(0,i) = "checkbox"
			ListValue(1,i) = TopicList(0,i)
			ListValue(2,i) = TopicList(1,i)
			ListValue(3,i) = TopicList(2,i)
			ListValue(4,i) = "action"
		Next

		Page = EA_M_XML.make(ListName,ListValue,Count)
	Else
		Page = EA_M_XML.make("","",0)
	End If

	Call EA_M_XML.Out(Page)
End Sub

Sub Add
	Dim PostId,Temp,Setting,i

	PostId=EA_Pub.SafeRequest(2,"ID",0,0,0)
	Call EA_M_XML.AppInfo("ID",PostId)
	
	Temp=EA_DBO.Get_Group_Setting(PostId)
	If IsArray(Temp) Then 
		Call EA_M_XML.AppInfo("GroupName",Temp(0,0))
		Call EA_M_XML.AppInfo("IsLogin",Abs(CInt(Temp(1,0))))

		Setting=Split(Temp(2,0),",")
	End If
	
	If Not IsArray(Setting) Then Setting=Split("0,0,8|23,1,10,10,1,0,1,0,50,10",",")
	If UBound(Setting)<11 Then Setting=Split("0,0,8|23,1,10,10,1,0,1,0,50,10",",")

	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_Group_Help",str_Group_Help)

	Call EA_M_XML.AppElements("Language_Group_Login_Option",str_Group_Login_Option)
	Call EA_M_XML.AppElements("Language_Group_Power_Option",str_Group_Power_Option)
	Call EA_M_XML.AppElements("Language_Group_Name",str_Group_Name)
	Call EA_M_XML.AppElements("Language_Group_IsLogin",str_Group_IsLogin)
	Call EA_M_XML.AppElements("Language_Group_Timer",str_Group_Timer)
	Call EA_M_XML.AppElements("Language_Group_Timer_Option",str_Group_Timer_Option)
	Call EA_M_XML.AppElements("Language_Group_Timer_Option_Help",str_Group_Timer_Option_Help)
	Call EA_M_XML.AppElements("Language_Group_Power",str_Group_Power)
	Call EA_M_XML.AppElements("Language_Group_ViewHide",str_Group_ViewHide)
	Call EA_M_XML.AppElements("Language_Group_PostReviewForRegLater",str_Group_PostReviewForRegLater)
	Call EA_M_XML.AppElements("Language_Group_PostReviewForRegLater_Help",str_Group_PostReviewForRegLater_Help)
	Call EA_M_XML.AppElements("Language_Group_PostVotedForRegLater",str_Group_PostVotedForRegLater)
	Call EA_M_XML.AppElements("Language_Group_PostVotedForRegLater_Help",str_Group_PostVotedForRegLater_Help)
	Call EA_M_XML.AppElements("Language_Group_IsReview",str_Group_IsReview)
	Call EA_M_XML.AppElements("Language_Group_ReviewForManager",str_Group_ReviewForManager)
	Call EA_M_XML.AppElements("Language_Group_IsPostArticle",str_Group_IsPostArticle)
	Call EA_M_XML.AppElements("Language_Group_PostForManager",str_Group_PostForManager)
	Call EA_M_XML.AppElements("Language_Group_DayMaxPost",str_Group_DayMaxPost)
	Call EA_M_XML.AppElements("Language_Group_FavMax",str_Group_FavMax)
	Call EA_M_XML.AppElements("Language_Group_FavMax_Help",str_Group_FavMax_Help)

	For i= 1 To 7
		Call EA_M_XML.AppElements("Language_Comm_Yes" & i,str_Comm_Yes)
		Call EA_M_XML.AppElements("Language_Comm_No" & i,str_Comm_No)
	Next

	For i= 1 To 2
		Call EA_M_XML.AppElements("btnSubmit" & i,str_Comm_Save_Button)
		Call EA_M_XML.AppElements("btnReturn" & i,str_Comm_Return_Button)
	Next

	Call EA_M_XML.AppInfo("Power",Setting(0))
	Call EA_M_XML.AppInfo("IsHide",Setting(1))
	Call EA_M_XML.AppInfo("OpenTime",Setting(2))
	Call EA_M_XML.AppInfo("IsClose",Setting(3))
	Call EA_M_XML.AppInfo("VotedSplit",Setting(4))
	Call EA_M_XML.AppInfo("ReviewSplit",Setting(5))
	Call EA_M_XML.AppInfo("IsReview",Setting(6))
	Call EA_M_XML.AppInfo("ReviewNeedManager",Setting(7))
	Call EA_M_XML.AppInfo("IsPost",Setting(8))
	Call EA_M_XML.AppInfo("PostNeedManager",Setting(9))
	Call EA_M_XML.AppInfo("FavMax",Setting(10))
	Call EA_M_XML.AppInfo("PostMax",Setting(11))

	Page = EA_M_XML.make("","",0)
	Call EA_M_XML.Out(Page)
End Sub

Sub Save
	Dim IsExist
	Dim PostId
	Dim Name,Setting,IsLogin

	FoundErr = False
	
	PostId	= EA_Pub.SafeRequest(2,"ID",0,0,0)
	Name	= EA_Pub.SafeRequest(2,"GroupName",1,"",0)
	IsLogin	= EA_Pub.SafeRequest(2,"IsLogin",0,0,0)
	
	Setting = EA_Pub.SafeRequest(2,"Power",0,0,0)
	Setting = Setting&","&EA_Pub.SafeRequest(2,"IsHide",0,0,0)
	Setting = Setting&","&EA_Pub.SafeRequest(2,"OpenTime",1,"8|23",0)
	Setting = Setting&","&EA_Pub.SafeRequest(2,"IsClose",0,0,0)
	Setting = Setting&","&EA_Pub.SafeRequest(2,"VotedSplit",0,0,0)
	Setting = Setting&","&EA_Pub.SafeRequest(2,"ReviewSplit",0,0,0)
	Setting = Setting&","&EA_Pub.SafeRequest(2,"IsReview",0,0,0)
	Setting = Setting&","&EA_Pub.SafeRequest(2,"ReviewNeedManager",0,0,0)
	Setting = Setting&","&EA_Pub.SafeRequest(2,"IsPost",0,0,0)
	Setting = Setting&","&EA_Pub.SafeRequest(2,"PostNeedManager",0,0,0)
	Setting = Setting&","&EA_Pub.SafeRequest(2,"FavMax",0,0,0)
	Setting = Setting&","&EA_Pub.SafeRequest(2,"PostMax",0,0,0)
	
	If Len(Name)>50 Or Len(Name)=0 Then FoundErr = True
	
	If FoundErr Then
		Response.Write "-1"
	Else
		If Rs.State=1 Then rs.Close
		If PostId<>0 Then
			Sql="Select * From [NB_UserGroup] Where [Id]="&PostId
			rs.Open Sql,Conn,2,2
		Else
			rs.Open "NB_UserGroup",Conn,2,2
			rs.AddNew
		End If
			rs("GroupName")=Name
			rs("IsLogin")=IsLogin
			rs("Setting")=Setting
			rs.update
		Rs.Close:Set Rs=Nothing
		
		Dim EA_Ini
		Set EA_Ini=New cls_Ini
		EA_Ini.OpenFile	= EA_Pub.sIniFilePath

		Call EA_Ini.WriteNode("GroupSetting","Group_"&PostId,Name&","&IsLogin&","&Setting)
		EA_Ini.Save
		EA_Ini.Close
		Set EA_Ini=Nothing
		
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

		If Tmp > 4 Then EA_M_DBO.Set_Group_Delete Tmp
	Next
	
	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing
	
	Response.Write "1"
	Response.End
End Sub
%>