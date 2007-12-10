<!--#Include File="../conn.asp" -->
<!--#Include File="comm/inc.asp" -->
<!--#Include File="../include/md5.asp" -->
<!--#include file="../include/_cls_teamplate.asp"-->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：/Manager/Admin_Master.asp
'= 摘    要：后台-管理员管理文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2006-11-12
'====================================================================

Call EA_Manager.Chk_IsMaster

If Not EA_Manager.Chk_Power(Admin_Power,"31") Then 
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
	Dim Count,i
	Dim TopicList
	Dim ListName(6),ListValue()

	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_Master_Help",str_Master_Help)

	Call EA_M_XML.AppElements("Language_Master_Account",str_Master_Account)
	Call EA_M_XML.AppElements("Language_Master_LastLoginTime",str_Master_LastLoginTime)
	Call EA_M_XML.AppElements("Language_Master_LastLoginIp",str_Master_LastLoginIp)
	Call EA_M_XML.AppElements("Language_Master_State",str_Master_State)
	Call EA_M_XML.AppElements("Language_Master_MasterTotal",str_Master_MasterTotal)

	Call EA_M_XML.AppElements("Comm_Add_Operation",str_Comm_Add_Operation)
	Call EA_M_XML.AppElements("Comm_Del_Operation",str_Comm_Del_Operation)

	Call EA_M_XML.AppElements("Language_Comm_Bar_Operation",str_Comm_Bar_Operation)

	TopicList=EA_M_DBO.Get_Master_List()
	If IsArray(TopicList) Then 
		ListName(0) = "checkbox"
		ListName(1) = "ID"
		ListName(2) = "Account"
		ListName(3) = "LastLoginTime"
		ListName(4) = "LastLoginIp"
		ListName(5) = "State"
		ListName(6) = "action"
		ForTotal = UBound(TopicList,2)

		For i=0 To ForTotal
			ReDim Preserve ListValue(6,i)

			ListValue(0,i) = "checkbox"
			ListValue(1,i) = TopicList(0,i)
			ListValue(2,i) = TopicList(1,i)
			ListValue(3,i) = TopicList(2,i)
			ListValue(4,i) = TopicList(3,i)
			ListValue(5,i) = TopicList(4,i)
			ListValue(6,i) = "action"
		Next

		Call EA_M_XML.AppElements("MasterTotal",UBound(TopicList,2)+1)

		Page = EA_M_XML.make(ListName,ListValue,UBound(TopicList,2)+1)
	Else
		Page = EA_M_XML.make("","",0)

		Call EA_M_XML.AppElements("MasterTotal","0")
	End If

	Call EA_M_XML.Out(Page)
End Sub

Sub Add
	Dim PostId,BackName,State
	Dim Temp,ListBlock,ColumnList,Temp1,ListBlock1,RN
	Dim Userid,Power,Admin_Power
	Dim Level,i,j
	Dim ColumnName

	PostId=EA_Pub.SafeRequest(2,"ID",0,0,0)
	
	Temp=EA_M_DBO.Get_Master_Info(PostId)
	If IsArray(Temp) Then
		BackName=Temp(0,0)
		State=Temp(1,0)
	End If

	Temp=EA_M_DBO.Get_Master_Info(PostId)
	If IsArray(Temp) Then
		Power=Temp(2,0)
		Admin_Power=Temp(3,0)
		Admin_Power=","&Admin_Power&","
		Power=Power&","
	End If

	Dim Template
	Set Template=New cls_NEW_TEMPLATE

	PageContent=Template.LoadTemplate("admin_master_option.htm")

	Template.SetVariable "Language_Comm_Save_Button",str_Comm_Save_Button,PageContent
	Template.SetVariable "Language_Comm_Return_Button",str_Comm_Return_Button,PageContent
	Template.SetVariable "Language_Comm_Enabled",str_Comm_Enabled,PageContent
	Template.SetVariable "Language_Comm_Disabled",str_Comm_Disabled,PageContent
	Template.SetVariable "Language_Comm_SelectAll",str_Comm_SelectAll,PageContent

	Template.SetVariable "Language_OperationNotice",str_OperationNotice,PageContent
	Template.SetVariable "Language_Master_Help",str_Master_Help,PageContent

	Template.SetVariable "Language_Master_AccountList",str_Master_AccountList,PageContent
	Template.SetVariable "Language_Master_AddAccount",str_Master_AddAccount,PageContent

	Template.SetVariable "Language_Master_EditInfo",str_Master_EditInfo,PageContent
	Template.SetVariable "Language_Master_Account",str_Master_Account,PageContent
	Template.SetVariable "Language_Master_AddMaster_AccountInfo",str_Master_AddMaster_AccountInfo,PageContent
	Template.SetVariable "Language_Master_Password",str_Master_Password,PageContent
	Template.SetVariable "Language_Master_AddMaster_PasswordInfo",str_Master_AddMaster_PasswordInfo,PageContent
	Template.SetVariable "Language_Master_State",str_Master_State,PageContent
	Template.SetVariable "Language_Master_ColumnPowerOption",str_Master_ColumnPowerOption,PageContent
	Template.SetVariable "Language_Master_MenuPowerOption",str_Master_MenuPowerOption,PageContent

	Template.SetVariable "Language_Master_Column_Add",str_Master_Column_Add,PageContent
	Template.SetVariable "Language_Master_Column_Manager",str_Master_Column_Manager,PageContent
	Template.SetVariable "Language_Master_Column_Edit",str_Master_Column_Edit,PageContent
	Template.SetVariable "Language_Master_Column_Del",str_Master_Column_Del,PageContent

	Template.SetVariable "MasterID",PostId,PageContent
	Template.SetVariable "MasterAccount",BackName,PageContent
	Template.SetVariable "State_" & Abs(CInt(State))," checked",PageContent

	ListBlock=Template.GetBlock("list",PageContent)

	ColumnList=EA_DBO.Get_Column_List()
	If IsArray(ColumnList) Then
		ForTotal = UBound(ColumnList,2)

		For i=0 To ForTotal
			Temp=ListBlock

			ColumnName = ""

			If Len(ColumnList(2,i))>4 Then ColumnName = "├"
			Level=(Len(ColumnList(2,i))/4-1)*4
			ColumnName = ColumnName & String(Level,"-")

			Template.SetVariable "ColumnID",ColumnList(0,i),Temp
			Template.SetVariable "ColumnName",ColumnName & ColumnList(1,i),Temp

			If InStr(Admin_Power,","&ColumnList(0,i)&"1,")>0 Then Template.SetVariable "Power_" & ColumnList(0,i) & "_1","checked",Temp
			If InStr(Admin_Power,","&ColumnList(0,i)&"2,")>0 Then Template.SetVariable "Power_" & ColumnList(0,i) & "_2","checked",Temp
			If InStr(Admin_Power,","&ColumnList(0,i)&"3,")>0 Then Template.SetVariable "Power_" & ColumnList(0,i) & "_3","checked",Temp
			If InStr(Admin_Power,","&ColumnList(0,i)&"4,")>0 Then Template.SetVariable "Power_" & ColumnList(0,i) & "_4","checked",Temp

			Template.SetBlock "list",Temp,PageContent
		Next
    End If
	Template.CloseBlock "list",PageContent


	ListBlock=Template.GetBlock("list1",PageContent)
	ForTotal = Ubound(str_LeftMenu)-1

	For i=0 To ForTotal
		Temp=ListBlock

		ListBlock1=Template.GetBlock("list2",Temp)
		For j=1 To Ubound(str_LeftMenu,2)
			If IsEmpty(str_LeftMenu(i,j)) Then Exit For
			Temp1=ListBlock1
			
			RN = ""

			If j Mod 4 =0 Then RN = "<br>"

			Template.SetVariable "MenuName",str_LeftMenu(i,j) & RN,Temp1
			Template.SetVariable "Power",i&j,Temp1

			If InStr(Power,i&j&",")>0 Then Template.SetVariable "Power_" & i & j,"checked",Temp1
			
			Template.SetBlock "list2",Temp1,Temp
		Next
		Template.CloseBlock "list2",Temp

		Template.SetVariable "MenuName",str_LeftMenu(i,0),Temp

		Template.SetBlock "list1",Temp,PageContent
	Next
	Template.CloseBlock "list1",PageContent
	

	Template.BaseReplace PageContent
	Template.OutStr PageContent
End Sub

Sub Save
	Dim UserId,Account,PassWord,PSQL,State
	Dim Admin_Power,Power

	FoundErr = False

	UserId	= EA_Pub.SafeRequest(2,"ID",0,0,0)
	Account	= EA_Pub.SafeRequest(2,"account",1,"",0)
	PassWord= EA_Pub.SafeRequest(2,"password",1,"",0)
	State	= EA_Pub.SafeRequest(2,"state",0,0,0)

	Admin_Power	= EA_Pub.SafeRequest(2,"admin_power",1,"",0)
	Power		= EA_Pub.SafeRequest(2,"power",1,"",0)
	
	Admin_Power	= Replace(Admin_Power," ","")
	Admin_Power	= ","&Admin_Power&","
	Power		= Replace(Power," ","")
	
	If PassWord<>"" Then 
		PassWord=Md5(PassWord)
		PSQL=",master_password='"&PassWord&"'"
	End If

	If Account="" Or Len(Account)>50 Then FoundErr = True
	
	If UserId<>0 Then
		Sql="UpDate NB_Master Set Master_Name='"&Account&"',State="&State&""&PSQL&",Column_Setting = '"&Power&"', Setting = '"&Admin_Power&"' Where Master_Id="&UserId
		Conn.Execute(SQL)
	Else
		Sql="Select Master_Id From NB_Master Where Master_Name='"&Account&"'"
		Set rs=Conn.Execute(SQL)
		If Not rs.eof And Not rs.bof Then 
			FoundErr=True
		ElseIf PassWord="" Then 
			FoundErr=True
		Else
			SQL="INSERT INTO NB_Master ( Master_Name, Master_Password, State, Column_Setting, Setting )"
			SQL=SQL&" VALUES ('"&Account&"', '"&PassWord&"', "&State&", '"&Power&"', '"&Admin_Power&"')"
			Conn.Execute(SQL)
		End If
		If FoundErr Then Response.Write "-1":Response.End
	End If
	
	Set Rs=Nothing
	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing
	
	Response.Write UserId
	Response.End
End Sub

Sub Del
	Dim IDs
	Dim i,Tmp

	IDs = Split(Request.Form("ID"), ",")
	ForTotal = UBound(IDs)

	For i = 0 To ForTotal
		Tmp = EA_Pub.SafeRequest(5,IDs(i),0,0,0)

		SQL = "DELETE"
		SQL = SQL&" FROM NB_Master"
		SQL = SQL&" WHERE Master_Id="&Tmp&" And Master_Name<>'"&Session(sCacheName&"master_name")&"'"
		Conn.Execute(SQL)
	Next
	
	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing
	
	Response.Write "1"
	Response.End
End Sub
%>
