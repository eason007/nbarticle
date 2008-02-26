<!--#Include File="../conn.asp" -->
<!--#Include File="comm/inc.asp" -->
<!--#Include File="../include/md5.asp" -->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：/Manager/Admin_Member.asp
'= 摘    要：后台-会员管理文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2006-07-27
'====================================================================

Call EA_Manager.Chk_IsMaster

If Not EA_Manager.Chk_Power(Admin_Power,"22") Then 
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
	Dim ListName(7),ListValue()
	Dim WSQL,Temp,Tmp
	Dim KeyWord,Field
	
	Page	= EA_Pub.SafeRequest(2,"nowPage",0,1,0)
	KeyWord	= EA_Pub.SafeRequest(2,"keyword",1,"",0)
	Field	= EA_Pub.SafeRequest(2,"group",0,0,0)

	WSQL = " WHERE 1=1"

	If KeyWord <> "" Then WSQL = WSQL & " AND reg_name like '%"&KeyWord&"%'"
	If Field > 0 Then WSQL = WSQL & " AND user_group="&Field
	
	Temp=EA_M_DBO.Get_Group_List()
	Tmp = "(build-select),0 " & str_Comm_Select & ",0"
	If IsArray(Temp) Then
		ForTotal = UBound(Temp,2)

		For i=0 To ForTotal
			Tmp = Tmp & " " & Replace(Temp(1,i)," ","") & "," & Temp(0,i)
		Next
	End If
	Call EA_M_XML.AppInfo("UserGroup",Tmp)
	
	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_Member_Help",str_Member_Help)

	Call EA_M_XML.AppElements("Language_Member_Account",str_Member_Account)
	Call EA_M_XML.AppElements("Language_Member_State",str_Member_State)
	Call EA_M_XML.AppElements("Language_Member_GroupName",str_Member_GroupName)
	Call EA_M_XML.AppElements("Language_Member_RegDate",str_Member_RegDate)

	Call EA_M_XML.AppElements("Language_QuickSearchUser",str_QuickSearchUser)
	Call EA_M_XML.AppElements("Language_UserGroup",str_UserGroup)

	Call EA_M_XML.AppElements("Comm_Add_Operation",str_Comm_Add_Operation)
	Call EA_M_XML.AppElements("Comm_Del_Operation",str_Comm_Del_Operation)

	Call EA_M_XML.AppElements("Language_Comm_Bar_Operation",str_Comm_Bar_Operation)

	SQL="Select Count([Id]) From [NB_User] "&WSQL
	Count=EA_M_DBO.DB_Query(SQL)(0, 0)
	If Count>0 Then 
		If Rs.State=1 Then Rs.Close
		If iDataBaseType=0 Then
			SQL="Select a.[Id],Reg_Name,b.GroupName,a.Email,a.RegTime,IIF(State=0,'" & str_Comm_State_NoPass & "','" & str_Comm_State_Pass & "') From [NB_User] a Left Join [NB_UserGroup] b On a.User_Group=b.[Id] "&WSQL&" Order By a.[Id] Desc"
		Else
			SQL="Select a.[Id],Reg_Name,b.GroupName,a.Email,a.RegTime,Case State When 0 Then '" & str_Comm_State_NoPass & "' Else '" & str_Comm_State_Pass & "' End From [NB_User] a Left Join [NB_UserGroup] b On a.User_Group=b.[Id] "&WSQL&" Order By a.[Id] Desc"
		End If
		'Response.Write sql
		Rs.Open SQL,Conn,1,1
		If Not rs.eof And Not rs.bof Then 
			Rs.AbsolutePosition=Rs.AbsolutePosition+((Abs(Page)-1)*10)
			TopicList=Rs.GetRows(10)
		End If
		Rs.Close

		ListName(0) = "checkbox"
		ListName(1) = "ID"
		ListName(2) = "Account"
		ListName(3) = "State"
		ListName(4) = "GroupName"
		ListName(5) = "E-Mail"
		ListName(6) = "RegDate"
		ListName(7) = "action"
		ForTotal = Ubound(TopicList,2)

	    For i=0 To ForTotal
			ReDim Preserve ListValue(7,i)

			ListValue(0,i) = "checkbox"
			ListValue(1,i) = TopicList(0,i)
			ListValue(2,i) = TopicList(1,i)
			ListValue(3,i) = TopicList(5,i)
			ListValue(4,i) = TopicList(2,i)
			ListValue(5,i) = TopicList(3,i)
			ListValue(6,i) = TopicList(4,i)
			ListValue(7,i) = "action"
		Next

		Page = EA_M_XML.make(ListName,ListValue,Count)
	Else
		Page = EA_M_XML.make("","",0)
	End If

	Call EA_M_XML.Out(Page)
End Sub

Sub Add
	Dim PostId
	Dim i,Temp,Tmp
	Dim UserGroup,ArticleCount
	
	PostId=EA_Pub.SafeRequest(2,"ID",0,0,0)
	Call EA_M_XML.AppInfo("ID",PostId)

	Temp=EA_M_DBO.Get_Member_Info(PostId)
	If IsArray(Temp) Then 
		Call EA_M_XML.AppElements("Account",Temp(1,0))
		Call EA_M_XML.AppElements("Email",Temp(3,0))
		Call EA_M_XML.AppElements("RegTime",Temp(4,0))
		Call EA_M_XML.AppElements("LoginTotal",Temp(5,0))
		Call EA_M_XML.AppElements("NickName",Temp(6,0))
		Call EA_M_XML.AppElements("BirtDay",Temp(7,0))

		If Temp(2,0) Then
			Call EA_M_XML.AppElements("Sex",str_Member_Sex_Man)
		Else
			Call EA_M_XML.AppElements("Sex",str_Member_Sex_Woman)
		End If

		UserGroup=Temp(8,0)
		Call EA_M_XML.AppInfo("State",Temp(9,0))
	End If

	ArticleCount=EA_DBO.Get_FlorilegiumStat(Temp(6,0),PostId)(0,0)
	Call EA_M_XML.AppElements("ArticleCount",ArticleCount)
	
	Temp=EA_M_DBO.Get_Group_List()
	Tmp = "(build-select)," & UserGroup & " " & str_Comm_Select & ",0"
	If IsArray(Temp) Then
		ForTotal = UBound(Temp,2)

		For i=0 To ForTotal
			Tmp = Tmp & " " & Replace(Temp(1,i)," ","") & "," & Temp(0,i)
		Next
	End If
	Call EA_M_XML.AppInfo("UserGroup",Tmp)

	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_Member_Help",str_Member_Help)
	Call EA_M_XML.AppElements("Language_Member_MemberInfo",str_Member_MemberInfo)

	Call EA_M_XML.AppElements("Language_Member_Account",str_Member_Account)
	Call EA_M_XML.AppElements("Language_Member_State",str_Member_State)
	Call EA_M_XML.AppElements("Language_Member_Sex",str_Member_Sex)
	Call EA_M_XML.AppElements("Language_Member_NickName",str_Member_NickName)
	Call EA_M_XML.AppElements("Language_Member_LoginTotal",str_Member_LoginTotal)
	Call EA_M_XML.AppElements("Language_Member_BirtDay",str_Member_BirtDay)
	Call EA_M_XML.AppElements("Language_Member_GroupName",str_Member_GroupName)
	Call EA_M_XML.AppElements("Language_Member_RegDate",str_Member_RegDate)
	Call EA_M_XML.AppElements("Language_Member_IsLogin",str_Member_IsLogin)
	Call EA_M_XML.AppElements("Language_Member_ArticleTotal",str_Member_ArticleTotal)
	Call EA_M_XML.AppElements("Language_Member_LoginPassword",str_Member_LoginPassword)
	Call EA_M_XML.AppElements("Language_Member_PasswordEditInfo",str_Member_PasswordEditInfo)

	Call EA_M_XML.AppElements("btnSubmit",str_Comm_Save_Button)
	Call EA_M_XML.AppElements("btnReturn",str_Comm_Return_Button)

	Page = EA_M_XML.make("","",0)
	Call EA_M_XML.Out(Page)
End Sub

Sub Save
	Dim UserGroup,State,Password
	Dim PostId
	
	PostId	 = EA_Pub.SafeRequest(2,"ID",0,0,0)
	UserGroup= EA_Pub.SafeRequest(2,"UserGroup",0,0,0)
	State	 = EA_Pub.SafeRequest(2,"State",0,0,0)
	Password = EA_Pub.SafeRequest(2,"logPass",1,"",0)
		
	If PostId<>0 And UserGroup<>0 Then
		SQL="UPDATE NB_User SET User_Group = "&UserGroup&", State = "&State
		If Password <> "" Then SQL = SQL & ",Reg_Pass = '" & Md5(Password) & "'"
		SQL=SQL&" WHERE Id="&PostId
		EA_M_DBO.DB_Execute(SQL)
	End If
	
	Set Rs=Nothing
	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing
	
	Response.Write "1"
	Response.End
End Sub

Sub Del
	Dim IDs
	Dim i,Tmp

	IDs = Split(Request.Form("ID"), ",")
	ForTotal = UBound(IDs)

	For i = 0 To ForTotal
		Tmp = EA_Pub.SafeRequest(5,IDs(i),0,0,0)

		EA_M_DBO.Set_Member_Delete Tmp
	Next
	
	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing
	
	Response.Write "1"
	Response.End
End Sub
%>