<!--#Include File="../include/inc.asp"-->
<!--#Include File="../include/md5.asp"-->
<!--#Include File="cls_db.asp"-->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：Member/Login.asp
'= 摘    要：会员-登陆文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-03-20
'====================================================================

Dim Action
Action=Request.QueryString ("action")

Select Case LCase(Action)
Case "login"
	Call Chk_Login
Case "logout"
	Call Get_Logout
End Select
EA_Pub.Close_Obj
Set EA_Pub=Nothing

Sub Chk_Login
	Call EA_Pub.Chk_Post
	
	Dim Login_Accout, Login_Password, SaveTime
	Dim Mem_Info(4)
	Dim TempArray, i, Key, Temp
	Dim EA_Mem_DBO

	Set EA_Mem_DBO = New cls_Member_DBOperation
	
	Login_Accout	= EA_Pub.SafeRequest(2, "username", 1,"", 1)
	Login_Password	= EA_Pub.SafeRequest(2, "password", 1,"", -1)
	Login_Password	= MD5(Login_Password)
	SaveTime		= EA_Pub.SafeRequest(2, "savetimes", 0, 0, 0)
	
	Temp = EA_DBO.Get_MemberLogin(Login_Accout)
	If Not IsArray(Temp) Then Call EA_Pub.ShowErrMsg(2, 2)
	If Trim(Temp(1,0)) <> Login_Password Then Call EA_Pub.ShowErrMsg(29, 2)
	If Not Temp(2,0) Then Call EA_Pub.ShowErrMsg(30, 2)
	
	EA_Pub.Mem_Info(0) = Temp(0, 0)
	EA_Pub.Mem_Info(1) = Login_Accout
	EA_Pub.Mem_Info(2) = Temp(3, 0)
	EA_Pub.Mem_Info(3) = Temp(4, 0)
	
	Call EA_Pub.Get_Member_GroupSetting(EA_Pub.Mem_Info(3))
	
	If EA_Pub.Mem_GroupSetting(1)="0" Then Call EA_Pub.ShowErrMsg(31, 2)
	If EA_Pub.Mem_GroupSetting(5)="1" And EA_Pub.Chk_SystemTimer(EA_Pub.Mem_GroupSetting(4)) Then Call EA_Pub.ShowErrMsg(33, 1)
	
	Randomize
	Key=CStr(Int((999999-1+100000)*Rnd+1))
	
	EA_Pub.Mem_Info(4)=Key
	EA_Pub.Mem_Info(5)=EA_Pub.Mem_GroupSetting(0)

	Session("UserData") = Join(EA_Pub.Mem_Info,"|")
	Response.Cookies("UserData") = Session("UserData")

	If SaveTime=10 Then Response.Cookies("UserData").Expires=Date()+720

	Call EA_Mem_DBO.Set_MemberLoginKey(EA_Pub.Get_UserIp,Key,EA_Pub.Mem_Info(0))

	Call EA_Pub.ShowErrMsg(39, 2)
End Sub

Sub Get_Logout()
	Session("UserKey")=Empty
	Session("UserData")=Empty
	
	Response.Cookies("UserData")=Empty
	
	Response.Redirect SystemFolder
End Sub
%> 
