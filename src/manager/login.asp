<!--#Include File="comm/inc.asp" -->
<!--#Include File="../include/md5.asp" -->
<%
'====================================================================
'= Team Elite - EliteCMS
'= Copyright (c) 2004 - 2008 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：/Manager/Login.asp
'= 摘    要：后台-登陆文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2007-03-09
'====================================================================

Dim Action
Action=Request.QueryString ("action")

Select Case LCase(Action)
Case "login"
	Call Chk_Login()
Case "logout"
	Call Get_LogOut()
End Select

Sub Chk_Login()
	Dim Login_Name,Login_Pass
	Dim Login_Key,Login_Id
	Dim Come_Ip
	Dim Temp
	
	FoundErr=False

	Login_Name=EA_Pub.SafeRequest(2,"name",1,"",0)
	Login_Pass=EA_Pub.SafeRequest(2,"pass",1,"",0)

	Temp=EA_M_DBO.Get_Master_Login(Login_Name)
	If Not IsArray(Temp) Then
		ErrMsg="对不起，输入的用户名不存在！"
	    FoundErr=True
	Else
		If Temp(0,0)=Md5(Login_Pass) Then
			If Temp(2,0) Then
				Randomize
				Login_Key=Cstr(Int((999999-1+100000)*Rnd+1))
				
				Login_Id=Temp(1,0)

				Session(sCacheName&"master_id")	 = Login_Id
				Session(sCacheName&"master_name")= Login_Name
				Session(sCacheName&"master_key") = Login_Key

				Response.Cookies("UserName") = Login_Name
	
				Come_Ip=EA_Pub.Get_UserIp
				
				EA_M_DBO.Set_Master_LoginLog Login_Key,Come_Ip,Login_Id
			Else
				ErrMsg="对不起，您的帐号没有进入管理中心的权限，请勿进入！"
				FoundErr=True
			End If
		Else
			ErrMsg="对不起，密码错误！"
			FoundErr=True
	    End If		
	End If
	
	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing
	 
	If FoundErr Then 
		Response.Clear 
		Response.Write ErrMsg
		Response.End 
	Else
		Response.Redirect "admin_index.htm"
	End If
End Sub

Sub Get_LogOut()
	Response.Clear 
	
	Session.Abandon 
	
	Set Rs=Nothing
	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing
	
	Response.Redirect "./"
	Response.Flush 
	Response.End 
End Sub
%>