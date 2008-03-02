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
'= 文件名称：/Manager/Admin_Config.asp
'= 摘    要：后台-系统设定文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-03-02
'====================================================================

Call EA_Manager.Chk_IsMaster

If Not EA_Manager.Chk_Power(Admin_Power,"01") Then 
	ErrMsg=str_Comm_NotAccess
	Call EA_Manager.Error(1)
End If

Dim Setting,Action
Action=Request.Form ("action")

Select Case LCase(Action)
Case "save"
	Call Save
Case Else 
	Call Main
End Select
Call EA_Pub.Close_Obj
Set EA_Pub=Nothing

Sub Main
	Dim Tmp,TempStr,Source,BadWord
	Dim Temp,i
	
	Temp=EA_DBO.Get_System_Info()
	If IsArray(Temp) Then 
		If Temp(5,0)<>"" And Not IsNull(Temp(5,0)) Then 
			Tmp=Split(Temp(5,0),",")
			TempStr=Temp(5,0)
			Source=Temp(6,0)
			BadWord=Temp(7,0)

			If Ubound(Tmp)<26 Then 
				TempStr=TempStr&",,,,,,,,,,,,,,,,,,,,,,,,"
				Tmp=Split(TempStr,",")
			End If

			Setting=Tmp
		End If
	End If

	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_Config_Help",str_Config_Help)
	Call EA_M_XML.AppElements("Language_Config_Base",str_Config_Base)
	Call EA_M_XML.AppElements("Language_Config_WebSiteName",str_Config_WebSiteName)
	Call EA_M_XML.AppElements("Language_Config_WebSiteURL",str_Config_WebSiteURL)
	Call EA_M_XML.AppElements("Language_Config_WebSiteState",str_Config_WebSiteState)
	Call EA_M_XML.AppElements("Language_Config_WebSiteState_Open",str_Config_WebSiteState_Open)
	Call EA_M_XML.AppElements("Language_Config_WebSiteState_Close",str_Config_WebSiteState_Close)
	Call EA_M_XML.AppElements("Language_Config_ClosedMsg",str_Config_ClosedMsg)
	Call EA_M_XML.AppElements("Language_Config_SystemTimer",str_Config_SystemTimer)
	Call EA_M_XML.AppElements("Language_Config_SystemTimer_Value",str_Config_SystemTimer_Value)
	Call EA_M_XML.AppElements("Language_Config_SystemTimer_Value_Help",str_Config_SystemTimer_Value_Help)
	Call EA_M_XML.AppElements("Language_Config_PageKeyWord",str_Config_PageKeyWord)
	Call EA_M_XML.AppElements("Language_Config_PageKeyWord_Help",str_Config_PageKeyWord_Help)
	Call EA_M_XML.AppElements("Language_Config_PageDescription",str_Config_PageDescription)
	Call EA_M_XML.AppElements("Language_Config_PageDescription_Help",str_Config_PageDescription_Help)
	Call EA_M_XML.AppElements("Language_Config_SystemMode",str_Config_SystemMode)
	Call EA_M_XML.AppElements("Language_Config_SystemStyle",str_Config_SystemStyle)
	Call EA_M_XML.AppElements("Language_Config_SystemIndexMode",str_Config_SystemIndexMode)
	Call EA_M_XML.AppElements("Language_Config_Article",str_Config_Article)
	Call EA_M_XML.AppElements("Language_Config_MemberEditor",str_Config_MemberEditor)
	Call EA_M_XML.AppElements("Language_Config_AutoRemote",str_Config_AutoRemote)
	Call EA_M_XML.AppElements("Language_Config_AutoRemote_Help",str_Config_AutoRemote_Help)
	Call EA_M_XML.AppElements("Language_Config_DefaultPoster",str_Config_DefaultPoster)
	Call EA_M_XML.AppElements("Language_Config_Reg",str_Config_Reg)
	Call EA_M_XML.AppElements("Language_Config_UserRegEnable",str_Config_UserRegEnable)
	Call EA_M_XML.AppElements("Language_Config_EmailById",str_Config_EmailById)
	Call EA_M_XML.AppElements("Language_Config_RegWaitAdmin",str_Config_RegWaitAdmin)
	Call EA_M_XML.AppElements("Language_Config_EMail",str_Config_EMail)
	Call EA_M_XML.AppElements("Language_Config_EMailAddress",str_Config_EMailAddress)
	Call EA_M_XML.AppElements("Language_Config_SMTPServerAddress",str_Config_SMTPServerAddress)
	Call EA_M_XML.AppElements("Language_Config_SMTPLoginAccout",str_Config_SMTPLoginAccout)
	Call EA_M_XML.AppElements("Language_Config_SMTPLoginPassWord",str_Config_SMTPLoginPassWord)
	Call EA_M_XML.AppElements("Language_Config_Other",str_Config_Other)
	Call EA_M_XML.AppElements("Language_Config_GhostPostReview",str_Config_GhostPostReview)
	Call EA_M_XML.AppElements("Language_Config_ReviewWaitAdmin",str_Config_ReviewWaitAdmin)
	Call EA_M_XML.AppElements("Language_Config_GhostPostVote",str_Config_GhostPostVote)
	Call EA_M_XML.AppElements("Language_Config_BadWord",str_Config_BadWord)
	Call EA_M_XML.AppElements("Language_Config_BadWord_Help",str_Config_BadWord_Help)
	Call EA_M_XML.AppElements("Language_Config_ArticleFrom",str_Config_ArticleFrom)
	Call EA_M_XML.AppElements("Language_Config_ArticleFrom_Help",str_Config_ArticleFrom_Help)
	Call EA_M_XML.AppElements("Language_Config_SEO",str_Config_SEO)

	For i= 1 To 6
		Call EA_M_XML.AppElements("btnSubmit" & i,str_Comm_Save_Button)
		Call EA_M_XML.AppElements("btnReset" & i,str_Comm_Reset_Button)
	Next

	For i= 1 To 11
		Call EA_M_XML.AppElements("Language_Comm_Yes" & i,str_Comm_Yes)
		Call EA_M_XML.AppElements("Language_Comm_No" & i,str_Comm_No)
	Next

	Call EA_M_XML.AppInfo("SiteName",Setting(0))
	Call EA_M_XML.AppInfo("SiteState",Setting(1))
	Call EA_M_XML.AppInfo("ClosedMsg",Setting(2))
	Call EA_M_XML.AppInfo("IsClose",Setting(3))
	Call EA_M_XML.AppInfo("SystemTimer",Setting(4))
	Call EA_M_XML.AppInfo("isreg",Setting(7))
	Call EA_M_XML.AppInfo("isemail",Setting(8))
	Call EA_M_XML.AppInfo("isadmin",Setting(9))
	Call EA_M_XML.AppInfo("isvote",Setting(10))
	Call EA_M_XML.AppInfo("SiteURL",Setting(11))
	Call EA_M_XML.AppInfo("mail",Setting(12))
	Call EA_M_XML.AppInfo("mail_n",Setting(13))
	Call EA_M_XML.AppInfo("mail_p",Setting(14))
	Call EA_M_XML.AppInfo("mail_s",Setting(15))
	Call EA_M_XML.AppInfo("keyword",Setting(16))
	Call EA_M_XML.AppInfo("description",Setting(17))
	Call EA_M_XML.AppInfo("skin",Setting(18))
	Call EA_M_XML.AppInfo("isreview",Setting(19))
	Call EA_M_XML.AppInfo("isreview_admin",Setting(20))
	Call EA_M_XML.AppInfo("author",Setting(21))
	Call EA_M_XML.AppInfo("autoremote",Setting(22))
	Call EA_M_XML.AppInfo("member_editor",Setting(24))
	Call EA_M_XML.AppInfo("index",Setting(26))
	Call EA_M_XML.AppInfo("badword",BadWord)
	Call EA_M_XML.AppInfo("source",Source)

	Page = EA_M_XML.make("","",0)

	Call EA_M_XML.Out(Page)
End Sub

Function OptionReplace(ByRef vNameArray,ByRef vValueArray,ByRef iCheckValue)
	Dim i
	Dim str
	Dim ForTotal

	str = "(build-select)," & iCheckValue
	ForTotal = UBound(vValueArray)

	For i = 0 To ForTotal
		str = str & " " & vNameArray(i) & "," & vValueArray(i)
	Next

	OptionReplace = str
End Function

Sub Save
	Dim Source,BadWord
	Dim SiteUrl
	
	SiteUrl=EA_Pub.SafeRequest(2,"SiteURL",1,"",0)
	If Right(SiteUrl,1)<>"/" And Right(SiteUrl,1)<>"\" Then SiteUrl=SiteUrl&"/"
	Source=EA_Pub.SafeRequest(2,"source",1,"",1)
	If Right(Source,1)=";" Then Source=Left(Source,Len(Source)-1)
	BadWord=EA_Pub.SafeRequest(2,"badword",1,"",1)
	If Right(BadWord,1)=";" Then BadWord=Left(BadWord,Len(BadWord)-1)

	Setting=EA_Pub.SafeRequest(2,"SiteName",1,"",1)&","
	Setting=Setting&EA_Pub.SafeRequest(2,"SiteState",0,1,0)&","
	Setting=Setting&EA_Pub.SafeRequest(2,"ClosedMsg",1,"",0)&","
	Setting=Setting&EA_Pub.SafeRequest(2,"IsClose",0,0,0)&","
	Setting=Setting&EA_Pub.SafeRequest(2,"SystemTimer",1,"1|23",-1)&","
	Setting=Setting&"0,"
	Setting=Setting&"0,"
	Setting=Setting&EA_Pub.SafeRequest(2,"isreg",0,1,0)&","
	Setting=Setting&EA_Pub.SafeRequest(2,"isemail",0,1,0)&","
	Setting=Setting&EA_Pub.SafeRequest(2,"isadmin",0,0,0)&","
	Setting=Setting&EA_Pub.SafeRequest(2,"isvote",0,1,0)&","
	Setting=Setting&SiteUrl&","
	Setting=Setting&EA_Pub.SafeRequest(2,"mail",1,"",1)&","
	Setting=Setting&EA_Pub.SafeRequest(2,"mail_n",1,"",1)&","
	Setting=Setting&EA_Pub.SafeRequest(2,"mail_p",1,"",1)&","
	Setting=Setting&EA_Pub.SafeRequest(2,"mail_s",1,"",1)&","
	Setting=Setting&EA_Pub.SafeRequest(2,"keyword",1,"",1)&","
	Setting=Setting&Replace(EA_Pub.SafeRequest(2,"description",1,"",1),",","，")&","
	Setting=Setting&EA_Pub.SafeRequest(2,"skin",0,1,1)&","
	Setting=Setting&EA_Pub.SafeRequest(2,"isreview",0,1,1)&","
	Setting=Setting&EA_Pub.SafeRequest(2,"isreview_admin",0,0,1)&","
	Setting=Setting&EA_Pub.SafeRequest(2,"author",1,"",1)&","
	Setting=Setting&EA_Pub.SafeRequest(2,"autoremote",1,"",1)&","
	Setting=Setting&","
	Setting=Setting&EA_Pub.SafeRequest(2,"member_editor",1,"",1)&","
	Setting=Setting&","
	Setting=Setting&EA_Pub.SafeRequest(2,"index",1,"",0)&","
	
	SQL="UpDate [NB_System] Set Info='"&Setting&"',Source='"&Source&"',BadWord='"&BadWord&"'"
	EA_M_DBO.DB_Execute(SQL)

	Setting = Split(Setting,",")
	Setting(14) = ""
	Setting = Join(Setting,",")
	
	Dim EA_Ini
	Set EA_Ini=New cls_Ini
	EA_Ini.OpenFile	= EA_Pub.sIniFilePath

	Call EA_Ini.WriteNode("System","Info",Setting)
	EA_Ini.Save
	EA_Ini.Close
	Set EA_Ini=Nothing

	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing
	
	Response.Write "1"
	Response.End
End Sub
%>
