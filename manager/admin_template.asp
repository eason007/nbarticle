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
'= 文件名称：/Manager/Admin_Template.asp
'= 摘    要：后台-模版管理文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2007-10-27
'====================================================================

Call EA_Manager.Chk_IsMaster

If Not EA_Manager.Chk_Power(Admin_Power,"44") Then 
	ErrMsg=str_Comm_NotAccess
	Call EA_Manager.Error(1)
End If

Dim Atcion
Atcion=Request ("action")

Select Case LCase(Atcion)
Case "add"
	Call Edit
Case "save"
	Call Save
Case "del"
	Call DelTemplate
Case "default"
	Call SaveDefaultTemplate
Case "clone"
	Call CloneTemplate
Case Else
	Call Main
End Select
Call EA_Pub.Close_Obj
Set EA_Pub=Nothing

Sub Main
	Dim i
	Dim TopicList
	Dim ListName(5),ListValue()
	Dim ForTotal

	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_Template_Info",str_Template_Info)

	Call EA_M_XML.AppElements("Language_Template_TempName",str_Template_TempName)
	Call EA_M_XML.AppElements("Language_Template_Default",str_Template_Default)

	Call EA_M_XML.AppElements("Comm_Add_Operation",str_Comm_Add_Operation)
	Call EA_M_XML.AppElements("Comm_Del_Operation",str_Comm_Del_Operation)

	Call EA_M_XML.AppElements("Language_Comm_Bar_Operation",str_Comm_Bar_Operation)

	TopicList=EA_M_DBO.Get_Template_List()

	ListName(0) = "checkbox"
	ListName(1) = "ID"
	ListName(2) = "IDs"
	ListName(3) = "Name"
	ListName(4) = "IsDefault"
	ListName(5) = "action"
	ForTotal = Ubound(TopicList,2)
	
	For i=0 To ForTotal
		ReDim Preserve ListValue(5,i)

		ListValue(0,i) = "checkbox"
		ListValue(1,i) = TopicList(0,i)
		ListValue(2,i) = TopicList(0,i)
		ListValue(3,i) = TopicList(1,i)
		If TopicList(2,i) Then
			ListValue(4,i) = "<strong>√</strong>"
		Else
			ListValue(4,i) = "<font color=""red""><strong>×</strong></font>"
		End If
		ListValue(5,i) = "action"
	Next

	Page = EA_M_XML.make(ListName,ListValue,Ubound(TopicList,2)+1)

	Call EA_M_XML.Out(Page)
End Sub

Sub Edit
	Dim TemplateName,TempStr,StyleId
	Dim i
	
	StyleId=EA_Pub.SafeRequest(2,"ID",0,0,0)
	Call EA_M_XML.AppInfo("ID",StyleId)
	
	TempStr=EA_M_DBO.Get_Template_Info(StyleId)

	If IsArray(TempStr) Then TemplateName=TempStr(1,0)
	If Not IsArray(TempStr) Then ReDim TempStr(13,1)

	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_Template_Info",str_Template_Info)
	Call EA_M_XML.AppElements("Language_Template_Manager",str_Template_Manager)

	Call EA_M_XML.AppElements("Language_Template_TempName",str_Template_TempName)

	Call EA_M_XML.AppElements("Language_Template_Css_Info",str_Template_Css_Info)
	Call EA_M_XML.AppElements("Language_Template_Head_Info",str_Template_Head_Info)
	Call EA_M_XML.AppElements("Language_Template_Foot_Info",str_Template_Foot_Info)
	Call EA_M_XML.AppElements("Language_Template_Index_Info",str_Template_Index_Info)
	Call EA_M_XML.AppElements("Language_Template_List_Info",str_Template_List_Info)
	Call EA_M_XML.AppElements("Language_Template_View_Info",str_Template_View_Info)
	Call EA_M_XML.AppElements("Language_Template_Search_Info",str_Template_Search_Info)
	Call EA_M_XML.AppElements("Language_Template_ImgList_Info",str_Template_ImgList_Info)
	Call EA_M_XML.AppElements("Language_Template_MemberList_Info",str_Template_MemberList_Info)
	Call EA_M_XML.AppElements("Language_Template_Error_Info",str_Template_Error_Info)
	Call EA_M_XML.AppElements("Language_Template_Success_Info",str_Template_Success_Info)
	Call EA_M_XML.AppElements("Language_Template_MemberList_Info",str_Template_MemberList_Info)
	Call EA_M_XML.AppElements("Language_Template_Login_Info",str_Template_Login_Info)

	Call EA_M_XML.AppElements("Language_Template_CheckForm",str_Template_CheckForm)
	Call EA_M_XML.AppElements("Language_Template_CurrentEdit",str_Template_CurrentEdit)

	Call EA_M_XML.AppElements("btnSubmit",str_Comm_Save_Button)
	Call EA_M_XML.AppElements("btnReturn",str_Comm_Return_Button)

	Call EA_M_XML.AppElements("TemplateName1",TemplateName)

	For i = 1 To 13
		Call EA_M_XML.AppElements("Comm_Edit_" & i,str_Comm_Edit_Operation)
	Next

	Call EA_M_XML.AppInfo("TemplateName",TemplateName)
	Call EA_M_XML.AppInfo("TemplateCSS",TempStr(2,0))
	Call EA_M_XML.AppInfo("TemplateHead",TempStr(3,0))
	Call EA_M_XML.AppInfo("TemplateFoot",TempStr(4,0))
	Call EA_M_XML.AppInfo("TemplateIndex",TempStr(5,0))
	Call EA_M_XML.AppInfo("TemplateList",TempStr(6,0))
	Call EA_M_XML.AppInfo("TemplateView",TempStr(7,0))
	Call EA_M_XML.AppInfo("TemplateSearch",TempStr(8,0))
	Call EA_M_XML.AppInfo("TemplateMemberList",TempStr(9,0))
	Call EA_M_XML.AppInfo("TemplateImgList",TempStr(10,0))
	Call EA_M_XML.AppInfo("TemplateError",TempStr(11,0))
	Call EA_M_XML.AppInfo("TemplateSuccess",TempStr(12,0))
	Call EA_M_XML.AppInfo("TemplateLogin",TempStr(13,0))

	Page = EA_M_XML.make("","",0)
	Call EA_M_XML.Out(Page)
End Sub

Sub Save
	If rs.state=1 Then rs.close

	Dim TemplateContent,TemplateTag
	Dim TemplateID
	
	TemplateID		= EA_Pub.SafeRequest(2,"ID",0,0,0)
	TemplateContent = EA_Pub.SafeRequest(2,"TemplateContent",1,"",-1)
	TemplateTag		= EA_Pub.SafeRequest(2,"TemplateTag",1,"",-1)

	If TemplateTag = "" Then Response.Write "-1":Response.End
	
	If TemplateID<>0 Then
		rs.open "select * from NB_Template where id="&TemplateID,conn,2,3
	Else
		rs.open "select * from NB_Template",conn,2,3
		rs.addnew
	End If
	
	If TemplateTag = "Name" Then
		rs("Temp_" & TemplateTag)=TemplateContent
	Else
		rs("Page_" & TemplateTag)=TemplateContent

		If LCase(TemplateTag) = "css" Then
			Dim re
			Set re=new RegExp
			re.IgnoreCase =true
			re.Global=True

			re.Pattern="<style([^>]*)>"
			TemplateContent=re.Replace(TemplateContent,"")

			TemplateContent=Replace(TemplateContent,"</style>","")

			TemplateContent=Replace(TemplateContent,"{$SystemPath$}",SystemFolder)

			Call EA_Pub.Save_HtmlFile("../common/css/style_" & TemplateID & ".css",TemplateContent)
		End If
	End If
	rs.update
		
	Rs.Close:Set Rs=Nothing
	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing
	
	Response.Write "1"
	Response.End
End Sub

Sub CloneTemplate
	Dim TempId
	TempId=EA_Pub.SafeRequest(2,"ID",0,0,0)
	
	EA_M_DBO.Set_Template_Clone TempId
	
	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing
	
	Response.Write "1"
	Response.End
End Sub

Sub SaveDefaultTemplate
	Dim DefId
	DefId=EA_Pub.SafeRequest(2,"ID",0,0,0)
	
	If DefId>0 Then EA_M_DBO.Set_DefaultTemplate DefId
	
	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing
	
	Response.Write "1"
	Response.End
End Sub

Sub DelTemplate
	Dim IDs
	Dim i,Tmp
	Dim TempStr
	Dim IsDel
	Dim ForTotal

	IDs = Split(Request("ID"), ",")
	ForTotal = UBound(IDs)

	For i = 0 To ForTotal
		IsDel=False

		Tmp = EA_Pub.SafeRequest(5,IDs(i),0,0,0)

		TempStr=EA_M_DBO.Get_Template_Info(Tmp)
		If IsArray(TempStr) Then 
			If TempStr(13,0)=0 Then 
				IsDel=True
			Else
				IsDel=False
			End If
		End If

		If IsDel Then EA_M_DBO.Set_Template_Delete Tmp
	Next
	
	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing
	
	Response.Write "1"
	Response.End
End Sub 
%>