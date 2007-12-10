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
'= 文件名称：/Manager/Admin_Link.asp
'= 摘    要：后台-联盟管理文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2006-07-27
'====================================================================

Call EA_Manager.Chk_IsMaster

If Not EA_Manager.Chk_Power(Admin_Power,"02") Then 
	ErrMsg=str_Comm_NotAccess
	Call EA_Manager.Error(1)
End If

Dim Action
Dim ForTotal
Action=Request("action")

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
	
	Page=EA_Pub.SafeRequest(2,"nowPage",0,1,0)

	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_Friend_Help",str_Friend_Help)

	Call EA_M_XML.AppElements("Language_Friend_Order",str_Friend_Order)
	Call EA_M_XML.AppElements("Language_Friend_SiteName",str_Friend_SiteName)
	Call EA_M_XML.AppElements("Language_Friend_SiteLogo",str_Friend_SiteLogo)
	Call EA_M_XML.AppElements("Language_Friend_Location",str_Friend_Location)
	Call EA_M_XML.AppElements("Language_Friend_State",str_Friend_State)

	Call EA_M_XML.AppElements("Comm_Add_Operation",str_Comm_Add_Operation)
	Call EA_M_XML.AppElements("Comm_Del_Operation",str_Comm_Del_Operation)

	Call EA_M_XML.AppElements("Language_Comm_Bar_Operation",str_Comm_Bar_Operation)

	Count=EA_M_DBO.Get_Friend_Total()(0,0)
	If Count>0 Then 
		TopicList=EA_M_DBO.Get_Friend_List(Page,10)

		ListName(0) = "checkbox"
		ListName(1) = "ID"
		ListName(2) = "SiteName"
		ListName(3) = "SiteImg"
		ListName(4) = "Location"
		ListName(5) = "Order"
		ListName(6) = "State"
		ListName(7) = "action"
		ForTotal = UBound(TopicList,2)

		For i = 0 To ForTotal
			ReDim Preserve ListValue(7,i)

			ListValue(0,i) = "checkbox"
			ListValue(1,i) = TopicList(0,i)
			ListValue(2,i) = "<a href='" & TopicList(3,i) & "' target='_blank'>" & TopicList(1,i)  & "</a>"
			If LCase(TopicList(2,i)) = "images/space.gif" Or TopicList(2,i) = "" Then
				ListValue(3,i) = "<img src='images/space.gif' alt='' />"
			Else
				ListValue(3,i) = "<img src='" & TopicList(2,i) & "' alt='' width='83' />"
			End If
			ListValue(4,i) = TopicList(4,i)
			ListValue(5,i) = TopicList(5,i)
			ListValue(6,i) = TopicList(6,i)
			ListValue(7,i) = "action"
		Next

		Page = EA_M_XML.make(ListName,ListValue,Count)
	Else
		Page = EA_M_XML.make("","",0)
	End If

	Call EA_M_XML.Out(Page)
End Sub

Sub Add
	Dim PostId,Info
	Dim ColumnId,State,Style
	Dim Level,Temp,Tmp,i
	
	PostId=EA_Pub.SafeRequest(2,"ID",0,0,0)
	Call EA_M_XML.AppInfo("ID",PostId)

	Info=EA_M_DBO.Get_Friend_Info(PostId)
	If IsArray(Info) Then 
		Call EA_M_XML.AppInfo("LinkName",Info(0,0))
		Call EA_M_XML.AppInfo("LinkURL",Info(1,0))
		Call EA_M_XML.AppInfo("LinkImgPath",Info(2,0))
		Call EA_M_XML.AppInfo("LinkInfo",Info(3,0))
		Call EA_M_XML.AppInfo("OrderNum",Info(5,0))

		ColumnId = Info(4,0)
		State	 = Info(6,0)
		Style	 = Info(7,0)
	End If

	Temp=EA_DBO.Get_Column_List()
	Tmp = "(build-select)," & ColumnId & " " & str_Friend_Index & ",0"
	If IsArray(Temp) Then
		ForTotal = UBound(Temp,2)

		For i=0 To ForTotal
			Level=(Len(Temp(2,i))/4-1)*3

			Tmp = Tmp & " ├" & String(Level,"-") & Replace(Temp(1,i)," ","") & "," & Temp(0,i)
		Next
	End If
	Call EA_M_XML.AppInfo("ColumnList",Tmp)

	Tmp = "(build-select)," & CInt(State) & " " & str_Comm_State_Pass & ",1 " & str_Comm_State_NoPass & ",0"
	Call EA_M_XML.AppInfo("state",Tmp)

	Tmp = "(build-select)," & Style & " " & str_Friend_Style_Img & ",1 " & str_Friend_Style_Txt & ",0"
	Call EA_M_XML.AppInfo("style",Tmp)

	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_Friend_Help",str_Friend_Help)
	Call EA_M_XML.AppElements("Language_Friend_Input_Info",str_Friend_Input_Info)

	Call EA_M_XML.AppElements("Language_Friend_Order",str_Friend_Order)
	Call EA_M_XML.AppElements("Language_Friend_SiteName",str_Friend_SiteName)
	Call EA_M_XML.AppElements("Language_Friend_SiteLogo",str_Friend_SiteLogo)
	Call EA_M_XML.AppElements("Language_Friend_SiteURL",str_Friend_SiteURL)
	Call EA_M_XML.AppElements("Language_Friend_SiteInfo",str_Friend_SiteInfo)
	Call EA_M_XML.AppElements("Language_Friend_Location",str_Friend_Location)
	Call EA_M_XML.AppElements("Language_Friend_State",str_Friend_State)
	Call EA_M_XML.AppElements("Language_Friend_Style",str_Friend_Style)

	Call EA_M_XML.AppElements("btnSubmit",str_Comm_Save_Button)
	Call EA_M_XML.AppElements("btnReturn",str_Comm_Return_Button)

	Page = EA_M_XML.make("","",0)
	Call EA_M_XML.Out(Page)
End Sub

Sub Save
	Dim LinkName,LinkURL,LinkImgPath,LinkInfo,ColumnId,OrderNum,State,Style
	Dim PostId

	FoundErr = False

	PostId		= EA_Pub.SafeRequest(2,"ID",0,0,0)
	LinkName	= EA_Pub.SafeRequest(2,"LinkName",1,"",1)
	LinkURL		= EA_Pub.SafeRequest(2,"LinkURL",1,"",1)
	LinkImgPath = EA_Pub.SafeRequest(2,"LinkImgPath",1,"",1)
	LinkInfo	= EA_Pub.SafeRequest(2,"LinkInfo",1,"",1)
	ColumnId	= EA_Pub.SafeRequest(2,"ColumnList",0,0,0)
	OrderNum	= EA_Pub.SafeRequest(2,"OrderNum",0,0,0)
	State		= EA_Pub.SafeRequest(2,"state",0,0,0)
	Style		= EA_Pub.SafeRequest(2,"style",0,0,0)

	If LinkName="" Or Len(LinkName)>50 Then FoundErr = True
	If LinkURL="" Or Len(LinkURL)>150 Then FoundErr = True
	If Style=1 And (LinkImgPath="" Or Len(LinkImgPath)>150) Then FoundErr = True
	
	If FoundErr Then
		Response.Write "-1"
	Else
		If Rs.State=1 Then rs.Close
		If PostId<>0 Then
			Sql="Select * From [NB_FriendLink] Where [Id]="&PostId
			rs.Open Sql,Conn,2,2
		Else
			rs.Open "NB_FriendLink",Conn,2,2
			rs.AddNew
		End If
			rs("LinkName")=LinkName
			rs("LinkUrl")=LinkURL
			rs("LinkImgPath")=LinkImgPath
			rs("LinkInfo")=LinkInfo
			rs("ColumnId")=ColumnId
			rs("OrderNum")=OrderNum
			rs("State")=State
			rs("Style")=Style
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

		EA_M_DBO.Set_Friend_Delete Tmp
	Next
	
	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing
	
	Response.Write "1"
	Response.End
End Sub
%>