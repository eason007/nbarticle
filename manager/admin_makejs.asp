<!--#Include File="comm/inc.asp" -->
<!--#Include File="comm/cls_makejs.asp" -->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：/Manager/Admin_MakeJs.asp
'= 摘    要：后台-自定义Js定义管理文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2006-08-19
'====================================================================

Call EA_Manager.Chk_IsMaster

If Not EA_Manager.Chk_Power(Admin_Power,"46") Then 
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
Case "preview"
	Call Preview
Case "make"
	Call MakeJs
Case Else
	Call Main
End Select
Call EA_Pub.Close_Obj
Set EA_Pub=Nothing

Sub Main
	Dim Count,Page,i
	Dim TopicList
	Dim ListName(5),ListValue()
	
	Page=EA_Pub.SafeRequest(2,"nowPage",0,1,0)

	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_Js_Help",str_Js_Help)

	Call EA_M_XML.AppElements("Language_Js_Title",str_Js_Title)
	Call EA_M_XML.AppElements("Language_Js_Detail",str_Js_Detail)
	Call EA_M_XML.AppElements("Language_Js_TransferPath",str_Js_FilePath)

	Call EA_M_XML.AppElements("Comm_Add_Operation",str_Comm_Add_Operation)
	Call EA_M_XML.AppElements("Comm_Del_Operation",str_Comm_Del_Operation)

	Call EA_M_XML.AppElements("Language_Comm_Bar_Operation",str_Comm_Bar_Operation)

	Call EA_M_XML.AppElements("Language_Js_UpDate",str_Js_UpDate)
	Call EA_M_XML.AppElements("Language_Js_Preview",str_Js_Preview)
	Call EA_M_XML.AppElements("Language_Js_UpDateSystemJs",str_Js_UpDateSystemJs)

	TopicList=EA_M_DBO.Get_Js_List()
	If IsArray(TopicList) Then
		ListName(0) = "checkbox"
		ListName(1) = "ID"
		ListName(2) = "Title"
		ListName(3) = "Detail"
		ListName(4) = "TransferPath"
		ListName(5) = "action"
		ForTotal = Ubound(TopicList,2)

	    For i=0 To ForTotal
			ReDim Preserve ListValue(5,i)

			ListValue(0,i) = "checkbox"
			ListValue(1,i) = TopicList(0,i)
			ListValue(2,i) = TopicList(1,i)
			ListValue(3,i) = TopicList(2,i)
			ListValue(4,i) = "<textarea name=""message"" cols=""45"" rows=""2"" wrap=""VIRTUAL"" onclick=""this.focus();this.select()""><script type=""text/javascript"" src=""" & SystemFolder & "jsfiles/" & TopicList(3,i) & ".js"" charset=""utf-8""></script></textarea>"
			ListValue(5,i) = "action"
		Next

		Page = EA_M_XML.make(ListName,ListValue,Count)
	Else
		Page = EA_M_XML.make("","",0)
	End If

	Call EA_M_XML.Out(Page)
End Sub

Sub Add
	Dim PostId,Info
	Dim Temp,Setting,Tmp,i,Level
	
	PostId=EA_Pub.SafeRequest(2,"ID",0,0,0)
	Call EA_M_XML.AppInfo("ID",PostId)

	Temp=EA_M_DBO.Get_Js_Info(PostId)
	If IsArray(Temp) Then 
		Call EA_M_XML.AppInfo("Title",Temp(0,0))
		Call EA_M_XML.AppInfo("Info",Temp(1,0))
		Call EA_M_XML.AppInfo("FileName",Replace(Temp(2,0), ".js", "", 1, -1, 1))

		Setting=Split(Temp(3,0),"|")

		Call EA_M_XML.AppInfo("IncludeChildColumn",Setting(1))
		Call EA_M_XML.AppInfo("TransferTotal",Setting(3))
		Call EA_M_XML.AppInfo("TitleLen",Setting(5))
		Call EA_M_XML.AppInfo("ContentLen",Setting(6))
		Call EA_M_XML.AppInfo("ShowColumn",Setting(7))
		Call EA_M_XML.AppInfo("ShowNew",Setting(8))
		Call EA_M_XML.AppInfo("ShowTime",Setting(9))
		Call EA_M_XML.AppInfo("ShowTypes",Setting(10))
		Call EA_M_XML.AppInfo("ShowReview",Setting(11))
		Call EA_M_XML.AppInfo("RowTotal",Setting(13))
		Call EA_M_XML.AppInfo("ImgSize",Setting(14))
	End If

	If Not IsArray(Setting) Then ReDim Setting(14):Setting(13)="1"

	Temp=EA_DBO.Get_Column_List()
	Tmp = "(build-select)," & Setting(0) & " " & str_InsideLink_All & ",0"
	If IsArray(Temp) Then
		ForTotal = UBound(Temp,2)

		For i=0 To ForTotal
			Level=(Len(Temp(2,i))/4-1)*3

			Tmp = Tmp & " ├" & String(Level,"-") & Replace(Temp(1,i)," ","") & "," & Temp(2,i)
		Next
	End If
	Call EA_M_XML.AppInfo("ColumnList",Tmp)

	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_Js_Help",str_Js_Help)
	Call EA_M_XML.AppElements("Language_Friend_Input_Info",str_Friend_Input_Info)

	Call EA_M_XML.AppElements("Language_Js_Title",str_Js_Title)
	Call EA_M_XML.AppElements("Language_Js_Detail",str_Js_Detail)
	Call EA_M_XML.AppElements("Language_Js_FilePath",str_Js_FilePath)
	Call EA_M_XML.AppElements("Language_Js_BaseSetting",str_Js_BaseSetting)
	Call EA_M_XML.AppElements("Language_Js_TransferSetting",str_Js_TransferSetting)
	Call EA_M_XML.AppElements("Language_Js_TransferColumn",str_Js_TransferColumn)
	Call EA_M_XML.AppElements("Language_Js_IncludeChildColumn",str_Js_IncludeChildColumn)
	Call EA_M_XML.AppElements("Language_Js_ListStyle",str_Js_ListStyle)
	Call EA_M_XML.AppElements("Language_Js_TransferTotal",str_Js_TransferTotal)
	Call EA_M_XML.AppElements("Language_Js_TransferType",str_Js_TransferType)
	Call EA_M_XML.AppElements("Language_Js_TitleLen",str_Js_TitleLen)
	Call EA_M_XML.AppElements("Language_Js_ContentLen",str_Js_ContentLen)
	Call EA_M_XML.AppElements("Language_Js_ShowFields",str_Js_ShowFields)
	Call EA_M_XML.AppElements("Language_Js_ShowFields_ColumnName",str_Js_ShowFields_ColumnName)
	Call EA_M_XML.AppElements("Language_Js_ShowFields_NewTag",str_Js_ShowFields_NewTag)
	Call EA_M_XML.AppElements("Language_Js_ShowFields_AddTime",str_Js_ShowFields_AddTime)
	Call EA_M_XML.AppElements("Language_Js_ShowFields_TypeTag",str_Js_ShowFields_TypeTag)
	Call EA_M_XML.AppElements("Language_Js_ShowFields_ReviewLink",str_Js_ShowFields_ReviewLink)
	Call EA_M_XML.AppElements("Language_Js_OpenWindowType",str_Js_OpenWindowType)
	Call EA_M_XML.AppElements("Language_Js_OpenWindowType_Parent",str_Js_OpenWindowType_Parent)
	Call EA_M_XML.AppElements("Language_Js_OpenWindowType_New",str_Js_OpenWindowType_New)
	Call EA_M_XML.AppElements("Language_Js_RowTotal",str_Js_RowTotal)
	Call EA_M_XML.AppElements("Language_Js_ImgSize",str_Js_ImgSize)

	Call EA_M_XML.AppElements("btnSubmit1",str_Comm_Save_Button)
	Call EA_M_XML.AppElements("btnSubmit2",str_Comm_Save_Button)
	Call EA_M_XML.AppElements("btnReturn1",str_Comm_Return_Button)
	Call EA_M_XML.AppElements("btnReturn2",str_Comm_Return_Button)

	Tmp = "(build-select)," & Setting(2) & " " & str_Js_ListStyle_Txt & ",0 " & str_Js_ListStyle_Detail & ",1 " & str_Js_ListStyle_Mix & ",2 " & str_Js_ListStyle_Img & ",3"
	Call EA_M_XML.AppInfo("Style",Tmp)

	Tmp = "(build-select)," & Setting(4) & " " & str_Js_Transfer_AllArticle & ",0 " & str_Js_Transfer_CommendArticle & ",1 " & str_Js_Transfer_HotArticle & ",2 " & str_Js_Transfer_ImgArticle & ",3"
	Call EA_M_XML.AppInfo("Type",Tmp)

	Tmp = "(build-select)," & Setting(12) & " " & str_Js_OpenWindowType_Parent & ",0 " & str_Js_OpenWindowType_New & ",1"
	Call EA_M_XML.AppInfo("OpenWindowType",Tmp)

	Page = EA_M_XML.make("","",0)
	Call EA_M_XML.Out(Page)
End Sub

Sub Preview
	Dim PostId
	Dim FilePath,Temp

	PostId=EA_Pub.SafeRequest(2,"ID",0,0,0)
	
	Temp=EA_M_DBO.Get_Js_Info(PostId)
	If IsArray(Temp) Then FilePath=Temp(2,0)
%>
<table width="90%">
  <tr> 
    <td><script language="JavaScript" src="../jsfiles/<%=FilePath%>"></script></td>
</table>
<%
End Sub

Sub Save
	Dim Title,Info,FileName,Setting
	Dim PostId
	
	PostId	= EA_Pub.SafeRequest(2,"ID",0,0,0)
	Title	= EA_Pub.SafeRequest(2,"Title",1,"",0)
	Info	= EA_Pub.SafeRequest(2,"Info",1,"",0)
	FileName= EA_Pub.SafeRequest(2,"FileName",1,"",0)
	
	Setting=EA_Pub.SafeRequest(2,"ColumnList",1,"",0)
	Setting=Setting&"|"&EA_Pub.SafeRequest(2,"IncludeChildColumn",0,0,0)
	Setting=Setting&"|"&EA_Pub.SafeRequest(2,"Style",0,0,0)
	Setting=Setting&"|"&EA_Pub.SafeRequest(2,"TransferTotal",0,10,0)
	Setting=Setting&"|"&EA_Pub.SafeRequest(2,"Type",0,0,0)
	Setting=Setting&"|"&EA_Pub.SafeRequest(2,"TitleLen",0,10,0)
	Setting=Setting&"|"&EA_Pub.SafeRequest(2,"ContentLen",0,50,0)
	Setting=Setting&"|"&EA_Pub.SafeRequest(2,"ShowColumn",0,0,0)
	Setting=Setting&"|"&EA_Pub.SafeRequest(2,"ShowNew",0,0,0)
	Setting=Setting&"|"&EA_Pub.SafeRequest(2,"ShowTime",0,0,0)
	Setting=Setting&"|"&EA_Pub.SafeRequest(2,"ShowTypes",0,0,0)
	Setting=Setting&"|"&EA_Pub.SafeRequest(2,"ShowReview",0,0,0)
	Setting=Setting&"|"&EA_Pub.SafeRequest(2,"OpenWindowType",0,0,0)
	Setting=Setting&"|"&EA_Pub.SafeRequest(2,"RowTotal",0,1,0)
	Setting=Setting&"|"&EA_Pub.SafeRequest(2,"ImgSize",1,"",0)
	
	If FoundErr Then
		Response.Write "-1"
	Else
		If Rs.State=1 Then rs.Close
		If PostId<>0 Then
			Sql="Select * From [NB_JsFile] Where [Id]="&PostId
			rs.Open Sql,Conn,2,2
		Else
			rs.Open "NB_JsFile",Conn,2,2
			rs.AddNew
		End If
			rs("Title")=Title
			rs("FileName")=FileName
			rs("Info")=Info
			rs("Setting")=Setting
			rs.update
		Rs.Close:Set Rs=Nothing
		
		Set Rs=Nothing
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

		EA_M_DBO.Set_Js_Delete Tmp
	Next
	
	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing
	
	Response.Write "1"
	Response.End
End Sub

Sub MakeJs()
	Dim NBArticle_Make,Content
	Dim TempArray,TSQL,TopicList,Temp,JsList,i
	Dim FileName,S
	Dim JsId(),IDs,ID
	
	Set NBArticle_Make=New Cls_MakeJs

	If Request("ID") = "" Then 
		SQL="Select FileName,Setting From [NB_JsFile]"
	Else
		IDs = Split(Request("ID"), ",")
		ForTotal = UBound(IDs)

		For i = 0 To ForTotal
			ReDim Preserve JsId(i)
			JsId(i) = EA_Pub.SafeRequest(5,IDs(i),0,0,0)
		Next

		ID = Join(JsId, ",")

		SQL="Select FileName,Setting From [NB_JsFile] Where [Id] IN (" & ID & ")"
	End If
	Set Rs=Conn.Execute(SQL)
	If Not rs.eof And Not rs.bof Then 
		JsList=Rs.GetRows()
		
		ID = 0
		ForTotal = UBound(JsList,2)

		For i=0 To ForTotal
			FileName=JsList(0,i)
			TempArray=Split(JsList(1,i),"|")
				
			TSQL=MakeSQLQuery(TempArray)
			
			If iDataBaseType=0 Then 
				SQL="Select top "&TempArray(3)&" [ID],COLUMNID,COLUMNNAME,TITLE,TCOLOR,AddDate,IsImg,IsTop,Img,left(Summary,"&TempArray(6)&") From [NB_Content] Where IsPass=-1 And IsDel=0"&TSQL
			Else
				SQL="Select top "&TempArray(3)&" [ID],COLUMNID,COLUMNNAME,TITLE,TCOLOR,AddDate,IsImg,IsTop,Img,substring(Summary,1,"&TempArray(6)&") From [NB_Content] Where IsPass=1 And IsDel=0"&TSQL
			End If
			'Response.Write sql
			Set Rs=Conn.Execute(SQL)
			If Not rs.eof And Not rs.bof Then 
				TopicList=Rs.GetRows(TempArray(3))
					
				Select Case TempArray(2)
				Case "0"
				'纯文本模式
					Content=NBArticle_Make.MakeTxtJs(TopicList,CInt(TempArray(7)),CInt(TempArray(9)),CInt(TempArray(8)),CInt(TempArray(10)),CInt(TempArray(11)),CInt(TempArray(5)),CInt(TempArray(6)),CInt(TempArray(12)),CInt(TempArray(13)))
				Case "1"
				'图文(上文+下左图+下右文)
					Temp=Split(TempArray(14),"&")
					
					Content=NBArticle_Make.MakeTxtMoreJs(TopicList,CInt(TempArray(7)),CInt(TempArray(9)),CInt(TempArray(8)),CInt(TempArray(10)),CInt(TempArray(11)),CInt(TempArray(5)),CInt(TempArray(12)),CInt(TempArray(13)),CInt(Temp(0)),CInt(Temp(1)))
				Case "2"
				'图文(上文+下左图+下右标题)
					Temp=Split(TempArray(14),"&")
					
					Content=NBArticle_Make.MakeGlsJs(TopicList,CInt(TempArray(7)),CInt(TempArray(9)),CInt(TempArray(8)),CInt(TempArray(10)),CInt(TempArray(11)),CInt(TempArray(5)),CInt(TempArray(6)),CInt(TempArray(12)),CInt(TempArray(13)),CInt(Temp(0)),CInt(Temp(1)))
				Case "3"
				'图片
					Temp=Split(TempArray(14),"&")
					
					Content=NBArticle_Make.MakeImgJs(TopicList,CInt(TempArray(7)),CInt(TempArray(9)),CInt(TempArray(8)),CInt(TempArray(10)),CInt(TempArray(11)),CInt(TempArray(5)),CInt(TempArray(12)),CInt(TempArray(13)),CInt(Temp(0)),CInt(Temp(1)))
				End Select
			Else
				Content="document.write ('<table><tr><td>·没有任何文章</td></tr></table>');"
			End If

			FileName="../jsfiles/" & FileName & ".js"
			Call EA_Pub.Save_HtmlFile(FileName,Content)

			ID = ID + 1
		Next
	End If
	
	Rs.Close
	Set Rs=Nothing
	
	Response.Write ID
	Response.End
End Sub

Function MakeSQLQuery(DataArray)
	Dim TempStr
	
	If DataArray(0)<>"0" Then 
		If DataArray(1)="1" Then 
			TempStr=" and columncode like '"&DataArray(0)&"%'"
		Else
			TempStr=" and columncode='"&DataArray(0)&"'"
		End If
	End If
	
	Select Case CInt(DataArray(4))
	Case 1
		TempStr=TempStr&" And istop="&EA_DBO.TrueValue
	Case 3
		TempStr=TempStr&" And isimg="&EA_DBO.TrueValue
	End Select
	
	Select Case DataArray(2)
	Case "0"
		If CInt(DataArray(4))=2 Then 
			TempStr=TempStr&" Order By ViewNum Desc,TrueTime Desc"
		Else
			TempStr=TempStr&" Order By TrueTime Desc"
		End If
	Case "1","2"
		If iDataBaseType=0 Then
			TempStr=TempStr&" Order By IsImg,TrueTime Desc"
		Else
			TempStr=TempStr&" Order By IsImg Desc,TrueTime Desc"
		End If
	Case "3"
		TempStr=TempStr&" Order By TrueTime Desc"
	End Select
	
	MakeSQLQuery=TempStr
End Function
%>