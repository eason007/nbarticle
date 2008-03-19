<!--#Include File="comm/inc.asp" -->
<!--#Include File="../include/page_column.asp"-->
<%
'====================================================================
'= Team Elite - EliteCMS
'= Copyright (c) 2004 - 2008 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：/Manager/Admin_MakeList.asp
'= 摘    要：后台-HTML栏目页生成文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-03-14
'====================================================================

Server.ScriptTimeout = 999

Call EA_Manager.Chk_IsMaster

If Not EA_Manager.Chk_Power(Admin_Power,"42") Then 
	ErrMsg=str_Comm_NotAccess
	Call EA_Manager.Error(1)
End If

Dim Atcion
Dim ForTotal

Atcion=Request.Form ("action")

Select Case LCase(Atcion)
Case "mark"
	Call MarkList
Case Else
	Call Main
End Select
Call EA_Pub.Close_Obj
Set EA_Pub=Nothing

Sub Main
	Dim Level,Temp,i
	Dim ColumnList

	Temp=EA_DBO.Get_Column_List()
	ColumnList = str_Comm_AllColumn & ",0"
	If IsArray(Temp) Then
		ForTotal = UBound(Temp,2)

		For i=0 To ForTotal
			Level=(Len(Temp(2,i))/4-1)*2

			ColumnList = ColumnList & " ├"
			ColumnList = ColumnList & String(Level,"-")
			ColumnList = ColumnList & Replace(Temp(1,i)," ","") & "," & Temp(0,i)
		Next
	End If
	Call EA_M_XML.AppInfo("ColumnList",ColumnList)
	Call EA_M_XML.AppInfo("total",UBound(Temp, 2) + 1)

	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_MakeList_Help",str_MakeList_Help)

	Call EA_M_XML.AppElements("Language_Comm_SelectAll",str_Comm_SelectAll)
	Call EA_M_XML.AppElements("Language_Comm_Select",str_Comm_Select)

	Call EA_M_XML.AppElements("Language_MakeList_Title",str_MakeList_Title)

	Call EA_M_XML.AppElements("btnSubmit",str_MakeIndex_StartNow)

	Page = EA_M_XML.make("","",0)
	Call EA_M_XML.Out(Page)
End Sub

Sub MarkList
	Dim ColumnList, PostId, Total
	Dim Tmp, i
	
	PostId = EA_Pub.SafeRequest(2,"ColumnList",1,"",1)
	Total = EA_Pub.SafeRequest(2,"total",0,0,0)
	ColumnList = Split(PostId, ", ")
	If Ubound(ColumnList) = -1 Then Exit Sub
	If PostId = "0" Then 
		Tmp = EA_DBO.Get_Column_List()

		Total = UBound(Tmp,2)

		ReDim ColumnList(Total)

		For i = 0 To Total
			ColumnList(i) = Tmp(0,i)
		Next
	End If
	If (Ubound(ColumnList) = 0 And PostId <> "0") Or (Ubound(ColumnList) > 0) Then Total = UBound(ColumnList) + 1
	
	PageContent=EA_Temp.Load_Template_File("admin_makelist_view.htm")
	EA_Temp.P_Prefix = "{$"
	EA_Temp.P_Suffix = "$}"

	EA_Temp.SetVariable "ColumnTotal", Total, PageContent
	EA_Temp.SetVariable "Language_MakeList_Column", str_MakeList_Column, PageContent
	EA_Temp.SetVariable "Language_MakeList_Now", str_MakeList_Now, PageContent
	EA_Temp.SetVariable "Language_MakeList_Task", str_MakeList_Task, PageContent
	EA_Temp.SetVariable "Language_MakeList_Page", str_MakeList_Page, PageContent

	Response.Write PageContent

	For i = 0 To Total - 1
		Response.Write "<script>page_complete.innerHTML=""1"";</script>" & VbCrLf
		Response.Write "<script>img3.width=1;</script>" & VbCrLf
		
		Call ForColumn(ColumnList(i))
		
		If i = Total - 1 Then
			Response.Write "<script>img1.width=400;" & VbCrLf
		Else
			Response.Write "<script>img1.width=" & Fix(((i + 1)/Total) * 400) & ";" & VbCrLf
		End If
		Response.Write "column_complete.innerHTML=""<font color=blue>"&i+1&"</font>"";</script>" & VbCrLf
		Response.Flush
	Next
	
	Response.Write "<script>make_msg.innerHTML="""&str_MakeList_AllComplate&""";</script>" & VbCrLf
End Sub

Sub ForColumn(ColumnId)
	Dim PageContent
	Dim PageCount, PageSize
	Dim Tmp, i, TempStr
	Dim ColumnInfo
	Dim FileName, Folder
	Dim re
	Dim clsColumn

	Set re = New RegExp

	re.IgnoreCase= True
	re.Global	 = True
	re.Pattern	 = Replace(SystemFolder, "/", "\/") & "(.*)\/(\w+).(\w+)"

	FileName = EA_Pub.Cov_ColumnPath(ColumnId, "0")
	Folder	 = re.Replace(FileName, "/$1/")

	Set re = Nothing

	If Not(EA_Pub.CheckDir(".." & Folder)) Then 
		Tmp		 = Split(Folder, "/")
		TempStr  = ""
		ForTotal = UBound(Tmp) - 1

		For i = 1 To ForTotal
			TempStr = TempStr & "/" & Tmp(i)
			
			If Not(EA_Pub.CheckDir(".." & TempStr)) Then EA_Pub.MakeNewsDir Server.MapPath(".." & TempStr)
		Next
	End If
	
	'load column data
	ColumnInfo	= EA_DBO.Get_Column_Info(ColumnId)

	If ColumnInfo(6,0) Then 
		PageContent = "<meta http-equiv=""refresh"" content=""0;URL=" & ColumnInfo(7, 0) & """>"
		
		Call EA_Pub.Save_HtmlFile(FileName, PageContent)
	Else
		Set clsColumn = New page_Column

		EA_Temp.P_Prefix = "<!--"
		EA_Temp.P_Suffix = "-->"

		PageSize	= ColumnInfo(17, 0)
		PageCount	= EA_Pub.Stat_Page_Total(PageSize, ColumnInfo(3, 0))
		If PageCount = 0 Then PageCount = 1
		Response.Write "<script>page_total.innerHTML=""" & PageCount & """;</script>" & VbCrLf
		
		EA_Pub.SysInfo(18) = "0"

		For i = 1 To PageCount
			Call clsColumn.Make(ColumnId, ColumnInfo, i)

			Call EA_Pub.Save_HtmlFile(Replace(FileName, "_1", "_" & i), clsColumn.PageContent)

			Response.Write "<script>img3.width=" & Fix((i/PageCount) * 100) & ";" & VbCrLf
			Response.Write "page_complete.innerHTML=""<b>" & i & "</b>"";</script>" & VbCrLf
			Response.Flush
		Next
	End If

	Response.Write "<script>img3.width=400;</script>" & VbCrLf
End Sub
%>