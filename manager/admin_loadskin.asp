<!--#Include File="comm/inc.asp" -->
<%
'====================================================================
'= Team Elite - EliteCMS
'= Copyright (c) 2004 - 2008 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：/Manager/Admin_LoadSkin.asp
'= 摘    要：后台-风格导入导出文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2007-07-09
'====================================================================

Call EA_Manager.Chk_IsMaster

If Not EA_Manager.Chk_Power(Admin_Power,"45") Then 
	ErrMsg=str_Comm_NotAccess
	Call EA_Manager.Error(1)
End If

Dim Atcion,TempConn
Atcion=Request.Form ("action")

Select Case Atcion
Case "load"
	Call Load()
Case "operation"
	Select Case Request.Form ("way")
	Case "0"
		OutputTemplate()
	Case "1"
		InputTemplate()
	End Select
Case Else
	Call Main()
End Select
Call EA_Pub.Close_Obj
Set EA_Pub=Nothing

Sub Main()
	Dim Operation
	Dim Temp,Page,i
	Dim ListName(2),ListValue()
	Dim mdbName

	If Atcion="loadthis" Then
		Operation=str_Skin_Input
	Else
		Operation=str_Skin_Output
	End If
	mdbName=Trim(Request.Form ("mdbName"))
	If mdbName = "" Then mdbName = "databackup/nb_template.mdb"

	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_Skin_Help",str_Skin_Help)

	Call EA_M_XML.AppElements("Language_Template_TempName",str_Template_TempName)
	Call EA_M_XML.AppElements("Language_Skin_Database",str_Skin_Database)

	Call EA_M_XML.AppElements("btnSubmit",Operation)

	Call EA_M_XML.AppInfo("mdbName",mdbName)

	SQL="Select [Id],Temp_Name From [NB_Template] Order By Id "

	If Atcion="loadthis" Then
		ConnectionSkinDataBase(mdbname)
		Set Rs=TempConn.Execute(SQL)
	Else
		Set Rs=Conn.Execute(SQL)
	End If

	ListName(0) = "checkbox"
	ListName(1) = "ID"
	ListName(2) = "Name"

	i = 0

	Do While Not Rs.eof
		ReDim Preserve ListValue(4,i)

		ListValue(0,i) = "checkbox"
		ListValue(1,i) = Rs("id")
		ListValue(2,i) = Rs("Temp_Name")

		i = i + 1

		Rs.MoveNext
	Loop

	Page = EA_M_XML.make(ListName,ListValue,i)

	Call EA_M_XML.Out(Page)
End Sub

Sub Load()
	Dim Page

	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_Skin_Help",str_Skin_Help)

	Call EA_M_XML.AppElements("Language_Skin_InputTemplate",str_Skin_InputTemplate)
	Call EA_M_XML.AppElements("Language_Skin_Database",str_Skin_Database)

	Call EA_M_XML.AppElements("btnSumbit",str_Comm_Next)

	Page = EA_M_XML.make("","",0)

	Call EA_M_XML.Out(Page)
End Sub

Sub OutputTemplate()
	On Error Resume Next

	Dim TempId,TmpArray
	Dim TmpRs
	Dim i
	
	TempId=EA_Pub.SafeRequest(2,"ID",1,"",0)
	
	ConnectionSkinDataBase(Request.Form ("dbname"))

	If TempId<>"" Then 
		Set TmpRs=Server.CreateObject("adodb.recordSet")
		SQL="Select * From [NB_Template]"
		TmpRs.Open SQL,TempConn,2,3
		
		SQL="Select * From [NB_Template] Where Id In ("&TempId&")"
		Set Rs=Conn.Execute(SQL)
		
		Do While Not rs.eof
			TmpRs.AddNew
			
			For i=1 To TmpRs.Fields.Count-1
				TmpRs(TmpRs(i).Name)=Rs(TmpRs(i).Name)
			Next
			
			TmpRs.Update 
			Rs.MoveNext
		Loop
		
		TempRs.Close:Set TempRs=Nothing
		TempConn.Close:Set TempConn=Nothing

		Call EA_Pub.Close_Obj
		Set EA_Pub=Nothing

		Response.Write "0"
		Response.End
	Else
		Response.Write "-1"
		Response.End
	End If
End Sub

Sub InputTemplate()
	On Error Resume Next
	
	Dim TempId,TmpArray
	Dim TmpRs
	Dim i
	
	TempId=EA_Pub.SafeRequest(2,"ID",1,"",0)
	
	ConnectionSkinDataBase(Request.Form ("dbname"))

	If TempId<>"" Then 
		Set Rs=Server.CreateObject("adodb.recordSet")
		SQL="Select * from [NB_Template]"
		Rs.Open SQL,Conn,2,3
		
		SQL="Select * from [NB_Template] Where Id In ("&TempId&")"
		Set TmpRs=TempConn.Execute(SQL)
		'Response.Write sql
		
		Do While Not TmpRs.eof
			Rs.AddNew
			
			For i=1 To Rs.Fields.Count-1
				'Response.Write i&Rs.Fields.Count&Rs(i).Name&"<br>"
				'Response.Flush 
				Rs(Rs(i).Name)=TmpRs(Rs(i).Name)
			Next
			
			Rs.Update 
			TmpRs.MoveNext
		Loop
		
		TempRs.Close:Set TempRs=Nothing
		TempConn.Close:Set TempConn=Nothing

		Call EA_Pub.Close_Obj
		Set EA_Pub=Nothing

		Response.Write "0"
		Response.End
	Else
		Response.Write "-1"
		Response.End
	End If
End Sub

Sub ConnectionSkinDataBase(mdbname)
	On Error Resume Next 
	Err.Clear 
	
	Set TempConn = Server.CreateObject("ADODB.Connection")
	TempConn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)
	
	If Err Then 
		Dim Page

		Page = EA_M_XML.make("","",0)

		Call EA_M_XML.Out(Page)
	End If
End Sub
%>