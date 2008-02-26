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
'= 文件名称：/Manager/Admin_Column.asp
'= 摘    要：后台-栏目管理文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-02-22
'====================================================================

Call EA_Manager.Chk_IsMaster

If Not EA_Manager.Chk_Power(Admin_Power,"11") Then 
	ErrMsg=str_Comm_NotAccess
	Call EA_Manager.Error(1)
End If

Dim Action
Dim ForTotal
Action=Request("action")

Select Case LCase(Action)
Case "save"
	Call Save
Case "del"
	Call del
Case "add"
	Call Add
Case "up"
	Call MoveColumn(1)
Case "down"
	Call MoveColumn(0)
Case "reset"
	Call ResetAllBoard()
Case Else
	Call Main
End Select
Call EA_Pub.Close_Obj
Set EA_Pub=Nothing

Sub Main
	Dim List
	Dim ListName(5),ListValue()
	Dim Level,i,ColumnRetract

	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_Column_Help",str_Column_Help)

	Call EA_M_XML.AppElements("Language_Column_Title",str_Column_Title)
	Call EA_M_XML.AppElements("Language_Column_ManagerArticleTotal",str_Column_ManagerArticleTotal)
	Call EA_M_XML.AppElements("Language_Column_ArticleTotal",str_Column_ArticleTotal)
	Call EA_M_XML.AppElements("Language_Column_ResetBoard",str_Column_ResetBoard)

	Call EA_M_XML.AppElements("Comm_Add_Operation",str_Comm_Add_Operation)
	Call EA_M_XML.AppElements("Comm_Del_Operation",str_Comm_Del_Operation)

	Call EA_M_XML.AppElements("Language_Comm_Bar_Operation",str_Comm_Bar_Operation)

	List=EA_DBO.Get_Column_List()
    If IsArray(List) Then 
		ListName(0) = "checkbox"
		ListName(1) = "ID"
		ListName(2) = "T_ID"
		ListName(3) = "Title"
		ListName(4) = "Stat"
		ListName(5) = "action"
		ForTotal = UBound(List,2)

		For i=0 To ForTotal
			ReDim Preserve ListValue(5,i)
			ColumnRetract = ""

			Level=(Len(List(2,i))/4-1)*4
			If Len(List(2,i))>4 Then ColumnRetract = ColumnRetract & "├"
			ColumnRetract = ColumnRetract & String(Level,"-")

			ListValue(0,i) = "checkbox"
			ListValue(1,i) = List(0,i)
			ListValue(2,i) = List(0,i)
			ListValue(3,i) = ColumnRetract & "<a href='" & EA_Pub.Cov_ColumnPath(List(0,i),EA_Pub.SysInfo(18)) & "' target='_blank'>" & List(1,i)  & "</a>"
			If List(5,i)>0 Then
				ListValue(4,i) = "<font color=red><strong>" & List(5,i) & "</strong></font>"
			Else
				ListValue(4,i) = List(5,i)
			End If
			ListValue(4,i) = ListValue(4,i) & "/" & List(4,i)
			ListValue(5,i) = "action"
		Next
		
		Page = EA_M_XML.make(ListName,ListValue,UBound(List,2)+1)
	Else
		Page = EA_M_XML.make("","",0)
	End If

	Call EA_M_XML.Out(Page)
End Sub

Sub Add
	Dim PostId
	Dim i,Level
	Dim Temp,Tmp
	Dim Code,Style,List_TempId,Article_TempId
	
	PostId=EA_Pub.SafeRequest(2,"ID",0,0,0)
	Call EA_M_XML.AppInfo("ID",PostId)

	List_TempId = 0
	Article_TempId = 0
	Code = 0
	
	Temp=EA_DBO.Get_Column_Info(PostId)
	If IsArray(Temp) Then
		Call EA_M_XML.AppInfo("Name",Temp(0,0))
		Call EA_M_XML.AppInfo("Info",Temp(2,0))
		Call EA_M_XML.AppInfo("OutUrl",Temp(7,0))
		Call EA_M_XML.AppInfo("Power",Temp(12,0))
		Call EA_M_XML.AppInfo("IsHide",Abs(CInt(Temp(13,0))))
		Call EA_M_XML.AppInfo("IsReview",Abs(CInt(Temp(14,0))))
		Call EA_M_XML.AppInfo("IsPost",Abs(CInt(Temp(15,0))))
		Call EA_M_XML.AppInfo("IsTop",Abs(CInt(Temp(16,0))))
		Call EA_M_XML.AppInfo("PageSize",Temp(17,0))

		Code			= Left(Temp(1,0),Len(Temp(1,0))-4)
		Style			= Temp(8,0)
		List_TempId		= Temp(9,0)
		Article_TempId	= Temp(10,0)
	Else
		Call EA_M_XML.AppInfo("Power","0")
		Call EA_M_XML.AppInfo("PageSize","10")
	End If

	Temp=EA_DBO.Get_Column_List()
	Tmp = "(build-select)," & Code & " " & str_Column_Root & ",0"
	If IsArray(Temp) Then
		ForTotal = UBound(Temp,2)

		For i=0 To ForTotal
			Level=(Len(Temp(2,i))/4-1)*3

			Tmp = Tmp & " ├" & String(Level,"-") & Replace(Temp(1,i)," ","") & "," & Temp(2,i)
		Next
	End If
	Call EA_M_XML.AppInfo("ColumnList",Tmp)


	Temp=EA_M_DBO.Get_DefaultModule_List()
	Tmp = "(build-select)," & List_TempId & " " & str_Comm_Select & ",0"
	If IsArray(Temp) Then
		ForTotal = UBound(Temp,2)

		For i=0 To ForTotal
			Tmp = Tmp & " " & Replace(Temp(1,i)," ","") & "," & Temp(0,i)
		Next
	End If
	Call EA_M_XML.AppInfo("List_TempId",Tmp)


	Temp=EA_M_DBO.Get_DefaultModule_List()
	Tmp = "(build-select)," & Article_TempId & " " & str_Comm_Select & ",0"
	If IsArray(Temp) Then
		ForTotal = UBound(Temp,2)

		For i=0 To ForTotal
			Tmp = Tmp & " " & Replace(Temp(1,i)," ","") & "," & Temp(0,i)
		Next
	End If
	Call EA_M_XML.AppInfo("Article_TempId",Tmp)


	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_Column_Help",str_Column_Help)
	Call EA_M_XML.AppElements("Language_Column_Input_Column",str_Column_Input_Column)

	Call EA_M_XML.AppElements("Language_Column_Title",str_Column_Title)
	Call EA_M_XML.AppElements("Language_Column_Attrib",str_Column_Attrib)
	Call EA_M_XML.AppElements("Language_Column_Info",str_Column_Info)
	Call EA_M_XML.AppElements("Language_Column_OutURL",str_Column_OutURL)
	Call EA_M_XML.AppElements("Language_Column_ViewPower",str_Column_ViewPower)
	Call EA_M_XML.AppElements("Language_Column_IsHide",str_Column_IsHide)
	Call EA_M_XML.AppElements("Language_Column_IsReview",str_Column_IsReview)
	Call EA_M_XML.AppElements("Language_Column_IsPost",str_Column_IsPost)
	Call EA_M_XML.AppElements("Language_Column_IsTop",str_Column_IsTop)
	Call EA_M_XML.AppElements("Language_Column_Style",str_Column_Style)
	Call EA_M_XML.AppElements("Language_Column_ListTemplate",str_Column_ListTemplate)
	Call EA_M_XML.AppElements("Language_Column_ArticleTemplate",str_Column_ArticleTemplate)
	Call EA_M_XML.AppElements("Language_Column_PageSize",str_Column_PageSize)

	Call EA_M_XML.AppElements("btnSubmit",str_Comm_Save_Button)
	Call EA_M_XML.AppElements("btnReturn",str_Comm_Return_Button)

	For i= 1 To 4
		Call EA_M_XML.AppElements("Language_Comm_Yes" & i,str_Comm_Yes)
		Call EA_M_XML.AppElements("Language_Comm_No" & i,str_Comm_No)
	Next

	Page = EA_M_XML.make("","",0)
	Call EA_M_XML.Out(Page)
End Sub

Sub ResetAllBoard()
	Dim i,Temp
	i=1
	
	Temp=EA_DBO.Get_Column_List()
	
	If IsArray(Temp) Then
		ForTotal = UBound(Temp,2)

		For i=0 To ForTotal
			SQL="UpDate [NB_Column] Set Code='"&Right("0000"&(i+1),4)&"' Where [Id]="&Temp(0,i)
			EA_M_DBO.DB_Execute(SQL)
		Next
	End If
	
	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing
	
	Response.Write "1"
	Response.End
End Sub

Sub MoveColumn(IsUp)
	On Error Resume Next
	Err.Clear

	Dim PostId
	Dim Temp
	Dim ColumnCode,TempStr,CodeLen,NBStr,WStr

	PostId=EA_Pub.SafeRequest(2,"ID",0,0,0)
	
	Temp=EA_DBO.Get_Column_Info(PostId)
	If IsArray(Temp) Then 
		ColumnCode=Temp(1,0)
		CodeLen=Len(ColumnCode)
		NBStr=String(CodeLen,"-")
	End If
	
	If ColumnCode<>"" And (Len(ColumnCode) Mod 4)=0 Then 
		If CodeLen>4 Then WStr="and left(code,"&CodeLen-4&")='"&Left(ColumnCode,CodeLen-4)&"'"
		If IsUp Then 
			SQL="select top 1 code from [NB_Column] where len(code)="&CodeLen&" and code<'"&ColumnCode&"' "&WStr&" order by code desc"
			TempStr=EA_M_DBO.DB_Execute(sql)(0)
		Else
			SQL="select top 1 code from [NB_Column] where len(code)="&CodeLen&" and code>'"&ColumnCode&"' "&WStr&" order by code"
			TempStr=EA_M_DBO.DB_Execute(sql)(0)
		End If

		If Err Then 
			Response.Write "-1"
			Response.End
		End If
			
		'Move Under Column
		If iDataBaseType = 0 Then
			SQL="update nb_column set code='"&NBStr&"'+mid(code,"&CodeLen+1&",len(code)) where left(code,"&CodeLen&")='"&TempStr&"'"
		Else
			SQL="update nb_column set code='"&NBStr&"'+SUBSTRING(code,"&CodeLen+1&",len(code)) where left(code,"&CodeLen&")='"&TempStr&"'"
		End If
		EA_M_DBO.DB_Execute(SQL)

		'Update Target Column
		If iDataBaseType = 0 Then
			SQL="update nb_column set code='"&TempStr&"'+mid(code,"&CodeLen+1&",len(code)) where left(code,"&CodeLen&")='"&ColumnCode&"'"
		Else
			SQL="update nb_column set code='"&TempStr&"'+SUBSTRING(code,"&CodeLen+1&",len(code)) where left(code,"&CodeLen&")='"&ColumnCode&"'"
		End If
		EA_M_DBO.DB_Execute(SQL)

		'Update Under Column
		If iDataBaseType = 0 Then
			SQL="update nb_column set code='"&ColumnCode&"'+mid(code,"&CodeLen+1&",len(code)) where left(code,"&CodeLen&")='"&NBStr&"'"
		Else
			SQL="update nb_column set code='"&ColumnCode&"'+SUBSTRING(code,"&CodeLen+1&",len(code)) where left(code,"&CodeLen&")='"&NBStr&"'"
		End If
		EA_M_DBO.DB_Execute(SQL)

		If iDataBaseType<>2 Then
			'Update Table Of Article's Column Information
			SQL="update [NB_Content] left join [NB_Column] c on [NB_Content].columnid=c.id set columncode=c.code,columnname=c.title where columncode like '"&ColumnCode&"%'"
			EA_M_DBO.DB_Execute(SQL)
		End If
	End If
	
	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing
	
	Response.Write "1"
	Response.End
End Sub

Sub Save
	Dim Name,Power,Info,IsHide,IsOut,OutUrl,Style,IsReview,IsPost,IsTop,List_TempId,Article_TempId,PageSize
	Dim MatchStr,ParentCode,TypeCode,SelfCode,StepLeng
	Dim EditCode,SourCode
	Dim PostId

	FoundErr = False
	EditCode	= False
	StepLeng	= 4
	MatchStr	= String(StepLeng,"____")
	
	PostId			= EA_Pub.SafeRequest(2,"ID",0,0,0)
	Name			= EA_Pub.SafeRequest(2,"Name",1,"",1)
	Power			= EA_Pub.SafeRequest(2,"Power",0,0,0)
	Info			= EA_Pub.SafeRequest(2,"Info",1,"",0)
	IsHide			= EA_Pub.SafeRequest(2,"IsHide",0,0,0)
	OutUrl			= EA_Pub.SafeRequest(2,"OutUrl",1,"",0)
	ParentCode		= EA_Pub.SafeRequest(2,"ColumnList",1,"",0)
	Style			= EA_Pub.SafeRequest(2,"Style",0,0,0)
	IsReview		= EA_Pub.SafeRequest(2,"IsReview",0,1,0)
	IsPost			= EA_Pub.SafeRequest(2,"IsPost",0,1,0)
	IsTop			= EA_Pub.SafeRequest(2,"IsTop",0,1,0)
	List_TempId		= EA_Pub.SafeRequest(2,"List_TempId",0,1,0)
	Article_TempId	= EA_Pub.SafeRequest(2,"Article_TempId",0,1,0)
	PageSize		= EA_Pub.SafeRequest(2,"PageSize",0,15,0)
	
	If ParentCode = "0" Then ParentCode = ""
	ParentCode		= CStr(ParentCode)
	
	If Name="" Or Len(Name)>50 Then Response.Write "-1":Response.End
	
	If Len(OutUrl)>0 Then
		IsOut=1
	Else
		IsOut=0
	End If
		
	If PostId<>0 Then
		SourCode=EA_DBO.Get_Column_Info(PostId)(1,0)
		SourCode=CStr(SourCode)
		If ParentCode<>CStr(Left(SourCode,Len(SourCode)-4)) Then 
			If ParentCode<>SourCode Then 
				If Left(ParentCode,Len(SourCode))<>SourCode Then EditCode=True
			End If
		End If
	Else
		EditCode=True
	End If
	
	If EditCode Then 
		Sql="Select Top 1 Code from [NB_Column] Where Code Like '"&ParentCode&MatchStr&"' Order By Code Desc"
		Set Rs=Conn.Execute(SQL)
		If rs.eof Then
			TypeCode=ParentCode&String(StepLeng-1,"0")&"1"
		Else
			SelfCode=Int(Right(rs(0),StepLeng))+1
			SelfCode=Right(String(StepLeng-1,"0")&SelfCode,StepLeng)
			TypeCode=ParentCode&SelfCode
		End If
		Rs.Close
	End If
	
	If PostId=0 Then 
		EA_DBO.Set_System_ColumnTotal 1
			
		Sql="INSERT INTO NB_Column ( Title, Code, Info, IsOut, OutUrl, StyleId, IsReview, IsPost, IsTop, List_TempId, Article_TempId, PageSize, ListPower, IsHide )"
		Sql=Sql&" VALUES ( '"&Name&"','"&TypeCode&"','"&Info&"',"&IsOut&",'"&OutUrl&"',"&Style&","&IsReview&","&IsPost&","&IsTop&","&List_TempId&","&Article_TempId&","&PageSize&","&Power&","&IsHide&")"
	ElseIf IsNumeric(PostId) And PostId<>"" And PostId<>"0" Then
		Sql="Update [NB_Column] Set Title='"&Name&"',Info='"&Info&"',IsOut="&IsOut&",OutUrl='"&OutUrl&"',StyleId="&Style&",IsReview="&IsReview&",IsPost="&IsPost&",IsTop="&IsTop&",List_TempId="&List_TempId&",Article_TempId="&Article_TempId&",PageSize="&PageSize&",ListPower="&Power&",IsHide="&IsHide
		If EditCode Then Sql=Sql&",Code='"&TypeCode&"'"
		Sql=Sql&" Where Id="&PostId
	End If
	EA_M_DBO.DB_Execute(SQL)

	If EditCode And PostId<>0 Then 
		SQL="update [NB_Column] Set Code='"&TypeCode&"'+Right(Code,Len(Code)-"&Len(SourCode)&") Where Code Like '"&SourCode&"%'"
		EA_M_DBO.DB_Execute(SQL)
	End If
		
	Set Rs=Nothing
	
	Response.Write "1"
	Response.End
End Sub

Sub Del
	Dim PostId
	Dim DelBool,TxtCount,Under

	PostId=EA_Pub.SafeRequest(2,"ID",0,0,0)
	DelBool=False
		
	TxtCount=EA_M_DBO.Get_Column_ArticleTotal(PostId)(0,0)
	If TxtCount>0 Then DelBool=True:ErrMsg=str_Column_ColumnIsNotEmpty
		
	Sql="Select Count(Id) From [NB_Column] Where Code Like (Select Code From [NB_Column] Where Id="&PostId&")+'%' And Id<>"&PostId
	Under=EA_M_DBO.DB_Execute(SQL)(0)
	If Under>0 Then DelBool=True:ErrMsg=str_Column_ColumnHaveUnder
		
	If Not DelBool Then 
		EA_M_DBO.Set_Column_Delete PostId
		EA_DBO.Set_System_ColumnTotal -1

		Application.Lock 
		Application(sCacheName&"IsFlush")=1
		Application.UnLock 

		Response.Write "0"
	Else
		Response.Write "-1"
	End If

	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing
	
	Response.End
End Sub
%>