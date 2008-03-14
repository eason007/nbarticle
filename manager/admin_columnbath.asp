<!--#Include File="comm/inc.asp" -->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2004 - 2008 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：/Manager/Admin_ColumnBath.asp
'= 摘    要：后台-文章批量管理文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2006-07-27
'====================================================================

Call EA_Manager.Chk_IsMaster

If Not EA_Manager.Chk_Power(Admin_Power,"13") Then 
	ErrMsg=str_Comm_NotAccess
	Call EA_Manager.Error(1)
End If

Dim Action
Action=Request.Form("action")

Select Case LCase(Action)
Case "move"
	Call Move
Case "bathmove"
	Call BathMove
Case "del"
	Call del
Case "bathdel"
	Call BathDel
End Select
Call EA_Pub.Close_Obj
Set EA_Pub=Nothing
	
Sub Move
	Dim Level,Temp,i,Tmp
	Dim ColumnList
	Dim ForTotal

	Temp=EA_DBO.Get_Column_List()
	Tmp = "(build-select),0 " & str_Comm_Select & ",0"
	If IsArray(Temp) Then 
		ForTotal = UBound(Temp,2)

		For i=0 To ForTotal
			Level=(Len(Temp(2,i))/4-1)*2
			
			Tmp = Tmp & " ├" & String(Level,"-") & Replace(Temp(1,i)," ","") & "," & Temp(0,i)
		Next
	End If
	Call EA_M_XML.AppInfo("SourColumn",Tmp)
	Call EA_M_XML.AppInfo("DestColumn",Tmp)
	
	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_Bath_Move_Help",str_Bath_Move_Help)

	Call EA_M_XML.AppElements("Language_Bath_BathMove",str_Bath_BathMove)
	Call EA_M_XML.AppElements("Language_Bath_BathDel",str_Bath_BathDel)

	Call EA_M_XML.AppElements("Language_Bath_From",str_Bath_From)
	Call EA_M_XML.AppElements("Language_Bath_To",str_Bath_To)
	Call EA_M_XML.AppElements("Language_Bath_Condition",str_Bath_Condition)
	Call EA_M_XML.AppElements("Language_Bath_All",str_Bath_All)
	Call EA_M_XML.AppElements("Language_Bath_ByDate",str_Bath_ByDate)
	Call EA_M_XML.AppElements("Language_Bath_ByKeyword",str_Bath_ByKeyword)
	Call EA_M_XML.AppElements("Language_Bath_In",str_Bath_In)

	Call EA_M_XML.AppElements("btnSubmit",str_Comm_Submit_Button)

	Tmp = "(build-select), " & str_Content_Title & ",0 " & str_Content_Keyword & ",1 " & str_Content_Author & ",2 " & str_Content_Summary & ",3 " & str_Content_Source & ",4 " & str_Content_Content & ",5"
	Call EA_M_XML.AppInfo("Field",Tmp)

	Page = EA_M_XML.make("","",0)
	Call EA_M_XML.Out(Page)
End Sub

Sub Del
	Dim Level,Temp,i
	Dim Tmp
	Dim ForTotal

	Temp=EA_DBO.Get_Column_List()
	Tmp = "(build-select),0 " & str_Comm_Select & ",0"
	If IsArray(Temp) Then 
		ForTotal = UBound(Temp,2)

		For i=0 To ForTotal
			Level=(Len(Temp(2,i))/4-1)*2
			
			Tmp = Tmp & " ├" & String(Level,"-") & Replace(Temp(1,i)," ","") & "," & Temp(0,i)
		Next
	End If
	Call EA_M_XML.AppInfo("Column",Tmp)

	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_Bath_Del_Help",str_Bath_Del_Help)

	Call EA_M_XML.AppElements("Language_Bath_BathMove",str_Bath_BathMove)
	Call EA_M_XML.AppElements("Language_Bath_BathDel",str_Bath_BathDel)

	Call EA_M_XML.AppElements("Language_Bath_DestColumn",str_Bath_DestColumn)
	Call EA_M_XML.AppElements("Language_Bath_Condition",str_Bath_Condition)
	Call EA_M_XML.AppElements("Language_Bath_All",str_Bath_All)
	Call EA_M_XML.AppElements("Language_Bath_ByDate",str_Bath_ByDate)
	Call EA_M_XML.AppElements("Language_Bath_ByKeyword",str_Bath_ByKeyword)
	Call EA_M_XML.AppElements("Language_Bath_In",str_Bath_In)
	Call EA_M_XML.AppElements("Language_Bath_Option",str_Bath_Option)
	Call EA_M_XML.AppElements("Language_Bath_RecycleBin",str_Bath_RecycleBin)
	Call EA_M_XML.AppElements("Language_Bath_NoRecycleBin",str_Bath_NoRecycleBin)

	Call EA_M_XML.AppElements("btnSubmit",str_Comm_Submit_Button)

	Tmp = "(build-select), " & str_Content_Title & ",0 " & str_Content_Keyword & ",1 " & str_Content_Author & ",2 " & str_Content_Summary & ",3 " & str_Content_Source & ",4 " & str_Content_Content & ",5"
	Call EA_M_XML.AppInfo("Field",Tmp)

	Page = EA_M_XML.make("","",0)
	Call EA_M_XML.Out(Page)
End Sub

Sub BathMove
	Dim Where,Sour,Dest,Count
	Dim DestColumn_Name,DestColumn_Code
	Dim Temp
	
	Count	= 0
	Where	= EA_Pub.SafeRequest(2,"where",0,-1,0)
	Sour	= EA_Pub.SafeRequest(2,"SourColumn",0,0,0)
	Dest	= EA_Pub.SafeRequest(2,"DestColumn",0,0,0)
	
	If Sour<>0 And Dest<>0 Then
		Temp=EA_DBO.Get_Column_Info(Dest)
		If IsArray(Temp) Then 
			DestColumn_Name=Temp(0,0)
			DestColumn_Code=Temp(1,0)
			
			Select Case Where
			Case 0
				Sql="Select Count([Id]) From [NB_Content] Where ColumnId="&Sour
				Count=EA_M_DBO.DB_Query(SQL)(0, 0)
					
				Sql="UpDate [NB_Column] Set CountNum=0 Where [Id]="&Sour
				EA_M_DBO.DB_Execute(SQL)
				
				Sql="UpDate [NB_Column] Set CountNum=CountNum+"&Count&" Where [Id]="&Dest
				EA_M_DBO.DB_Execute(SQL)
				
				Sql="UpDate [NB_Content] Set ColumnId="&Dest&",ColumnName='"&DestColumn_Name&"',ColumnCode='"&DestColumn_Code&"' Where ColumnId="&Sour
				EA_M_DBO.DB_Execute(SQL)
			Case 1
				Dim Time
				Time=EA_Pub.SafeRequest(2,"date",2,"1900",0)
				
				Sql="Select Count([Id]) From [NB_Content] Where DateDiff('d',AddDate,'"&Time&"')=0 And ColumnId="&Sour
				Count=EA_M_DBO.DB_Query(SQL)(0, 0)
				
				Sql="UpDate [NB_Column] Set CountNum=CountNum-"&Count&" Where [Id]="&Sour
				EA_M_DBO.DB_Execute(SQL)
				
				Sql="UpDate [NB_Column] Set CountNum=CountNum+"&Count&" Where [Id]="&Dest
				EA_M_DBO.DB_Execute(SQL)
				
				Sql="UpDate [NB_Content] Set ColumnId="&Dest&",ColumnName='"&DestColumn_Name&"',ColumnCode='"&DestColumn_Code&"' Where DateDiff('d',AddDate,'"&Time&"')=0 And ColumnId="&Sour
				EA_M_DBO.DB_Execute(SQL)
			Case 2
				Dim KeyWord,Field,WSQL
				KeyWord=EA_Pub.SafeRequest(2,"keyword",1,"",0)
				Field=EA_Pub.SafeRequest(2,"Field",0,-1,0)
				
				If Field<>-1 Then
					Select Case Field
					Case 0
						WSQL=" Where ColumnId="&Sour&" And InStr(1,Title,'"&KeyWord&"')>0"
					Case 1
						WSQL=" Where ColumnId="&Sour&" And InStr(1,Keyword,'"&KeyWord&"')>0"
					Case 2
						WSQL=" Where ColumnId="&Sour&" And InStr(1,Author,'"&KeyWord&"')>0"
					Case 3
						WSQL=" Where ColumnId="&Sour&" And InStr(1,Summary,'"&KeyWord&"')>0"
					Case 4
						WSQL=" Where ColumnId="&Sour&" And InStr(1,Source,'"&KeyWord&"')>0"
					Case 5
						WSQL=" Where ColumnId="&Sour&" And InStr(1,Content,'"&KeyWord&"')>0"
					End Select
				
					Sql="Select Count([Id]) From [NB_Content]"&WSQL
					'Response.Write sql
					Count=EA_M_DBO.DB_Query(SQL)(0, 0)
					
					Sql="UpDate [NB_Column] Set CountNum=CountNum-"&Count&" Where [Id]="&Sour
					EA_M_DBO.DB_Execute(SQL)
				
					Sql="UpDate [NB_Column] Set CountNum=CountNum+"&Count&" Where [Id]="&Dest
					EA_M_DBO.DB_Execute(SQL)
				
					Sql="UpDate [NB_Content] Set ColumnId="&Dest&",ColumnName='"&DestColumn_Name&"',ColumnCode='"&DestColumn_Code&"' "&WSQL
					'Response.Write sql
					EA_M_DBO.DB_Execute(SQL)
				End If
			End Select
		End If
	End If
	
	ErrMsg=Replace(str_Bath_MoveMsg,"$1",Count)
	Response.Write ErrMsg
End Sub

Sub BathDel
	Dim Where,Sour,Count,WSQL,Back
	Dim DestColumn_Name,DestColumn_Code
	
	Count	= 0
	Where	= EA_Pub.SafeRequest(2,"where",0,-1,0)
	Sour	= EA_Pub.SafeRequest(2,"Column",0,0,0)
	Back	= EA_Pub.SafeRequest(2,"back",0,0,0)
	
	If Sour<>0 Then
		Select Case Where
		Case 0
			WSQL=" Where [ColumnId]="&Sour
			
			Sql="Select Count([Id]) From [NB_Content]"&WSQL
			Count=EA_M_DBO.DB_Query(SQL)(0, 0)
				
			Sql="UpDate [NB_Column] Set CountNum=0 Where [Id]="&Sour
			EA_M_DBO.DB_Execute(SQL)
		Case 1
			Dim Time
			Time=EA_Pub.SafeRequest(2,"date",2,"1900",0)
			
			WSQL=" where datediff('d',adddate,'"&Time&"')=0 And ColumnId="&Sour
				
			Sql="Select Count([Id]) From [NB_Content]"&WSQL
			Count=EA_M_DBO.DB_Query(SQL)(0, 0)
				
			Sql="UpDate [NB_Column] Set CountNum=CountNum-"&Count&" Where [Id]="&Sour
			EA_M_DBO.DB_Execute(SQL)
		Case 2
			Dim KeyWord,Field
			KeyWord=EA_Pub.SafeRequest(2,"keyword",1,"",0)
			Field=EA_Pub.SafeRequest(2,"Field",0,-1,0)
			
			If Field<>-1 Then
				Select Case Field
				Case 0
					WSQL=" Where ColumnId="&Sour&" And InStr(1,Title,'"&KeyWord&"')>0"
				Case 1
					WSQL=" Where ColumnId="&Sour&" And InStr(1,Keyword,'"&KeyWord&"')>0"
				Case 2
					WSQL=" Where ColumnId="&Sour&" And InStr(1,Author,'"&KeyWord&"')>0"
				Case 3
					WSQL=" Where ColumnId="&Sour&" And InStr(1,Summary,'"&KeyWord&"')>0"
				Case 4
					WSQL=" Where ColumnId="&Sour&" And InStr(1,Source,'"&KeyWord&"')>0"
				Case 5
					WSQL=" Where ColumnId="&Sour&" And InStr(1,Content,'"&KeyWord&"')>0"
				End Select
			
				Sql="Select Count([Id]) From [NB_Content]"&WSQL
				Count=EA_M_DBO.DB_Query(SQL)(0, 0)
				
				Sql="UpDate [NB_Column] Set CountNum=CountNum-"&Count&" Where [Id]="&Sour
				EA_M_DBO.DB_Execute(SQL)
			End If
		End Select
		
		Sql="UpDate [NB_System] Set TopicNum=TopicNum-"&Count
		EA_M_DBO.DB_Execute(SQL)
		
		If Back=0 Then 
			Sql="UpDate [NB_Content] Set IsDel=1"&WSQL
		Else
			Sql="Delete From [NB_Content]"&WSQL
		End If
		EA_M_DBO.DB_Execute(SQL)
		
	End If
	
	ErrMsg=Replace(str_Bath_DelMsg,"$1",Count)
	Response.Write ErrMsg
End Sub
%>