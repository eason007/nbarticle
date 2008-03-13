<!--#Include File="include/inc.asp"-->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：Search.asp
'= 摘    要：搜索文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-03-13
'====================================================================

Dim Action
Dim ForTotal
Action=Request.QueryString ("action")

Select Case LCase(Action)
Case "query"
	Call Request_Query()
End Select
Call EA_Pub.Close_Obj
Set EA_Pub=Nothing

Sub Request_Query()
	Dim PageNum,KeyWord,Column,Types,StartDate,EndTime,IsInclude,PageSize
	
	PageNum	= EA_Pub.SafeRequest(1,"page",0,1,0)
	KeyWord	= EA_Pub.SafeRequest(1,"keyword",1,"",0)
	Column	= EA_Pub.SafeRequest(1,"column",1,"0|",0)
	Column	= Split(Column,"|")
	If UBound(Column)<>1 Then Column=Split("0|","|")
	Column(0)= EA_Pub.SafeRequest(0,Column(0),0,0,0)
	Column(1)= EA_Pub.SafeRequest(0,Column(1),1,"",0)
	Types	 = EA_Pub.SafeRequest(1,"field",0,0,0)
	StartDate= EA_Pub.SafeRequest(1,"stime",2,"1900-1-1",0)
	EndTime	 = EA_Pub.SafeRequest(1,"etime",2,FormatDateTime(Now()+1,2),0)
	IsInclude= EA_Pub.SafeRequest(1,"isinclude",0,0,0)
	PageSize = 15

	If KeyWord="" And Column(0)=0 And StartDate="1900-1-1" And CStr(EndTime)=CStr(FormatDateTime(Now()+1,2)) Then Exit Sub

	Dim PageCount,ReCount
	Dim i,WSQL
	Dim QueryArray,QueryList
	Dim PageContent
	Dim SQL
	Dim Temp,ListBlock, Url
	
	If iDataBaseType=0 Then
		WSQL="Where IsPass="&EA_DBO.TrueValue&" And IsDel=0 And AddDate Between #"&StartDate&"# And #"&EndTime&"# "
	Else
		WSQL="Where IsPass="&EA_DBO.TrueValue&" And IsDel=0 And AddDate Between '"&StartDate&"' And '"&EndTime&"' "
	End If
	Select Case Types
	Case 1
		WSQL=WSQL&MakeSQLQuery("keyword",KeyWord)
	Case 2
		WSQL=WSQL&MakeSQLQuery("Author",KeyWord)
	Case 3
		WSQL=WSQL&MakeSQLQuery("Summary",KeyWord)
	Case Else
		WSQL=WSQL&MakeSQLQuery("Title",KeyWord)
	End Select
	
	If Column(0)>0 Then 
		If IsInclude=1 Then 
			WSQL=WSQL&" And Left(ColumnCode,"&Len(Column(1))&")='"&Column(1)&"'"
		Else
			WSQL=WSQL&" And ColumnId="&Column(0)
		End If
	End If


	PageContent=EA_Temp.Load_Template(0, 6)

	ListBlock = EA_Temp.GetBlock("Search.Topic", PageContent)

	SQL="Select Count([Id]) From [NB_Content] "&WSQL
	ReCount=EA_DBO.DB_Query(SQL)(0,0)
	PageCount=EA_Pub.Stat_Page_Total(PageSize,ReCount)
	If PageNum>PageCount And PageCount>0 Then PageNum=PageCount
	
	If PageCount>0 Then 
		WSQL = WSQL & " ORDER BY TrueTime DESC"
		SQL="Select [Id],ColumnId,ColumnName,IsImg,IsTop,TColor,Title,AddDate,Author,ViewNum,CommentNum,Summary From [NB_Content] "&WSQL
		QueryArray=EA_DBO.DB_CutPageQuery(SQL,PageNum,PageSize)

		If IsArray(QueryArray) Then
			ForTotal = UBound(QueryArray,2)

			For i=0 To ForTotal
				Temp=ListBlock
  
				EA_Temp.SetVariable "ColumnUrl",EA_Pub.Cov_ColumnPath(QueryArray(1,i),EA_Pub.SysInfo(18)),Temp
				EA_Temp.SetVariable "Column",QueryArray(2,i),Temp
				EA_Temp.SetVariable "Title",EA_Pub.Add_ArticleColor(QueryArray(5,i),QueryArray(6,i)),Temp
				EA_Temp.SetVariable "Url",EA_Pub.Cov_ArticlePath(QueryArray(0,i),QueryArray(7,i),EA_Pub.SysInfo(18)),Temp
				EA_Temp.SetVariable "Icon",EA_Pub.Chk_ArticleType(QueryArray(3,i),QueryArray(4,i)),Temp
				EA_Temp.SetVariable "Date", FormatDateTime(QueryArray(7,i), 2), Temp
				EA_Temp.SetVariable "Time", FormatDateTime(QueryArray(7,i), 4), Temp
				EA_Temp.SetVariable "ViewNum",QueryArray(9,i),Temp
				EA_Temp.SetVariable "Author",QueryArray(8,i),Temp
				EA_Temp.SetVariable "CommentNum",QueryArray(10,i),Temp
				EA_Temp.SetVariable "Summary",EA_Pub.Full_HTMLFilter(QueryArray(11,i)),Temp

				EA_Temp.SetBlock "Search.Topic", Temp, PageContent
			Next 
		End If
	End If
	EA_Temp.CloseBlock "Search.Topic", PageContent

	If EA_Temp.ChkTag("Search.PageNav", PageContent) Then 
		Url = "?keyword=" & KeyWord & "&column=" & Join(Column,"|") & "&field=" & Types & "&stime=" & StartDate & "&etime=" & EndTime & "&isinclude=" & IsInclude & "&action=query&page=$page"

		EA_Temp.SetVariable "Search.PageNav", EA_Temp.PageList(PageCount, PageNum, Url), PageContent
	End If

	EA_Temp.Title= SysMsg(28) & " - " & EA_Pub.SysInfo(0)
	EA_Temp.Nav	 = "<a href=""" & SystemFolder & """>" & EA_Pub.SysInfo(0) & "</a> - " & SysMsg(28)

	PageContent = EA_Temp.Replace_PublicTag(PageContent)
	
	Response.Write PageContent
End Sub

Function MakeSQLQuery(QueryField,QueryStr)
	Dim TagStart,TagEnd
	Dim TempStr,TempArray
	Dim FullQueryStr
	Dim i,Way
	
	'先找引号定界符
	Do
		TagStart=InStr(QueryStr,"\")
		If TagStart>0 Then 
			TagEnd=InStr(TagStart+1,QueryStr,"\")
			
			TempStr=Mid(QueryStr,TagStart+1,TagEnd-TagStart-1)
			TempStr=Replace(TempStr," ","#")
			
			QueryStr=Left(QueryStr,TagStart-1)&TempStr&Right(QueryStr,Len(QueryStr)-TagEnd)
		End If
	Loop While TagStart>0
	
	
	QueryStr = Replace(QueryStr,"|"," @")	'处理or定界符
	TempArray= Split(QueryStr," ")			'分隔关键字
	ForTotal = UBound(TempArray)

	For i=0 To ForTotal
		If Left(TempArray(i),1)="@" Then
			FullQueryStr=FullQueryStr&" Or "&QueryField
			TempArray(i)=Right(TempArray(i),Len(TempArray(i))-1)
		Else
			FullQueryStr=FullQueryStr&" And "&QueryField
		End If
		
		If Left(TempArray(i),1)="-" Then 
			FullQueryStr=FullQueryStr&" Not "
			TempArray(i)=Right(TempArray(i),Len(TempArray(i))-1)
		End If
		
		FullQueryStr=FullQueryStr&" Like '%"&TempArray(i)&"%'"
		
		FullQueryStr=Replace(FullQueryStr,"%$","")
		FullQueryStr=Replace(FullQueryStr,"$%","")
		FullQueryStr=Replace(FullQueryStr,"#"," ")
	Next
	
	MakeSQLQuery=FullQueryStr
End Function
%>