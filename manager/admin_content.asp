<!--#Include File="comm/inc.asp" -->
<%
'====================================================================
'= Team Elite - EliteCMS
'= Copyright (c) 2004 - 2008 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：/Manager/Admin_Content.asp
'= 摘    要：后台-文章管理文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-04-01
'====================================================================

Response.Clear

Call EA_Manager.Chk_IsMaster

If Not EA_Manager.Chk_Power(Admin_Power,"12") Then 
	ErrMsg=str_Comm_NotAccess
	Call EA_Manager.Error(1)
End If

Dim Action
Dim ForTotal

Action=Request.Form("action")

Select Case LCase(Action)
Case "add"
	Call Add
Case "save"
	Call Save
Case "del"
	Call Del
Case "batch"
	Call Batch
Case Else
	Call Main
End Select
Call EA_Pub.Close_Obj
Set EA_Pub=Nothing

Sub Main
	Dim Count,Page,i
	Dim TopicList
	Dim ListName(8),ListValue()
	Dim Level
	Dim WStr,ColumnId,KeyWord,Field
	Dim Temp,Tmp
	
	Page	= EA_Pub.SafeRequest(2,"nowPage",0,1,0)
	WStr	= EA_Pub.SafeRequest(2,"w",1,"",0)
	ColumnId= EA_Pub.SafeRequest(2,"Column",0,0,0)
	KeyWord	= EA_Pub.SafeRequest(2,"Keyword",1,"",0)
	Field	= EA_Pub.SafeRequest(2,"Field",0,-1,0)
	
	Select Case WStr
	Case "1d"
		If iDataBaseType=0 Then 
			WStr=" where datediff('d',adddate,Now())=0 and isdel=0 "
		Else
			WStr=" where datediff(d,adddate,GetDate())=0 and isdel=0 "
		End If
	Case "1w"
		If iDataBaseType=0 Then 
			WStr=" where datediff('ww',adddate,Now())=0 and isdel=0 "
		Else
			WStr=" where datediff(ww,adddate,GetDate())=0 and isdel=0 "
		End If
	Case "1m"
		If iDataBaseType=0 Then 
			WStr=" where datediff('m',adddate,Now())=0 and isdel=0 "
		Else
			WStr=" where datediff(m,adddate,GetDate())=0 and isdel=0 "
		End If
	Case "pass"
		WStr=" where ispass="&EA_DBO.TrueValue&" and isdel=0 "
	Case "npass"
		WStr=" where ispass=0 and isdel=0 "
	Case "search"
		WStr=" where isdel=0"
		If Len(KeyWord) > 0 Then
			Select Case Field
			Case 0
				If iDataBaseType=0 Then 
					WStr=WStr&" and InStr(1,title,'"&KeyWord&"')>0"
				Else
					WStr=WStr&" and CharIndex('"&KeyWord&"',title)>0"
				End If
			Case 1
				If iDataBaseType=0 Then 
					WStr=WStr&" and InStr(1,keyword,'"&KeyWord&"')>0"
				Else
					WStr=WStr&" and CharIndex('"&KeyWord&"',keyword)>0"
				End If
			Case 2
				If iDataBaseType=0 Then 
					WStr=WStr&" and InStr(1,author,'"&KeyWord&"')>0"
				Else
					WStr=WStr&" and CharIndex('"&KeyWord&"',author)>0"
				End If
			End Select
		End If
		If ColumnId<>0 Then WStr=WStr&" and columnid="&ColumnId
	Case "reback"
		WStr=" where isdel="&EA_DBO.TrueValue
	Case Else
		WStr=" where isdel=0"
	End Select

	Temp=EA_DBO.Get_Column_List()
	Tmp = "(build-select),0 " & str_Comm_Select & ",0"
	If IsArray(Temp) Then
		ForTotal = UBound(Temp,2)

		For i=0 To ForTotal
			Level = (Len(Temp(2,i))/4-1)*3

			Tmp = Tmp & " ├" & String(Level,"-") & Replace(Temp(1,i)," ","") & "," & Temp(0,i)
		Next
	End If
	Call EA_M_XML.AppInfo("ColumnList",Tmp)

	Tmp = "(build-select),0 " & str_Content_Title & ",0 " & str_Content_Keyword & ",1 " & str_Content_Author & ",2"
	Call EA_M_XML.AppInfo("Field",Tmp)

	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_Content_Help",str_Content_Help)

	Call EA_M_XML.AppElements("Language_Comm_AllColumn",str_Comm_AllColumn)
	Call EA_M_XML.AppElements("Language_Comm_Today",str_Comm_Today)
	Call EA_M_XML.AppElements("Language_Comm_Week",str_Comm_Week)
	Call EA_M_XML.AppElements("Language_Comm_Month",str_Comm_Month)
	Call EA_M_XML.AppElements("Language_Comm_State_NoPass",str_Comm_State_NoPass)
	Call EA_M_XML.AppElements("Language_Comm_State_Pass",str_Comm_State_Pass)
	Call EA_M_XML.AppElements("Language_Comm_RecycleBin",str_Comm_RecycleBin)

	Call EA_M_XML.AppElements("Language_Content_Title",str_Content_Title)
	Call EA_M_XML.AppElements("Language_Content_Column",str_Content_Column)
	Call EA_M_XML.AppElements("Language_Content_State",str_Content_State)
	Call EA_M_XML.AppElements("Language_Content_ViewNum",str_Content_ViewNum)
	Call EA_M_XML.AppElements("Language_Content_Review",str_Content_Review)
	Call EA_M_XML.AppElements("Language_Content_Date",str_Content_Date)

	Call EA_M_XML.AppElements("Language_QuickSearchArticle",str_QuickSearchArticle)

	Call EA_M_XML.AppElements("Comm_Add_Operation",str_Comm_Add_Operation)
	Call EA_M_XML.AppElements("Comm_Del_Operation",str_Comm_Del_Operation)
	Call EA_M_XML.AppElements("Comm_Resume_Operation",str_Content_Resume)
	Call EA_M_XML.AppElements("Comm_Pass_Operation",str_Comm_State_Pass)
	Call EA_M_XML.AppElements("Comm_NoPass_Operation",str_Comm_State_NoPass)

	Call EA_M_XML.AppElements("Language_Comm_Bar_Operation",str_Comm_Bar_Operation)
	Call EA_M_XML.AppElements("btnSubmit",str_Comm_Submit_Button)

	SQL="Select Count(Id) From [NB_Content]"&WStr
	Count=EA_M_DBO.DB_Query(SQL)(0, 0)
	If Count>0 Then 
		If Rs.State=1 Then Rs.Close
		SQL="Select Id,Title,ColumnName,AddDate,ViewNum,CommentNum,IsPass,ColumnId From [NB_Content] "&WStr&" Order By TrueTime Desc"
		Rs.Open SQL,Conn,1,1
		If Not rs.eof And Not rs.bof Then 
			Rs.AbsolutePosition=Rs.AbsolutePosition+((Abs(Page)-1)*10)
			TopicList=Rs.GetRows(10)
		End If
		Rs.Close
		Set Rs=Nothing

		ListName(0) = "checkbox"
		ListName(1) = "ID"
		ListName(2) = "C_ID"
		ListName(3) = "Title"
		ListName(4) = "Column"
		ListName(5) = "State"
		ListName(6) = "Stat"
		ListName(7) = "Date"
		ListName(8) = "action"
		ForTotal = Ubound(TopicList,2)

		For i=0 To ForTotal
			ReDim Preserve ListValue(8,i)

			ListValue(0,i) = "checkbox"
			ListValue(1,i) = TopicList(0,i)
			ListValue(2,i) = TopicList(0,i)
			ListValue(3,i) = "<a href='" & EA_Pub.Cov_ArticlePath(TopicList(0,i),TopicList(3,i),EA_Pub.SysInfo(18)) & "' target='_blank'>" & TopicList(1,i)  & "</a>"
			ListValue(4,i) = TopicList(2,i)
			If TopicList(6,i) Then
				ListValue(5,i) = "<strong>√</strong>"
			Else
				ListValue(5,i) = "<font color=""red""><strong>×</strong></font>"
			End If
			ListValue(6,i) = TopicList(4,i) & "/" & TopicList(5,i)
			ListValue(7,i) = FormatDateTime(TopicList(3,i),2)
			ListValue(8,i) = "action"
		Next

		Page = EA_M_XML.make(ListName,ListValue,Count)
	Else
		Page = EA_M_XML.make("","",0)
	End If

	Call EA_M_XML.Out(Page)
End Sub

Sub Add
	Dim PostId,ArticleTemplate_Id,Column_Id
	Dim T_Color
	Dim Level,TStr,Tmp
	Dim TempArray,i,Temp
	
	PostId = EA_Pub.SafeRequest(2,"ID",0,0,0)
	Call EA_M_XML.AppInfo("ID",PostId)

	Call EA_M_XML.AppElements("Language_Comm_Yes",str_Comm_Yes)
	Call EA_M_XML.AppElements("Language_Comm_No",str_Comm_No)

	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_Content_Add_Help",str_Content_Add_Help)
	Call EA_M_XML.AppElements("Language_Content_AddArticle",str_Content_AddArticle)

	Call EA_M_XML.AppElements("Language_Content_Title",str_Content_Title)
	Call EA_M_XML.AppElements("Language_Content_Column",str_Content_Column)
	Call EA_M_XML.AppElements("Language_Content_Keyword",str_Content_Keyword)
	Call EA_M_XML.AppElements("Language_Content_Author",str_Content_Author)
	Call EA_M_XML.AppElements("Language_Content_Summary",str_Content_Summary)
	Call EA_M_XML.AppElements("Language_Content_SummaryFromText",str_Content_SummaryFromText)
	Call EA_M_XML.AppElements("Language_Content_Img",str_Content_Img)
	Call EA_M_XML.AppElements("Language_Content_ReviewImg",str_Content_ReviewImg)
	Call EA_M_XML.AppElements("Language_Content_Top",str_Content_Top)
	Call EA_M_XML.AppElements("Language_Content_OutURL",str_Content_OutURL)
	Call EA_M_XML.AppElements("Language_Content_Source",str_Content_Source)
	Call EA_M_XML.AppElements("Language_Content_ViewNum",str_Content_ViewNum)
	Call EA_M_XML.AppElements("Language_Content_Date",str_Content_Date)
	Call EA_M_XML.AppElements("Language_Content_PassNow",str_Content_PassNow)
	Call EA_M_XML.AppElements("Language_Content_Column",str_Content_Column)
	Call EA_M_XML.AppElements("Language_Content_Date",str_Content_Date)
	Call EA_M_XML.AppElements("Language_Content_Word",str_Content_Word)
	Call EA_M_XML.AppElements("Language_Content_ArticleTemplate",str_Content_ArticleTemplate)
	Call EA_M_XML.AppElements("Language_Content_Title",str_Content_Title)
	Call EA_M_XML.AppElements("Language_Content_Column",str_Content_Column)
	Call EA_M_XML.AppElements("Language_Content_SubTitle",str_Content_SubTitle)
	Call EA_M_XML.AppElements("Language_Content_SubUrl",str_Content_SubUrl)

	TempArray=EA_DBO.Get_Article_Info(PostId,0)
	If IsArray(TempArray) Then 
		Call EA_M_XML.AppInfo("title",TempArray(3,0))
		Call EA_M_XML.AppInfo("keyword",TempArray(12,0))
		Call EA_M_XML.AppInfo("authorid",TempArray(7,0))
		Call EA_M_XML.AppInfo("summary",EA_Pub.Un_Full_HTMLFilter(TempArray(4,0)))
		Call EA_M_XML.AppInfo("img",TempArray(18,0))
		Call EA_M_XML.AppInfo("istop",Abs(TempArray(19,0)))
		Call EA_M_XML.AppInfo("outurl",TempArray(11,0))
		Call EA_M_XML.AppInfo("source",TempArray(15,0))
		Call EA_M_XML.AppInfo("sourceurl",TempArray(16,0))
		Call EA_M_XML.AppInfo("viewnum",TempArray(6,0))
		Call EA_M_XML.AppInfo("adddate",TempArray(13,0))
		Call EA_M_XML.AppInfo("author",TempArray(8,0))
		Call EA_M_XML.AppInfo("ispass",Abs(TempArray(20,0)))
		Call EA_M_XML.AppInfo("Content",TempArray(5,0))
		Call EA_M_XML.AppInfo("subtitle",TempArray(26,0))
		Call EA_M_XML.AppInfo("suburl",TempArray(27,0))
		
		T_Color = TempArray(17,0)
		Column_Id = TempArray(0,0)

		Call EA_M_XML.AppElements("Language_Content_SaveAs",str_Content_SaveAs)
	Else
		Call EA_M_XML.AppInfo("ispass","1")
		Call EA_M_XML.AppInfo("adddate",Now())
		Call EA_M_XML.AppInfo("author",EA_Pub.SysInfo(21))
		Call EA_M_XML.AppInfo("viewnum","0")
		Call EA_M_XML.AppInfo("authorid","-" & EA_Manager.MasterID)

		ArticleTemplate_Id=EA_Pub.SafeRequest(2,"temp_id",0,0,0)
		If ArticleTemplate_Id > 0 Then Call EA_M_XML.AppInfo("Content",EA_M_DBO.Get_ArticleTemp_Info(ArticleTemplate_Id)(1,0))
	End If

	Temp=EA_DBO.Get_Column_List()
	Tmp = "(build-select)," & Column_Id & " " & str_Comm_Select & ",0"
	If IsArray(Temp) Then
		ForTotal = UBound(Temp,2)

		For i=0 To ForTotal
			If InStr(1,","&Column_Power&",",","&Temp(0,i)&"1,") > 0 Or InStr(1,","&Column_Power&",",","&Temp(0,i)&"2,") > 0 Or InStr(1,","&Column_Power&",",","&Temp(0,i)&"3,") > 0 Then
				Level	= (Len(Temp(2,i))/4-1)*3

				Tmp		= Tmp & " ├" & String(Level,"-") & Replace(Temp(1,i)," ","") & "-"
				If InStr(1,","&Column_Power&",",","&Temp(0,i)&"1,")>0 Then Tmp	= Tmp & "[" & str_Master_Column_Add & "]"
				If InStr(1,","&Column_Power&",",","&Temp(0,i)&"2,")>0 Then Tmp	= Tmp & "[" & str_Master_Column_Manager & "]"
				If InStr(1,","&Column_Power&",",","&Temp(0,i)&"3,")>0 Then Tmp	= Tmp & "[" & str_Master_Column_Edit & "]"

				Tmp		= Tmp & "," & Temp(0,i) & " "
			End If
		Next
	End If
	Call EA_M_XML.AppInfo("Column",Tmp)
	Tmp = ""

	Temp=EA_M_DBO.Get_ArticleTemp_List(1,50)
	Tmp = "(build-select)," & ArticleTemplate_Id & " " & str_Comm_Select & ",0"
	If IsArray(Temp) Then
		ForTotal = UBound(Temp,2)

		For i=0 To ForTotal
			Tmp = Tmp & " " & Temp(1,i) & "," & Temp(0,i)
		Next
	End If
	Call EA_M_XML.AppInfo("atemplate",Tmp)
	Tmp = ""

	Temp=EA_DBO.Get_System_Info()
	Tmp = "(build-select),0 " & str_Comm_Select & ",=="
	If IsArray(Temp) Then
		If Not IsNull(Temp(6,0)) Then 
			TempArray= Split(Temp(6,0),";")
			ForTotal = UBound(TempArray)

			For i=0 To ForTotal
				TStr=""
				TStr=Split(TempArray(i),"==")

				Tmp = Tmp & " " & TStr(0) & "," & TempArray(i)
			Next
		End If
	End If
	Call EA_M_XML.AppInfo("choosesource",Tmp)
	Tmp = ""

	Tmp = "(build-select)," & T_Color & " " & str_Content_TColor & ",0 " & str_Content_Color_Red & ",1 " & str_Content_Color_Green & ",2 " & str_Content_Color_Blue & ",3"
	Call EA_M_XML.AppInfo("color",Tmp)

	Call EA_M_XML.AppElements("btnSubmit",str_Comm_Save_Button)
	Call EA_M_XML.AppElements("btnReturn",str_Comm_Return_Button)

	Page = EA_M_XML.make("","",0)
	Call EA_M_XML.Out(Page)
End Sub

Sub Save
	Dim Title,Author,Text,KeyWord,ColumnId,ColumnName,ColumnCode,Byter,TColor,IsImg,ImgPath,IsTop,IsDis,OutUrl,IsOut,AuthorId,ViewNum,AddDate,IsPass,Source,SourceUrl,Summary, SubTitle, SubUrl, ArticleID
	Dim PostId,TempStr,TrueTime,IsSaveAs
	Dim Key,i

	FoundErr=False
	
	PostId		= EA_Pub.SafeRequest(2,"ID",0,0,0)
	Title		= EA_Pub.SafeRequest(2,"title",1,"",1)
	SubTitle	= EA_Pub.SafeRequest(2,"subtitle",1,"",1)
	SubUrl		= EA_Pub.SafeRequest(2,"suburl",1,"",1)

	Text		= Request.Form("Content")
	Text		= EA_Pub.BadWords_Filter(Text)

	Summary		= EA_Pub.BadWords_Filter(EA_Pub.SafeRequest(2,"summary",1,"",2))
	Summary		= Left(Summary,140)

	KeyWord		= EA_Pub.DistinctStr(EA_Pub.SafeRequest(2,"keyword",1,"",1))

	ColumnId	= EA_Pub.SafeRequest(2,"Column",0,0,0)
	If ColumnId > 0 Then
		TempStr		= EA_DBO.Get_Column_Info(ColumnId)
		ColumnName	= EA_Pub.SafeRequest(0,Trim(TempStr(0,0)),1,"",1)
		ColumnCode	= EA_Pub.SafeRequest(0,Trim(TempStr(1,0)),1,"",0)
	End If

	TColor		= EA_Pub.SafeRequest(2,"color",0,0,0)
	ImgPath		= EA_Pub.SafeRequest(2,"img",1,"",1)
	IsTop		= EA_Pub.SafeRequest(2,"istop",0,0,0)
	OutUrl		= EA_Pub.SafeRequest(2,"outurl",1,"",1)
	Author		= EA_Pub.SafeRequest(2,"author",1,"",1)
	AuthorId	= EA_Pub.SafeRequest(2,"authorid",0,0,0)
	ViewNum		= EA_Pub.SafeRequest(2,"viewnum",0,0,0)
	AddDate		= Trim(EA_Pub.SafeRequest(2,"adddate",2,Now(),0))
	IsPass		= EA_Pub.SafeRequest(2,"ispass",0,0,0)
	Source		= EA_Pub.SafeRequest(2,"source",1,"",1)
	SourceUrl	= EA_Pub.SafeRequest(2,"sourceurl",1,"",1)
	IsSaveAs	= EA_Pub.SafeRequest(2,"saveas",0,0,0)
	Byter		= Lenb(Text)

	If ColumnId = "" Or ColumnId = "0" Then FoundErr= True

	If PostId<>0 Then
		If InStr(1,","&Column_Power&",",","&ColumnId&"3,")<=0 Then 
			ErrMsg	= Replace(str_Comm_NotAccess,"$1",str_Comm_Edit_Operation)
			FoundErr= True
		End If

		If IsSaveAs=1 Then PostId=0
	Else
		If InStr(1,","&Column_Power&",",","&ColumnId&"1,")<=0 Then 
			ErrMsg	= Replace(str_Comm_NotAccess,"$1",str_Comm_Add_Operation)
			FoundErr= True
		End If
	End If

	If Len(Title)>150 Or Len(Title)=0 Then FoundErr= True
	If Len(Author)>16 Then 
		FoundErr=True
	ElseIf Len(Author)=0 then 
		Author="本站编辑"
	End If
	
	If FoundErr Then 
		Response.Write "-1"
	Else
		If OutUrl="" Then 
			IsOut=0
		Else
			IsOut=1
		End If
		If ImgPath="" Then 
			IsImg=0
		Else
			IsImg=1
		End If
		
		TrueTime=Year(CDate(AddDate))
		TrueTime=TrueTime&Right("00"&Month(CDate(AddDate)),2)
		TrueTime=TrueTime&Right("00"&Day(CDate(AddDate)),2)
		TrueTime=TrueTime&Right("00"&Hour(CDate(AddDate)),2)
		TrueTime=TrueTime&Right("00"&Minute(CDate(AddDate)),2)
		TrueTime=TrueTime&Right("00"&Second(CDate(AddDate)),2)
		
		Randomize Timer
		key="000000"&Cstr(Int((999999-1+100000)*Rnd+1))
		TrueTime=TrueTime&Right(Key,6)

		IsDis=EA_DBO.Get_Column_Info(ColumnId)(11,0)

		If Rs.State=1 Then rs.Close
		If PostId<>0 Then
			Sql="Select * From [NB_Content] Where [Id]="&PostId
			rs.Open Sql,Conn,2,2
		Else
			rs.Open "SELECT * FROM [NB_Content] WHERE 0=1",Conn,2,2
			rs.AddNew
		End If
			rs("title")		= Title
			rs("author")	= Author
			rs("authorid")	= AuthorId
			rs("Content")	= Text&" "
			rs("KeyWord")	= KeyWord
			rs("ColumnId")	= ColumnId
			rs("ColumnName")= ColumnName
			rs("ColumnCode")= ColumnCode
			rs("byte")		= Byter
			rs("tcolor")	= TColor
			rs("isimg")		= IsImg
			rs("img")		= ImgPath
			rs("istop")		= IsTop
			rs("IsDis")		= IsDis
			rs("outurl")	= OutUrl
			rs("isout")		= IsOut
			rs("ViewNum")	= ViewNum
			rs("AddDate")	= AddDate
			rs("IsPass")	= IsPass
			rs("Source")	= Source
			rs("SourceUrl")	= SourceUrl
			rs("Summary")	= Summary
			rs("TrueTime")	= TrueTime
			rs("SubTitle")	= SubTitle
			rs("SubUrl")	= SubUrl
			rs.update
		Rs.Close:Set Rs=Nothing

		If PostId=0 Then 
			If IsPass=0 Then 
				If iDataBaseType<>2 Then EA_DBO.Set_System_ManagerTopicTotal 1

				EA_DBO.Set_Column_ManagerTopicTotal ColumnId,1
			Else
				If iDataBaseType<>2 Then EA_DBO.Set_System_TopicTotal 1

				EA_DBO.Set_Column_TopicTotal ColumnId,1
			End If
		End If

		If PostId=0 Then 
			ArticleID = EA_M_DBO.Get_ArticleID(AuthorId, ColumnId, Byter)

			If IsArray(ArticleID) Then Call SetTag(KeyWord, ArticleID(0, 0), ColumnId)
		End If

		Response.Write PostId

		Application.Lock 
		Application(sCacheName&"IsFlush")=1
		Application.UnLock 
		
		Call EA_Pub.Close_Obj
		Set EA_Pub=Nothing
	End If

	Response.End
End Sub

Sub SetTag (KeyWord, ArticleID, ColumnID) 
	Dim TempArray, i
	Dim TagList

	If Len(KeyWord) > 0 Then
		TempArray = Split(KeyWord, ",")
		ForTotal = UBound(TempArray)
		TagList	= ","

		For i = 0 To ForTotal
			TempArray(i) = Trim(TempArray(i))

			If Len(TempArray(i)) > 0 And InStr(TagList, "," & TempArray(i) & ",") = 0 Then 
				EA_M_DBO.Set_Tag_Create TempArray(i), ArticleID, ColumnID
				TagList = TagList & TempArray(i) & ","
			End If
		Next
	End If
End Sub

Sub Batch()
	Dim Opertion,Id,ColumnId
	Dim TempArray,i,j,Temp
	
	Id=Request.Form ("ID")
	Opertion=Request.Form ("opertion")
	Id=Replace(Id," ","")
	TempArray=Split(Id,",")
	j=0
	
	Select Case Opertion
	Case "nodel"
		ForTotal = UBound(TempArray)

		For i=0 To ForTotal
			EA_M_DBO.Set_Article_Resume TempArray(i)
		Next

		j = UBound(TempArray) + 1
	Case "pass"
		ForTotal = UBound(TempArray)

		For i=0 To ForTotal
			Temp=EA_DBO.Get_Article_Info_Single(TempArray(i))
			If IsArray(Temp) Then 
				If InStr(1,","&Column_Power&",",","&Temp(0,0)&"2,")>0 Then 
					EA_M_DBO.Set_Article_PassState 1,TempArray(i)

					j=j+1
				End If
			End If
		Next
	Case "nopass"
		ForTotal = UBound(TempArray)

		For i=0 To ForTotal
			Temp=EA_DBO.Get_Article_Info_Single(TempArray(i))
			If IsArray(Temp) Then 
				If InStr(1,","&Column_Power&",",","&Temp(0,0)&"2,")>0 Then 
					EA_M_DBO.Set_Article_PassState 0,TempArray(i)

					j=j+1
				End If
			End If
		Next
	Case "move"
		Dim ColumnCode,ColumnName
		ColumnId=Request.Form ("column")
		
		If IsNumeric(ColumnId) And ColumnId<>"" And ColumnId<>"0" Then
			Temp=EA_DBO.Get_Column_Info(ColumnId)
			If IsArray(Temp) Then
				ColumnCode=Temp(1,0)
				ColumnName=Temp(0,0)
				
				SQL="UpDate [NB_Content] Set ColumnId="&ColumnId&",ColumnCode='"&ColumnCode&"',ColumnName='"&ColumnName&"' Where [Id] In ("&Id&")"
				'Response.Write sql
				EA_M_DBO.DB_Execute(SQL)
			End If
		End If
	Case "make"
		Response.Write "<table width=""550"" border=""0"" cellpadding=""0"" cellspacing=""0"" align=""center"">"
		Response.Write "<form method=""post"" action=""admin_makeview.asp?atcion=mark"" name=""makeform"">"
		Response.Write "<tr valign=""middle"" align=""center"" height=""22"">"
		Response.Write "<td>"
		Response.Write "<input type=hidden name=sid value="""&Id&""">"
		Response.Write "<input type=hidden name=tag value=""4"">"
		Response.Write "</td>"
		Response.Write "</tr>"
		Response.Write "</form>"
		Response.Write "</table>"
		Response.Write "<script language=""JavaScript"">makeform.submit()</script>"
	End Select
	
	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing

	Response.Write j
	Response.End
End Sub

Sub Del
	Dim IDs
	Dim i,Tmp,Temp
	Dim ColumnId,AuthroId,IsPass,ReviewTotal,PostTime
	Dim sHTMLFilePath
	Dim objFSO

	IDs = Split(Request("ID"), ",")
	ForTotal = UBound(IDs)

	For i = 0 To ForTotal
		Tmp = EA_Pub.SafeRequest(5,IDs(i),0,0,0)

		Temp=EA_DBO.Get_Article_Info_Single(Tmp)
		If IsArray(Temp) Then
			ColumnId=Temp(0,0)
			AuthroId=Temp(7,0)
			PostTime=Temp(13,0)
			IsPass=Temp(20,0)
		End If
		
		If InStr(1,","&Column_Power&",",","&ColumnId&"4,")<=0 Then 
			ErrMsg = Replace(str_Comm_NotAccess,"$1",str_Comm_Del_Operation)
			Call EA_Manager.Error(1)
		End If

		If iDataBaseType<>2 Then
			'更新会员发表统计
			EA_M_DBO.Set_Member_PostTotal -1,AuthroId
		
			'删除评论及更新系统统计
			SQL="Select Count(Id) From [NB_Review] Where ArticleId="&Tmp
			ReviewTotal=EA_M_DBO.DB_Query(SQL)(0, 0)
			SQL="Delete From [NB_Review] Where ArticleId="&Tmp
			EA_M_DBO.DB_Execute(SQL)

			'更新系统信息
			EA_DBO.Set_System_ReviewTotal ReviewTotal-(ReviewTotal*2)
			If IsPass Then 
				EA_DBO.Set_System_TopicTotal -1
				EA_DBO.Set_Column_TopicTotal ColumnId,-1
			Else
				EA_DBO.Set_System_ManagerTopicTotal -1
				EA_DBO.Set_Column_ManagerTopicTotal ColumnId,-1
			End If
		End If
		
		EA_M_DBO.Set_Article_Del Tmp
		
		On Error Resume Next
		'Delete HTML File
		sHTMLFilePath=EA_Pub.Cov_ArticlePath(Tmp,PostTime,"0")
		If EA_Pub.Chk_IsExistsHtmlFile(sHTMLFilePath) Then 
			Set objFSO = CreateObject("Scripting.FileSystemObject")
			objFSO.DeleteFile server.MapPath (sHTMLFilePath)
		End If
		
		Application.Lock 
		Application(sCacheName&"IsFlush")=1
		Application.UnLock 
	Next
	
	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing
	
	Response.Write "1"
	Response.End
End Sub
%>
