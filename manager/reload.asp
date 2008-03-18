<!--#Include File="comm/inc.asp" -->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2004 - 2008 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：/Manager/ReLoad.asp
'= 摘    要：后台-数据更新文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-03-18
'====================================================================

Server.ScriptTimeout=9999999

Call EA_Manager.Chk_IsMaster

Dim Atcion
Dim ForTotal
Atcion=Request.Form ("action")

Select Case LCase(Atcion)
Case "updata"
	If Not EA_Manager.Chk_Power(Admin_Power,"07") Then 
		ErrMsg=str_Comm_NotAccess
		Call EA_Manager.Error(1)
	Else
		Call UpData
	End If
Case "sitemaps"
	Call Make_Sitemaps()
Case "baidu_newsop"
	Call Make_BaiduNewsop()
Case Else
	If Not EA_Manager.Chk_Power(Admin_Power,"07") Then 
		ErrMsg=str_Comm_NotAccess
		Call EA_Manager.Error(1)
	Else
		Call Main
	End If
End Select
Call EA_Pub.Close_Obj
Set EA_Pub=Nothing

Sub Main
	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_ReLoad_Help",str_ReLoad_Help)

	Call EA_M_XML.AppElements("Language_ReLoad_UpDateSystem",str_ReLoad_UpDateSystem)
	Call EA_M_XML.AppElements("Language_ReLoad_MakeSitemaps",str_ReLoad_MakeSitemaps)
	Call EA_M_XML.AppElements("Language_ReLoad_MakeSitemaps_Desc",str_ReLoad_MakeSitemaps_Desc)
	Call EA_M_XML.AppElements("Language_ReLoad_MakeBaiduNewsop",str_ReLoad_MakeBaiduNewsop)
	Call EA_M_XML.AppElements("Language_ReLoad_MakeBaiduNewsop_Desc",str_ReLoad_MakeBaiduNewsop_Desc)

	Call EA_M_XML.AppElements("btnSubmit1",str_Comm_Submit_Button)
	Call EA_M_XML.AppElements("btnSubmit2",str_Comm_Submit_Button)
	Call EA_M_XML.AppElements("btnSubmit3",str_Comm_Submit_Button)

	Page = EA_M_XML.make("","",0)
	Call EA_M_XML.Out(Page)
End Sub

Sub UpData()
	Call EA_Pub.Chk_Post

	Dim EA_Ini,strConfigFile

	Set EA_Ini=New cls_Ini

	strConfigFile	= Server.MapPath (SystemFolder&"include/config.ini")
	EA_Ini.OpenFile	= strConfigFile
	
	Dim ArticleTotal,MangerArticleTotal,MemberTotal,ColumnTotal,ReviewTotal
	Dim i,TempTotal_A,TempTotal_B,TempArray
	
	SQL="Select Count([Id]) From [NB_Content] Where IsPass="&EA_M_DBO.TrueValue&" And IsDel=0"
	ArticleTotal=EA_M_DBO.DB_Query(SQL)(0, 0)
	
	SQL="Select Count([Id]) From [NB_Content] Where IsPass=0 And IsDel=0"
	MangerArticleTotal=EA_M_DBO.DB_Query(SQL)(0, 0)
	
	SQL="Select Count([Id]) From [NB_Review] Where IsPass="&EA_M_DBO.TrueValue
	ReviewTotal=EA_M_DBO.DB_Query(SQL)(0, 0)
	
	SQL="Select Count([Id]) From [NB_Column]"
	ColumnTotal=EA_M_DBO.DB_Query(SQL)(0, 0)
	
	SQL="Select Count([Id]) From [NB_User]"
	MemberTotal=EA_M_DBO.DB_Query(SQL)(0, 0)
	
	SQL="UpDate [NB_System] Set "
	SQL=SQL&" RegUser="&MemberTotal
	SQL=SQL&",TopicNum="&ArticleTotal
	SQL=SQL&",ColumnNum="&ColumnTotal
	SQL=SQL&",MangerTopicNum="&MangerArticleTotal
	SQL=SQL&",ReviewNum="&ReviewTotal
	EA_M_DBO.DB_Execute(SQL)
	
	Call EA_Ini.WriteNode("System","Column_Total",ColumnTotal)
	Call EA_Ini.WriteNode("System","Topic_Total",ArticleTotal)
	Call EA_Ini.WriteNode("System","M_Topic_Total",MangerArticleTotal)
	Call EA_Ini.WriteNode("System","User_Total",MemberTotal)
	Call EA_Ini.WriteNode("System","Review_Total",ReviewTotal)
	EA_Ini.Save
	EA_Ini.Close
	Set EA_Ini=Nothing
	
	TempArray=EA_DBO.Get_Column_List()
	If IsArray(TempArray) Then 
		ForTotal = UBound(TempArray,2)

		For i=0 To ForTotal
			SQL="Select Count([Id]) From [NB_Content] Where ColumnId="&TempArray(0,i)&" And IsPass="&EA_M_DBO.TrueValue&" And IsDel=0"
			TempTotal_A=EA_M_DBO.DB_Query(SQL)(0, 0)
			
			SQL="Select Count([Id]) From [NB_Content] Where ColumnId="&TempArray(0,i)&" And IsPass=0 And IsDel=0"
			TempTotal_B=EA_M_DBO.DB_Query(SQL)(0, 0)
			
			SQL="UpDate [NB_Column] Set CountNum="&TempTotal_A&",MangerNum="&TempTotal_B&" Where [Id]="&TempArray(0,i)
			EA_M_DBO.DB_Execute(SQL)
		Next
	End If
	
	TempArray=EA_M_DBO.Get_Group_List()
	If IsArray(TempArray) Then 
		ForTotal = UBound(TempArray,2)

		For i=0 To ForTotal
			SQL="Select Count([Id]) From [NB_User] Where User_Group="&TempArray(0,i)
			TempTotal_A=EA_M_DBO.DB_Query(SQL)(0, 0)
			
			SQL="UpDate [NB_UserGroup] Set UserTotal="&TempTotal_A&" Where [Id]="&TempArray(0,i)
			EA_M_DBO.DB_Execute(SQL)
		Next
	End If
	
	If iDataBaseType=0 Then
		SQL="UpDate [NB_Content] a Left Join [NB_Column] b On a.ColumnId=b.[Id] Set a.ColumnName=b.Title,a.ColumnCode=b.Code"
	Else
		SQL="UpDate [NB_Content] Set ColumnName=b.Title,ColumnCode=b.Code From [NB_Content] a Join [NB_Column] b On a.ColumnId=b.[Id] "
	End If
	EA_M_DBO.DB_Execute(SQL)
	
	Set Rs=Nothing
	
	Response.Write str_BatchOperationMessageForSucess
End Sub

Sub Make_Sitemaps()
	Dim FileIndex, FileTotal
	Dim IndexContent, Content
	Dim IndexListBlock, ContentListBlock
	Dim Block
	Dim SitemapsFront
	Dim i, j, k, TempArray
	Dim PageCount, FileName
	Dim Template

	Set Template=New cls_NEW_TEMPLATE

	FileIndex = 1
	FileTotal = 0

	IndexContent = Template.LoadTemplate("sitemap-index.xml")
	Content		 = Template.LoadTemplate("sitemaps.xml")

	IndexListBlock		= Template.GetBlock("list",IndexContent)
	ContentListBlock	= Template.GetBlock("list",Content)

	SitemapsFront = EA_Pub.SysInfo(11)

	Block = ContentListBlock
	Template.SetVariable "SitemapsFile", EA_Pub.SysInfo(11), Block
	Template.SetVariable "priority", "1.0", Block
	Template.SetVariable "changefreq", "weekly", Block

	Template.SetBlock "list", Block, Content

	'Sort List
	SQL = "SELECT [ID],CountNum,IsOut,PageSize,ListPower,IsHide FROM [NB_Column]"
	TempArray = EA_DBO.DB_Query(SQL)

	If IsArray(TempArray) Then
		ForTotal = UBound(TempArray,2)

		For i = 0 To ForTotal
			Block = ContentListBlock

			If (TempArray(2,i) = "1") Or (CDbl(TempArray(4,i)) > 0) Or (TempArray(5,i) = "1") Then
				Template.SetVariable "SitemapsFile", SitemapsFront & "list.asp?classid=" & TempArray(0,i), Block
				Template.SetVariable "priority", "0.5", Block
				Template.SetVariable "changefreq", "weekly", Block

				Template.SetBlock "list", Block, Content

				FileTotal = FileTotal + 1

				If FileTotal = 50000 Then Call Save_Sitemaps(FileIndex, FileTotal, Content, IndexListBlock, IndexContent, SitemapsFront, Template)
			Else
				PageCount = EA_Pub.Stat_Page_Total(TempArray(3,i),TempArray(1,i))

				For j = 1 To PageCount
					If EA_Pub.SysInfo(18) = "1" Then
						Template.SetVariable "SitemapsFile", SitemapsFront & "list.asp?classid=" & TempArray(0,i) & "&amp;page=" & j, Block
						Template.SetVariable "priority", "0.5", Block
						Template.SetVariable "changefreq", "weekly", Block

						Template.SetBlock "list", Block, Content

						Block = ContentListBlock

						FileTotal = FileTotal + 1

						If FileTotal = 50000 Then Call Save_Sitemaps(FileIndex, FileTotal, Content, IndexListBlock, IndexContent, SitemapsFront, Template)
					Else
						Template.SetVariable "SitemapsFile", SitemapsFront & "articlelist/article_" & TempArray(0,i) & "_adddate_desc_" & j & ".htm", Block
						Template.SetVariable "priority", "0.5", Block
						Template.SetVariable "changefreq", "weekly", Block

						Template.SetBlock "list", Block, Content

						Block = ContentListBlock

						FileTotal = FileTotal + 1

						If FileTotal = 50000 Then Call Save_Sitemaps(FileIndex, FileTotal, Content, IndexListBlock, IndexContent, SitemapsFront, Template)
					End If
				Next
			End If
		Next
	End If

	'Article List
	SQL = "SELECT a.[ID],a.AddDate,a.IsOut,b.ListPower,b.IsHide"
	SQL = SQL & " FROM [NB_Content] a RIGHT JOIN [NB_Column] b ON a.ColumnID=b.[ID]"
	SQL = SQL & " WHERE IsDel=0 AND IsPass=" & EA_DBO.TrueValue
	TempArray = EA_DBO.DB_Query(SQL)
	
	If IsArray(TempArray) Then
		ForTotal = UBound(TempArray,2)

		For i = 0 To ForTotal
			Block = ContentListBlock

			If (TempArray(2,i) = "1") Or (CDbl(TempArray(3,i)) > 0) Or (TempArray(4,i) = "1") Then
				Template.SetVariable "SitemapsFile", SitemapsFront & "article.asp?articleid=" & TempArray(0,i), Block
				Template.SetVariable "priority", "0.8", Block
				Template.SetVariable "changefreq", "daily", Block
			Else
				Template.SetVariable "SitemapsFile", SitemapsFront & "articleview/" & Year(TempArray(1,i)) & "-" & Month(TempArray(1,i)) & "-" & Day(TempArray(1,i)) & "/article_view_" & TempArray(0,i) & ".htm", Block
				Template.SetVariable "priority", "0.8", Block
				Template.SetVariable "changefreq", "daily", Block
			End If

			Template.SetBlock "list", Block, Content

			FileTotal = FileTotal + 1

			If FileTotal = 50000 Then Call Save_Sitemaps(FileIndex, FileTotal, Content, IndexListBlock, IndexContent, SitemapsFront, Template)
		Next
	End If

	Call Save_Sitemaps(FileIndex, FileTotal, Content, IndexListBlock, IndexContent, SitemapsFront, Template)

	Template.CloseBlock "list",IndexContent

	FileName = "../sitemap-index.xml"

	Template.SetVariable "LastModTime",Year(Date()) & "-" & Right("00"&Month(Date()),2) & "-" & Right("00"&Day(Date()),2),IndexContent

	Call EA_Pub.Save_HtmlFile(FileName, IndexContent)

	Response.Write str_BatchOperationMessageForSucess
End Sub

Sub Save_Sitemaps(ByRef iFileIndex, ByRef iFileTotal, sFileContent, sIndexListBlock, ByRef sIndexContent, sSitemapsFront,ByRef Template)
	Dim FileName

	FileName = "../sitemaps-" & Date() & "-" & iFileIndex & ".xml"

	Template.CloseBlock "list",sFileContent

	Template.SetVariable "LastModTime",Year(Date()) & "-" & Right("00"&Month(Date()),2) & "-" & Right("00"&Day(Date()),2),sFileContent

	Call EA_Pub.Save_HtmlFile(FileName, sFileContent)

	Template.SetVariable "SitemapsFile", sSitemapsFront & "sitemaps-" & Date() & "-" & iFileIndex & ".xml", sIndexListBlock
	Template.SetBlock "list", sIndexListBlock, sIndexContent

	iFileIndex = iFileIndex + 1
	iFileTotal = 0
End Sub

Sub Make_BaiduNewsop ()
	Dim Content
	Dim ContentListBlock
	Dim Block
	Dim SitemapsFront
	Dim i,TempArray
	Dim FileName
	Dim Template

	Set Template=New cls_NEW_TEMPLATE

	Content				= Template.LoadTemplate("baidu-newsop.xml")
	ContentListBlock	= Template.GetBlock("list",Content)

	SitemapsFront = EA_Pub.SysInfo(11)

	Template.SetVariable "webSite", EA_Pub.SysInfo(11), Content
	Template.SetVariable "webMaster", EA_Pub.SysInfo(12), Content

	Template.SetBlock "list", Block, Content

	'Article List
	SQL = "SELECT TOP 100 a.[ID],a.AddDate,a.IsOut,b.ListPower,b.IsHide,a.Title,a.Summary,a.Content,a.KeyWord,b.Title,a.Author,a.Source,a.Img"
	SQL = SQL & " FROM [NB_Content] a RIGHT JOIN [NB_Column] b ON a.ColumnID=b.[ID]"
	SQL = SQL & " WHERE IsDel=0 AND IsPass=" & EA_DBO.TrueValue
	SQL = SQL & " ORDER BY TrueTime DESC"
	TempArray = EA_DBO.DB_Query(SQL)

	If IsArray(TempArray) Then
		ForTotal = UBound(TempArray,2)

		For i = 0 To ForTotal
			Block = ContentListBlock

			If (TempArray(2,i) = "1") Or (CDbl(TempArray(3,i)) > 0) Or (TempArray(4,i) = "1") Then
				Template.SetVariable "link", SitemapsFront & "article.asp?articleid=" & TempArray(0,i), Block
			Else
				Template.SetVariable "link", SitemapsFront & "articleview/" & Year(TempArray(1,i)) & "-" & Month(TempArray(1,i)) & "-" & Day(TempArray(1,i)) & "/article_view_" & TempArray(0,i) & ".htm", Block
			End If

			Template.SetVariable "title", TempArray(5,i), Block
			Template.SetVariable "description", EA_Pub.SafeRequest(0, TempArray(6,i), 1, "", 3), Block
			Template.SetVariable "text", EA_Pub.SafeRequest(0, TempArray(7,i), 1, "", 3), Block
			Template.SetVariable "headlineImg", TempArray(12,i), Block
			Template.SetVariable "keywords", Replace(TempArray(8,i), ",", " "), Block
			Template.SetVariable "category", TempArray(9,i), Block
			Template.SetVariable "author", TempArray(10,i), Block
			Template.SetVariable "source", TempArray(11,i), Block
			Template.SetVariable "pubDate", TempArray(1,i), Block

			Template.SetBlock "list", Block, Content
		Next
	End If

	Template.CloseBlock "list",Content

	Call EA_Pub.Save_HtmlFile("../baidu-newsop.xml", Content)
	Response.Write str_BatchOperationMessageForSucess
End Sub
%>