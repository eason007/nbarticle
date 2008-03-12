
<!--#Include File="comm/inc.asp" -->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：/Manager/ReLoad.asp
'= 摘    要：后台-数据更新文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2007-10-27
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
Case "markjs"
	If Not EA_Manager.Chk_Power(Admin_Power,"46") Then 
		ErrMsg=str_Comm_NotAccess
		Call EA_Manager.Error(1)
	Else
		Call MarkJs
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

Sub MarkJs
		On Error Resume Next
		Dim i,j,k,List,TopicList
		Dim file
		Dim OutStr
		
		OutStr="mpmenu1=new mMenu('首页','"&SystemFolder&"','self','','','','');"
		OutStr=OutStr&Chr(10)
		OutStr=OutStr&"mpmenu1.addItem(new mMenuItem('图片文章','"&SystemFolder&"img_list.asp','self',false,'图片文章',null,'','','',''));"
		OutStr=OutStr&Chr(10)
		OutStr=OutStr&"mpmenu1.addItem(new mMenuItem('会员列表','"&SystemFolder&"member_list.asp','self',false,'会员列表',null,'','','',''));"
		OutStr=OutStr&Chr(10)
		OutStr=OutStr&"mpmenu1.addItem(new mMenuItem('高级搜索','"&SystemFolder&"search.asp','self',false,'高级搜索',null,'','','',''));"
		
		'第一层菜单选择
		Sql="select id,title,code from NB_Column where len(code)=4 and IsTop="&EA_M_DBO.TrueValue&" order by code"
		Set rs=conn.execute(sql)
		If Not rs.eof And Not rs.bof Then 
			TopicList=rs.getrows()
			rs.close:Set rs=Nothing
			j=1		'初始化二级菜单标识号
			For i=0 To Ubound(TopicList,2)
				'生成一级菜单项
				OutStr=OutStr&"mpmenu"&i+2&"=new mMenu('"&TopicList(1,i)&"','"&EA_Pub.Cov_ColumnPath(TopicList(0,i),EA_Pub.SysInfo(18))&"','self','','','','');"
				OutStr=OutStr&Chr(10)
				
				'筛选当次一级菜单的下属菜单（二级）
				Sql="select id,title,code from NB_Column where left(code,4)='"&TopicList(2,i)&"' and len(code)=8 and id<>"&TopicList(0,i)&" and IsTop="&EA_M_DBO.TrueValue&" order by code"
				Set rs=conn.execute(sql)
				If Not rs.eof And Not rs.bof Then 
					List=rs.getrows()
					rs.close
					For k=0 To Ubound(List,2)
						'筛选当次二级菜单的下属菜单（三级）
						Sql="select id,title from nb_column where left(code,8)='"&List(2,k)&"' and len(code)=12 and id<>"&List(0,k)&" and IsTop="&EA_M_DBO.TrueValue&" order by code"
						Set rs=conn.execute(sql)
						If rs.eof And rs.bof Then		'判断是否有第三层
							OutStr=OutStr&"mpmenu"&i+2&".addItem(new mMenuItem('"&List(1,k)&"','"&EA_Pub.Cov_ColumnPath(List(0,k),EA_Pub.SysInfo(18))&"','self',false,'"&List(1,k)&"',null,'','','',''));"
							OutStr=OutStr&Chr(10)
						Else
							OutStr=OutStr&"msub"&j&"=new mMenuItem('"&List(1,k)&"','"&EA_Pub.Cov_ColumnPath(List(0,k),EA_Pub.SysInfo(18))&"','self',false,'','1','','','','');"
							OutStr=OutStr&Chr(10)

							Do While Not rs.eof		'历遍第三层项目
								OutStr=OutStr&"msub"&j&".addsubItem(new mMenuItem('"&rs(1)&"','"&EA_Pub.Cov_ColumnPath(Rs(0),EA_Pub.SysInfo(18))&"','self',false,'"&rs(1)&"',null,'','','',''));"
								OutStr=OutStr&Chr(10)

								rs.movenext
							Loop
							OutStr=OutStr&"mpmenu"&i+2&".addItem(msub"&j&")"		'关闭当次第三层
							OutStr=OutStr&Chr(10)

							j=j+1
						End If
						rs.close
					Next
				End If
				rs.close
			Next
		End If
	    OutStr=OutStr&"mwritetodocument();"
	    Rs.Close

		file="../jsfiles/menu.js"
		Call EA_Pub.Save_HtmlFile(file,OutStr)


		Dim Level
		List=EA_DBO.Get_Column_List()
		If IsArray(ColumnArray) Then 
			OutStr = "document.write ('<table>');"&chr(10)
			OutStr = OutStr & "document.write ('<form method=""post"" name=""SearchForm"" action="""&SystemFolder&"search.asp?action=query"" target=""_blank"">');"&chr(10)
			OutStr = OutStr & "document.write ('<tr>');"&chr(10)
			OutStr = OutStr & "document.write ('<td align=""center""><span style=""color: #000000;"">站内搜索：</span></td>');"&chr(10)
			OutStr = OutStr & "document.write ('<td align=""center"">&nbsp;');"&chr(10)
			OutStr = OutStr & "document.write ('<select name=""field"">');"&chr(10)
			OutStr = OutStr & "document.write ('<option value=""0"">标题</option>');"&chr(10)
			OutStr = OutStr & "document.write ('<option value=""1"">关键字</option>');"&chr(10)
			OutStr = OutStr & "document.write ('<option value=""2"">作者</option>');"&chr(10)
			OutStr = OutStr & "document.write ('<option value=""3"">摘要</option>');"&chr(10)
			OutStr = OutStr & "document.write ('</select>&nbsp;');"&chr(10)
			OutStr = OutStr & "document.write ('<select name=""column"">');"&chr(10)
			OutStr = OutStr & "document.write ('<option value=""0"">--栏 目--</option>');"&chr(10)
			ForTotal = Ubound(List,2)

			For i=0 To ForTotal
				Level=(Len(List(2,i))/4-1)
				OutStr = OutStr & "document.write ('<option value="""&List(0,i)&"|"&List(2,i)&""">');"&chr(10)
				If Len(List(2,i))>4 Then OutStr = OutStr & "document.write ('├');"&chr(10)
				OutStr = OutStr & "document.write ('"&String(Level,"-")&"');"&chr(10)
				OutStr = OutStr & "document.write ('"&List(1,i)&"</option>');"&chr(10)
			Next
			
			OutStr = OutStr & "document.write ('</select>&nbsp;<input name=""keyword"" type=""text"" value=""关键字"" onfocus=""this.select();"" size=""20"" maxlength=""50"">&nbsp;<input name=""Submit"" type=""submit"" value=""搜索""></td>');"&chr(10)
			OutStr = OutStr & "document.write ('<td align=""center"">&nbsp;<a href="""&SystemFolder&"search.asp"">高级搜索</a></td>');"&chr(10)
			OutStr = OutStr & "document.write ('</tr>');"&chr(10)
			OutStr = OutStr & "document.write ('</form>');"&chr(10)
			OutStr = OutStr & "document.write ('</table>');"&chr(10)
		End if
		
		file="../jsfiles/searchbar.js"
		Call EA_Pub.Save_HtmlFile(file,OutStr)


		OutStr = "document.write ('<table>');"&Chr(10)
		OutStr = OutStr & "document.write ('<tr>');"&Chr(10)
		OutStr = OutStr & "document.write ('<td align=""center""><marquee style=""word-break:break-all;FONT-SIZE: 9pt; LEFT: 2px; MARGIN-LEFT: 2px; WIDTH: 100%; TOP: 2px; HEIGHT: 100px; TEXT-ALIGN: center"" onMouseOver=this.stop() onMouseOut=this.start() scrollamount=1 scrolldelay=50 direction=up behavior=loop>');"&Chr(10)

		SQL="Select Top 8 LinkURL,LinkImgPath,LinkName,LinkInfo From [NB_FriendLink] Where ColumnId=0 And State="&EA_M_DBO.TrueValue&" And Style=1 Order By OrderNum Desc"
		Set Rs=Conn.Execute(SQL)
		If Not rs.EOF And Not rs.BOF Then
			List=rs.getrows()

			For i=0 To UBound(List,2)
				OutStr = OutStr & "document.write ('<a href="""&List(0,i)&""" target=_blank title="""&List(3,i)&"""><img src="""&List(1,i)&""" align=""absmiddle"" width=""88"" height=""31"" src="""&List(3,i)&"""></a><br>');"&Chr(10)
			Next
		End If
		OutStr = OutStr & "document.write ('</marquee></td>');"&Chr(10)
		OutStr = OutStr & "document.write ('</tr>');"&Chr(10)
		OutStr = OutStr & "document.write ('<tr><td align=""center"" height=""5""></td></tr>');"&Chr(10)
		OutStr = OutStr & "document.write ('<tr><td align=""center"">');"&Chr(10)
		OutStr = OutStr & "document.write ('<select name=""textfriend"" onChange=""if(this.selectedIndex) window.open(this.options[this.selectedIndex].value);"" style=""width:150"">');"&Chr(10)
		OutStr = OutStr & "document.write ('<option value="""">--文字连接站点--</option>');"&Chr(10)

		SQL="Select Top 10 LinkURL,LinkName From [NB_FriendLink] Where ColumnId=0 And State="&EA_M_DBO.TrueValue&" And Style=0 Order By OrderNum Desc"
		Set Rs=Conn.Execute(SQL)
		If Not rs.EOF And Not rs.BOF Then
			List=rs.getrows()

			For i=0 To UBound(List,2)
				OutStr = OutStr & "document.write ('<option value="""&List(0,i)&""">"&List(1,i)&"</option>');"&Chr(10)
			Next
		End If
		OutStr = OutStr & "document.write ('</select></td>');"&Chr(10)
		OutStr = OutStr & "document.write ('</tr>');"&Chr(10)
		OutStr = OutStr & "document.write ('<tr><td align=""center"" height=""5""></td></tr>');"&Chr(10)
		OutStr = OutStr & "document.write ('<tr><td align=""center"" height=""25""><a href=""#"" onclick=""javascript:window.open(\'"&SystemFolder&"app_link.asp\',\'\',\'height=320,width=550\')"">申请连接</a>&nbsp;&nbsp;<a href="""&SystemFolder&"morelink.asp"" target=""_blank"">更多连接</a></td>');"&Chr(10)
		OutStr = OutStr & "document.write ('</tr>');"&Chr(10)
		OutStr = OutStr & "document.write ('</table>');"&Chr(10)
		
		file="../jsfiles/friend.js"
		Call EA_Pub.Save_HtmlFile(file,OutStr)

		Response.Write "1"
End Sub
%>