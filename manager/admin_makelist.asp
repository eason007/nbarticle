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
'= 文件名称：/Manager/Admin_MakeList.asp
'= 摘    要：后台-HTML栏目页生成文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2007-10-17
'====================================================================

Server.ScriptTimeout=9999999

Call EA_Manager.Chk_IsMaster

If Not EA_Manager.Chk_Power(Admin_Power,"42") Then 
	ErrMsg=str_Comm_NotAccess
	Call EA_Manager.Error(1)
End If

Dim Atcion,PostId
Dim ChildColumnList
Dim SysKeyword, SysDesc
Dim ForTotal

Atcion=Request.Form ("action")

Select Case LCase(Atcion)
Case "mark"
	SysKeyword	= EA_Pub.SysInfo(16)
	SysDesc		= EA_Pub.SysInfo(17)

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
	ColumnList = "(build-select),0 " & str_Comm_AllColumn & ",0"
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

	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_MakeList_Help",str_MakeList_Help)

	Call EA_M_XML.AppElements("Language_Comm_SelectAll",str_Comm_SelectAll)
	Call EA_M_XML.AppElements("Language_Comm_Select",str_Comm_Select)

	Call EA_M_XML.AppElements("Language_MakeList_Title",str_MakeList_Title)

	Call EA_M_XML.AppElements("Language_MakeList_Option_1",str_MakeList_Option_1)
	Call EA_M_XML.AppElements("Language_MakeList_Option_2",str_MakeList_Option_2)
	Call EA_M_XML.AppElements("Language_MakeList_Option_3",str_MakeList_Option_3)
	Call EA_M_XML.AppElements("Language_MakeList_Option_4",str_MakeList_Option_4)
	Call EA_M_XML.AppElements("Language_MakeList_Option_5",str_MakeList_Option_5)
	Call EA_M_XML.AppElements("Language_MakeList_Option_6",str_MakeList_Option_6)
	Call EA_M_XML.AppElements("Language_MakeList_Option_7",str_MakeList_Option_7)
	Call EA_M_XML.AppElements("Language_MakeList_Option_8",str_MakeList_Option_8)

	Call EA_M_XML.AppElements("btnSubmit",str_MakeIndex_StartNow)

	Page = EA_M_XML.make("","",0)
	Call EA_M_XML.Out(Page)
End Sub

Sub MarkList
	Dim SelType,ColumnList,PostId
	Dim k,l,TmpStr
	
	PostId	= EA_Pub.SafeRequest(2,"ColumnList",0,0,0)
	SelType	= EA_Pub.SafeRequest(2,"seltype",1,"",0)
	
	SQL="Select [Id] from [NB_Column]"
	If PostId<>0 Then SQL=SQL&" Where [Id]="&PostId
	Set Rs=Conn.Execute(SQL)
	
	If Not rs.eof And Not rs.bof Then 
		ColumnList=Rs.GetRows()

		PageContent=EA_Temp.LoadTemplate("admin_makelist_view.htm")

		Template.SetVariable "ColumnTotal",Ubound(ColumnList,2)+1,PageContent
		Template.SetVariable "Language_MakeList_Column",str_MakeList_Column,PageContent
		Template.SetVariable "Language_MakeList_Now",str_MakeList_Now,PageContent
		Template.SetVariable "Language_MakeList_Task",str_MakeList_Task,PageContent
		Template.SetVariable "Language_MakeList_Page",str_MakeList_Page,PageContent

		Template.BaseReplace PageContent
		Response.Write PageContent

		Set Template = Nothing
		
		ForTotal = Ubound(ColumnList,2)

		For k=0 To ForTotal
			TmpStr=Split(SelType,", ")
			
			Response.Write "<script>task_total.innerHTML="""&UBound(TmpStr)+1&""";</script>" & VbCrLf
			For l=0 To UBound(TmpStr)
				Response.Write "<script>page_total.innerHTML=""1"";</script>" & VbCrLf
				Response.Write "<script>page_complete.innerHTML=""1"";</script>" & VbCrLf
				Response.Write "<script>img3.width=1;</script>" & VbCrLf
					
				Select Case TmpStr(l)
				Case "1"
					MakeColumn ColumnList(0,k),"adddate","desc"
				Case "2"
					MakeColumn ColumnList(0,k),"adddate","asc"
				Case "3"
					MakeColumn ColumnList(0,k),"title","desc"
				Case "4"
					MakeColumn ColumnList(0,k),"title","asc"
				Case "5"
					MakeColumn ColumnList(0,k),"viewnum","desc"
				Case "6"
					MakeColumn ColumnList(0,k),"viewnum","asc"
				Case "7"
					MakeColumn ColumnList(0,k),"commentnum","desc"
				Case "8"
					MakeColumn ColumnList(0,k),"commentnum","asc"
				End Select
				
				If UBound(TmpStr)=0 Then
					Response.Write "<script>img2.width=400;" & VbCrLf
				Else
					Response.Write "<script>img2.width=" & Fix((l/UBound(TmpStr)) * 400) & ";" & VbCrLf
				End If
				Response.Write "task_complete.innerHTML=""<font color=green>"&l+1&"</font>"";</script>" & VbCrLf
				Response.Flush
			Next
			
			If Ubound(ColumnList,2)=0 Then
				Response.Write "<script>img1.width=400;" & VbCrLf
			Else
				Response.Write "<script>img1.width=" & Fix((k/Ubound(ColumnList,2)) * 400) & ";" & VbCrLf
			End If
			Response.Write "column_complete.innerHTML=""<font color=blue>"&k+1&"</font>"";</script>" & VbCrLf
			Response.Flush
		Next
	End If
	
	Response.Write "<script>make_msg.innerHTML="""&str_MakeList_AllComplate&""";</script>" & VbCrLf
	
	Rs.Close
	Set Rs=Nothing
End Sub

Function MakeColumn(ColumnId,Field,Order)
	Dim TopicNav,PageNumNav
	Dim PageContent
	Dim PageCount,PageSize
	Dim Tmp,i,j,TempStr
	Dim ArticleList,ColumnInfo,ArticleOutStr
	Dim FileName,Folder,FileNameExample
	Dim re
	Dim Temp,ListBlock
	Dim Template

	Set re=New RegExp
	re.IgnoreCase =true
	re.Global=True

	FileName=EA_Pub.Cov_ColumnPath(ColumnId, "0")

	re.Pattern=Replace(SystemFolder, "/", "\/") & "(.*)\/(\w+).(\w+)"
	Folder = re.Replace(FileName,"/$1/")
	Folder = Replace(Folder, "adddate", Field)
	Folder = Replace(Folder, "desc", Order)

	re.Pattern="(.*)_(\d+).(\w+)"
	FileNameExample = re.Replace(FileName,"$1")

	If Not(EA_Pub.CheckDir(".." & Folder)) Then 
		Tmp = Split(Folder, "/")
		TempStr = ""
		ForTotal = UBound(Tmp)-1

		For j = 1 To ForTotal
			TempStr = TempStr & "/" & Tmp(j)

			If Not(EA_Pub.CheckDir(".." & TempStr)) Then EA_Pub.MakeNewsDir Server.MapPath(".." & TempStr)
		Next
	End If
	
	'load column data
	ColumnInfo=EA_DBO.Get_Column_Info(ColumnId)
	
	Set EA_Temp=New cls_Template
	
	PageSize=ColumnInfo(17,0)
	
	PageContent=EA_Temp.Load_Template(ColumnInfo(9,0),"list")

	PageCount=EA_Pub.Stat_Page_Total(PageSize,ColumnInfo(3,0))
	
	If PageCount=0 Then
		Response.Write "<script>page_total.innerHTML=""1"";</script>" & VbCrLf
	Else
		Response.Write "<script>page_total.innerHTML="""&PageCount&""";</script>" & VbCrLf
	End If
	
	EA_Temp.Title=ColumnInfo(0,0)&" - "&EA_Pub.SysInfo(0)
	EA_Temp.Nav="<a href="""&SystemFolder&"""><b>"&EA_Pub.SysInfo(0)&"</b></a> - "&EA_Pub.Get_NavByColumnCode(ColumnInfo(1,0))

	'jump to out url
	If ColumnInfo(6,0) Then 
		PageContent="<meta http-equiv=""refresh"" content=""0;URL="&ColumnInfo(7,0)&""">"

		FileName = Replace(FileName, "adddate", Field)
		FileName = Replace(FileName, "desc", Order)
		
		Call EA_Pub.Save_HtmlFile(FileName,PageContent)
		Exit Function	
	End If
		
	TopicNav="<table>"&Chr(10)
	TopicNav=TopicNav&"<tr>"&Chr(10)
	TopicNav=TopicNav&"<td style=""width: 65%;"">"
	If Field="Title" And Order="Desc" Then 
		TopicNav=TopicNav&"<a href="""&Replace(Replace(FileName, "adddate", "title"), "desc", "asc")&""" title=""按文章标题升序排列"">↓"
	Else
		TopicNav=TopicNav&"<a href="""&Replace(FileName, "adddate", "title")&""" title=""按文章标题降序排列"">"
		If Field="Title" Then TopicNav=TopicNav&"↑"
	End If
	TopicNav=TopicNav&"<strong>标题</strong></a></td>"&Chr(10)
	TopicNav=TopicNav&"<td style=""width: 20%;"">"
	If Field="AddDate" And Order="Desc" Then 
		TopicNav=TopicNav&"<a href="""&Replace(FileName, "desc", "asc")&""" title=""按文章发布日期升序排列"">↓"
	Else
		TopicNav=TopicNav&"<a href="""&FileName&""" title=""按文章发布日期降序排列"">"
		If Field="AddDate" Then TopicNav=TopicNav&"↑"
	End If
	TopicNav=TopicNav&"<strong>发布日期</strong></a></td>"&Chr(10)
	TopicNav=TopicNav&"<td style=""width: 15%;"">"
	If Field="ViewNum" And Order="Desc" Then 
		TopicNav=TopicNav&"<a href="""&Replace(Replace(FileName, "adddate", "viewnum"), "desc", "asc")&""" title=""按浏览人数升序排列"">↓"
	Else
		TopicNav=TopicNav&"<a href="""&Replace(FileName, "adddate", "viewnum")&""" title=""按浏览人数降序排列"">"
		If Field="ViewNum" Then TopicNav=TopicNav&"↑"
	End If
	TopicNav=TopicNav&"<strong>浏览</strong></a><strong>/</strong>"
	If Field="CommentNum" And Order="Desc" Then 
		TopicNav=TopicNav&"<a href="""&Replace(Replace(FileName, "adddate", "commentnum"), "desc", "asc")&""" title=""按回复人数升序排列"">↓"
	Else
		TopicNav=TopicNav&"<a href="""&Replace(FileName, "adddate", "commentnum")&""" title=""按回复人数降序排列"">"
		If Field="CommentNum" Then TopicNav=TopicNav&"↑"
	End If
	TopicNav=TopicNav&"<strong>回复</strong></a></td>"&Chr(10)
	TopicNav=TopicNav&"</tr>"
	TopicNav=TopicNav&"</table>"

	PageContent=Replace(PageContent,"{$ColumnId$}",ColumnId)
	PageContent=Replace(PageContent,"{$ColumnName$}",ColumnInfo(0,0))
	PageContent=Replace(PageContent,"{$ColumnInfo$}",ColumnInfo(2,0))
	PageContent=Replace(PageContent,"{$ColumnTopicTotal$}",ColumnInfo(3,0))
	PageContent=Replace(PageContent,"{$ColumnMangerTotal$}",ColumnInfo(4,0))
	PageContent=Replace(PageContent,"{$ColumnTitleNav$}",TopicNav)
	PageContent=Replace(PageContent,"{$ColumnNav$}",SiteColumnNav(ColumnId,ColumnInfo(1,0)))
	PageContent=Replace(PageContent,"{$ChildColumn$}",ColumnChild(ColumnInfo(1,0)))

	EA_Temp.Find_TemplateTagByInput "ChildColumnNav",ChildColumnNav(PageContent),PageContent

	EA_Pub.SysInfo(16)=ColumnInfo(0,0)&","&SysKeyword
	If Len(ColumnInfo(2,0)) Then EA_Pub.SysInfo(17)=ColumnInfo(2,0)
	
	PageContent=EA_Temp.Replace_PublicTag(PageContent)

	Call EA_Temp.Find_TemplateTags("Friend",PageContent)
	
	If PageCount=0 Then 
		Set Template=New cls_NEW_TEMPLATE

		ListBlock=Template.GetBlock("list",PageContent)

		PageNumNav=PageList(0,0,Replace(Replace(FileNameExample, "adddate", Field), "desc", Order) & "_")

		PageContent=Replace(PageContent,"{$ColumnPageNumNav$}",PageNumNav)

		FileName = Replace(FileName, "adddate", Field)
		FileName = Replace(FileName, "desc", Order)
		
		Call EA_Pub.Save_HtmlFile(FileName,PageContent)
		
		Exit Function 
	End If

	If Not IsObject(Template) Then Set Template=New cls_NEW_TEMPLATE
	
	For j=1 To PageCount
		Response.Write "<script>img3.width=" & Fix((j/PageCount) * 400) & ";" & VbCrLf
		Response.Write "page_complete.innerHTML=""<b>"&j&"</b>"";</script>" & VbCrLf
		Response.Flush

		'load article list
		If ColumnInfo(3,0)>0 Then 
			If Rs.State=1 Then Rs.Close
			
			SQL="SELECT [Id], TColor, Title, AddDate, CommentNum, Summary, LastComment, ViewNum, IsImg, Img, IsTop, Author, AuthorId, [KeyWord]"
			SQL=SQL&"FROM NB_Content "
			SQL=SQL&"WHERE ColumnId="&ColumnId&" And IsPass="&EA_DBO.TrueValue&" And IsDel=0 "
			SQL=SQL&"ORDER BY "
			If Field="AddDate" Then 
				SQL=SQL&"TrueTime"
			Else
				SQL=SQL&Field
			End If
			SQL=SQL&" "&Order
			Rs.Open SQL,Conn,1,1

			If Not rs.eof And Not rs.bof Then 
				Rs.AbsolutePosition=Rs.AbsolutePosition+((j-1)*PageSize)
				ArticleList=Rs.GetRows(PageSize)
			End If
		End If

		TempStr=PageContent
		ListBlock=Template.GetBlock("list",TempStr)

		If IsArray(ArticleList) Then
			ForTotal = UBound(ArticleList,2)

			For i=0 To ForTotal
				Temp=ListBlock
  
				Template.SetVariable "Url",EA_Pub.Cov_ArticlePath(ArticleList(0,i),ArticleList(3,i),EA_Pub.SysInfo(18)),Temp
				Template.SetVariable "Title",EA_Pub.Add_ArticleColor(ArticleList(1,i),EA_Pub.Base_HTMLFilter(ArticleList(2,i))),Temp
				Template.SetVariable "Date",ArticleList(3,i),Temp
				Template.SetVariable "CommentNum",ArticleList(4,i),Temp
				Template.SetVariable "Summary",ArticleList(5,i),Temp
				Template.SetVariable "LastComment",ArticleList(6,i),Temp
				Template.SetVariable "ViewNum",ArticleList(7,i),Temp
				Template.SetVariable "Icon",EA_Pub.Chk_ArticleType(ArticleList(8,i),ArticleList(10,i)),Temp
				Template.SetVariable "Img",ArticleList(9,i),Temp
				Template.SetVariable "Author","<a href='"&SystemFolder&"florilegium.asp?a_name="&ArticleList(11,i)&"&a_id="&ArticleList(12,i)&"' rel=""external"">"&ArticleList(11,i)&"</a>",Temp
				Template.SetVariable "Tag",TagList(ArticleList(13,i)),Temp

				Template.SetBlock "list",Temp,TempStr
			Next
			Template.CloseBlock "list",TempStr
		End If

		PageNumNav=PageList(j,PageCount,Replace(Replace(FileNameExample, "adddate", Field), "desc", Order) & "_")
		
		TempStr=Replace(TempStr,"{$ColumnTopicList$}",ArticleOutStr)
		TempStr=Replace(TempStr,"{$ColumnPageNumNav$}",PageNumNav)

		Tmp = Replace(FileName, "adddate", Field)
		Tmp = Replace(Tmp, "desc", Order)
		Tmp = Replace(Tmp, "_1.", "_" & j & ".")

		Call EA_Pub.Save_HtmlFile(Tmp,TempStr)
	Next

	Response.Write "<script>img3.width=400;</script>" & VbCrLf
End Function

Function PageList(iCurrentPage,iPageCount,Http)
	Dim OutStr
	Dim PageRoot, PageFoot
	Dim i

		If CLng(iCurrentPage)<=0 Then 
			iCurrentPage=1
		ElseIf CLng(iCurrentPage)>CLng(iPageCount) Then
			iCurrentPage=iPageCount
		End if
		
		If iCurrentPage-4<=1 Then 
			PageRoot=1
		Else
			PageRoot=iCurrentPage-4
		End If	
		If iCurrentPage+4>=iPageCount Then 
			PageFoot=iPageCount
		Else
			PageFoot=iCurrentPage+4
		End If

	OutStr="<div id=""pageList"">"

	If Clng(iCurrentPage) > 1 Then 
		OutStr=OutStr&"<a href='"
		OutStr=OutStr&Http&"1.htm' title=""Go to first page"" class=""first"">"
		OutStr=OutStr&"&laquo;</a>&nbsp;"
		OutStr=OutStr&"<a href='"
		OutStr=OutStr&Http&iCurrentPage-1&".htm' title=""Go to previous page"" class=""list"">"
		OutStr=OutStr&"&lt;</a>&nbsp;"
	End If

	For i=PageRoot To PageFoot
		If i=Cint(iCurrentPage) Then
			OutStr=OutStr&"<span class=""current"">"&i&"</span>&nbsp;"
		Else
			OutStr=OutStr&"<a href="""&Http&i&".htm"
			OutStr=OutStr&""" title="""&i&""" class=""list"">"&i&"</a>&nbsp;"
		End If
		If i=iPageCount Then Exit For
	Next

	If Clng(iCurrentPage) < iPageCount Then 
		OutStr=OutStr&"<a href='"
		OutStr=OutStr&Http&iCurrentPage+1&".Htm' title=""Go to next page"" class=""list"">"
		OutStr=OutStr&"&gt;</a>&nbsp;"
		OutStr=OutStr&"<a href='"
		OutStr=OutStr&Http&iPageCount&".Htm' title=""Go to last page"" class=""last"">"
		OutStr=OutStr&"&raquo;</a>&nbsp;"
	End If
	
	OutStr=OutStr&"<span class=""total"">"&Clng(iCurrentPage)&"/"&iPageCount&"</span>"

	OutStr=OutStr&"</div>"

	PageList=OutStr
End Function
%>