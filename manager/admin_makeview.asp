<!--#Include File="comm/inc.asp" -->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2004 - 2008 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：/Manager/Admin_MakeView.asp
'= 摘    要：后台-HTML栏目页生成文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2007-01-20
'====================================================================

Server.ScriptTimeout=9999999

Call EA_Manager.Chk_IsMaster

If Not EA_Manager.Chk_Power(Admin_Power,"43") Then 
	ErrMsg=str_Comm_NotAccess
	Call EA_Manager.Error(1)
End If

Dim Atcion
Dim ForTotal
Dim PageIndex(),PageStr()

Atcion=Request.Form("action")

Select Case LCase(Atcion)
Case "mark"
	Set EA_Temp=New cls_Template
			
	Call MarkView
Case Else
	Call Main
End Select

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
	Call EA_M_XML.AppElements("Language_MakeView_Help",str_MakeView_Help)

	Call EA_M_XML.AppElements("Language_MakeView_MakeById",str_MakeView_MakeById)
	Call EA_M_XML.AppElements("Language_MakeView_StartId",str_MakeView_StartId)
	Call EA_M_XML.AppElements("Language_MakeView_EndId",str_MakeView_EndId)

	Call EA_M_XML.AppElements("Language_MakeView_MakeByDate",str_MakeView_MakeByDate)
	Call EA_M_XML.AppElements("Language_MakeView_StartDate",str_MakeView_StartDate)
	Call EA_M_XML.AppElements("Language_MakeView_EndDate",str_MakeView_EndDate)

	Call EA_M_XML.AppElements("Language_MakeView_MakeByColumn",str_MakeView_MakeByColumn)

	Call EA_M_XML.AppElements("btnSubmit1",str_MakeIndex_StartNow)
	Call EA_M_XML.AppElements("btnSubmit2",str_MakeIndex_StartNow)
	Call EA_M_XML.AppElements("btnSubmit3",str_MakeIndex_StartNow)

	Page = EA_M_XML.make("","",0)
	Call EA_M_XML.Out(Page)
End Sub

Sub MarkView
	Dim TopicList,i,j,k
	Dim Tag
	Dim StartBorder,EndBorder
	Dim WSQL
	
	Tag=EA_Pub.SafeRequest(2,"tag",0,0,0)
	
	Select Case Tag
	Case 1
		StartBorder=EA_Pub.SafeRequest(2,"sid",0,0,0)
		EndBorder=EA_Pub.SafeRequest(2,"eid",0,0,0)
		
		WSQL=" And a.[id] between "&StartBorder&" and "&EndBorder
	Case 2
		StartBorder=EA_Pub.SafeRequest(2,"date1",2,"1970-1-1",0)
		EndBorder=EA_Pub.SafeRequest(2,"date2",2,Date(),0)
		
		If iDataBaseType=0 Then
			WSQL=" And AddDate between #"&StartBorder&"# and #"&DateAdd("d",1,EndBorder)&"#"
		Else
			WSQL=" And AddDate between '"&StartBorder&"' and '"&DateAdd("d",1,EndBorder)&"'"
		End If
	Case 3
		StartBorder=EA_Pub.SafeRequest(2,"ColumnList",0,0,0)
		
		If StartBorder<>0 Then WSQL=" And ColumnId="&StartBorder
	Case 4
		StartBorder=EA_Pub.SafeRequest(2,"sid",1,"",0)

		If Len(StartBorder)>0 Then WSQL=" And a.[id] In ("&StartBorder&")"
	End Select
		
	'0=ColumnId,1=ColumnCode,2=ColumnName,3=Title,4=Summary,5=Content,6=ViewNum,7=AuthorId,8=Author,9=CommentNum,10=IsOut
	'11=OutUrl,12=[KeyWord],13=AddDate,14=CutArticle,15=Source,16=SourceUrl,17=TColor,18=Img,19=IsTop,20=ListPower
	'21=IsHide,22=Article_TempId,23=[Id],24=TrueTime
	SQL="Select ColumnId,ColumnCode,ColumnName,a.Title,Summary,Content,a.ViewNum,AuthorId,Author,CommentNum,a.IsOut"
	SQL=SQL&",a.OutUrl,[KeyWord],AddDate,CutArticle,Source,SourceUrl,TColor,Img,a.IsTop,b.ListPower,b.IsHide,b.Article_TempId,a.[Id],TrueTime"
	SQL=SQL&" FROM NB_Content AS a INNER JOIN NB_Column AS b ON a.ColumnId=b.Id Where IsPass="&EA_M_DBO.TrueValue&" And IsDel=0 "&WSQL&" Order By ColumnId Asc,a.[Id] Asc"
	'response.write sql
	'response.end
	Set Rs=Conn.Execute(SQL)
	If Not rs.eof And Not rs.bof Then 
		TopicList=Rs.GetRows()
		Rs.Close
		
		Dim Template
		Set Template=New cls_NEW_TEMPLATE

		PageContent=Template.LoadTemplate("admin_makeview_view.htm")

		Template.SetVariable "PageTotal",Ubound(TopicList,2)+1,PageContent
		Template.SetVariable "Language_MakeList_Now",str_MakeList_Now,PageContent
		Template.SetVariable "Language_MakeList_Page",str_MakeList_Page,PageContent

		Template.BaseReplace PageContent
		Response.Write PageContent

		Set Template = Nothing

		PageContent = ""

		Dim CurrentTemplateId,CurrentColumnId
		Dim IsReplace
		Dim PageContent,ArticleContent,PageKeyword
		Dim TempStr,TempArray
		Dim Folder,sHTMLFilePath
		Dim FirstArticle,NextArticle
		Dim NewFolderList
		Dim re
		Dim Tmp

		Set re=New RegExp
		re.IgnoreCase =true
		re.Global=True

		NewFolderList = ","

		PageKeyword = EA_Pub.SysInfo(16)
		ForTotal = UBound(TopicList,2)
		k=0

		For i=0 To ForTotal
			IsReplace=True
			sHTMLFilePath=EA_Pub.Cov_ArticlePath(TopicList(23,i), TopicList(13,i), "0")
			
			'check folder isexists
			re.Pattern=Replace(SystemFolder, "/", "\/") & "(.*)\/(\w+).(\w+)"
			Folder = re.Replace(sHTMLFilePath,"/$1/")

			If InStr(NewFolderList, "," & Folder & ",") = 0 Then
				If Not(EA_Pub.CheckDir(".." & Folder)) Then 
					TempArray = Split(Folder, "/")
					Tmp = ""

					For j = 1 To UBound(TempArray)-1
						Tmp = Tmp & "/" & TempArray(j)

						If InStr(NewFolderList, "," & Tmp & "/,") = 0 Then
							If Not(EA_Pub.CheckDir(".." & Tmp)) Then EA_Pub.MakeNewsDir Server.MapPath(".." & Tmp)
							NewFolderList = NewFolderList & Tmp & "/,"
						End If
					Next
				End If
			End If
			
			'check is member
			If TopicList(20,i)>0 Or TopicList(21,i)<>0 Then
				IsReplace=False
				TempStr="<meta http-equiv=""refresh"" content=""0;URL="&SystemFolder&"article.asp?articleid="&TopicList(23,i)&""">"
			End If
			
			'check is out
			If TopicList(10,i) And IsReplace Then
				IsReplace=False
				TempStr="<meta http-equiv=""refresh"" content=""0;URL="&TopicList(11,i)&""">"
			End If
			
			'load template
			If CurrentTemplateId<>TopicList(22,i) And IsReplace Then
				CurrentTemplateId=TopicList(22,i)
				PageContent=EA_Temp.Load_Template(CurrentTemplateId,"view")
			End If

			If Len(PageContent)<=0 And IsReplace Then
				Response.Write "栏目[ "&TopicList(2,i)&" ]尚未设定文章页模版，导致文章："&TopicList(3,i)&" 未能生成HTML。<br>"
				IsReplace = False
			End If

			If CurrentColumnId<>TopicList(0,i) Then
				EA_Temp.Nav="<a href="""&SystemFolder&"""><b>"&EA_Pub.SysInfo(0)&"</b></a>"&EA_Pub.Get_NavByColumnCode(TopicList(1,i))&" -=> 正文"
				CurrentColumnId=TopicList(0,i)
			End If
			
			If Not IsReplace Then
			'not replace template tag
				Call EA_Pub.Save_HtmlFile(sHTMLFilePath,TempStr)
			Else
				TempStr=PageContent
				
				EA_Temp.Title=TopicList(3,i)&" - "&TopicList(2,i)&" - "&EA_Pub.SysInfo(0)

				EA_Temp.ReplaceTag "ColumnId",TopicList(0,i),TempStr
				EA_Temp.ReplaceTag "ArticleId",TopicList(23,i),TempStr
				EA_Temp.ReplaceTag "ArticleTitle",EA_Pub.Add_ArticleColor(TopicList(17,i),TopicList(3,i)),TempStr
				EA_Temp.ReplaceTag "ArticlePostTime",TopicList(13,i),TempStr
				EA_Temp.ReplaceTag "ArticleSummary",TopicList(4,i),TempStr
				
				EA_Temp.ReplaceTag "ArticleAuthor","<a href='"&SystemFolder&"florilegium.asp?a_name="&TopicList(8,i)&"&amp;a_id="&TopicList(7,i)&"' rel=""external"">"&TopicList(8,i)&"</a>",TempStr
				
				If Len(TopicList(16,i))>0 Then
					EA_Temp.ReplaceTag "ArticleFrom","<a href='"&TopicList(16,i)&"' rel=""external"">"&TopicList(15,i)&"</a>",TempStr
				Else
					EA_Temp.ReplaceTag "ArticleFrom","本站",TempStr
				End If
				
				EA_Temp.ReplaceTag "ArticleViewTotal","<script type=""text/javascript"" src="""&SystemFolder&"articleinfo.asp?action=viewtotal&amp;articleid="&TopicList(23,i)&"""></script>",TempStr
				EA_Temp.ReplaceTag "ArticleCommentTotal","<script type=""text/javascript"" src="""&SystemFolder&"articleinfo.asp?action=commenttotal&amp;articleid="&TopicList(23,i)&"""></script>",TempStr

				If InStr(TempStr, "{$FirstArticle$}") > 0 Then
					FirstArticle=EA_DBO.Get_Article_FirstArticle(TopicList(0,i),TopicList(24,i),TopicList(23,i))

					If IsArray(FirstArticle) Then
						TempStr=Replace(TempStr,"{$FirstArticle$}","<a href='"&EA_Pub.Cov_ArticlePath(FirstArticle(0,0),FirstArticle(3,0),EA_Pub.SysInfo(18))&"' rel=""external"">"&EA_Pub.Add_ArticleColor(FirstArticle(2,0),FirstArticle(1,0))&"</a>")
					Else
						TempStr=Replace(TempStr,"{$FirstArticle$}","<span style=""color: #800000;"">已到尽头</span>")
					End If
				End If

				If InStr(TempStr, "{$NextArticle$}") > 0 Then
					NextArticle=EA_DBO.Get_Article_NextArticle(TopicList(0,i),TopicList(24,i),TopicList(23,i))

					If IsArray(NextArticle) Then
						TempStr=Replace(TempStr,"{$NextArticle$}","<a href='"&EA_Pub.Cov_ArticlePath(NextArticle(0,0),NextArticle(3,0),EA_Pub.SysInfo(18))&"' rel=""external"">"&EA_Pub.Add_ArticleColor(NextArticle(2,0),NextArticle(1,0))&"</a>")
					Else
						TempStr=Replace(TempStr,"{$NextArticle$}","<span style=""color: #800000;"">已到尽头</span>")
					End If
				End If
				
				EA_Pub.SysInfo(16)=TopicList(12,i) & "," & PageKeyword 
				EA_Pub.SysInfo(17)=TopicList(4,i)

				Call CorrList(TopicList(12,i)&"",TopicList(0,i),TopicList(23,i),TempStr)
				Call TagList(TopicList(12,i),TempStr)

				TempStr=EA_Temp.Replace_PublicTag(TempStr)

				TopicList(5,i)=EA_Pub.Cov_InsideLink(TopicList(5,i),TopicList(0,i))
				
				Call RegExpTest("\[NextPage([^\]])*\]", TopicList(5,i), re)

				If UBound(PageIndex) = 1 Then
					EA_Temp.ReplaceTag "ArticleText","<div id=""article"">"&TopicList(5,i)&"</div>",TempStr

					Call EA_Pub.Save_HtmlFile(sHTMLFilePath,TempStr)
				Else
					For j = 1 To UBound(PageIndex)
						Tmp = TempStr

						ArticleContent = Mid(TopicList(5,i), PageIndex(j - 1) + Len(PageStr(j - 1)) + 1, PageIndex(j) - PageIndex(j - 1) - Len(PageStr(j - 1)))
						ArticleContent = "<div id=""article"">" & ArticleContent & "</div>"
						ArticleContent = ArticleContent & "<div style='TEXT-ALIGN: center;margin-bottom: 5px;'>" & PageNav(UBound(PageIndex), j, sHTMLFilePath, re) & "</div>"

						EA_Temp.ReplaceTag "ArticleText",ArticleContent,Tmp

						re.Pattern = "(.*)\/(\w+)_(\d+).(\w+)"
						
						If j = 1 Then
							Call EA_Pub.Save_HtmlFile(sHTMLFilePath,Tmp)
						Else
							Call EA_Pub.Save_HtmlFile(re.Replace(sHTMLFilePath,"$1/$2_$3_" & j & ".$4"),Tmp)
						End If
					Next
				End If
			End If
			
			k = k+1
			
			If ForTotal=0 Then 
				Response.Write "<script>img1.width=400;" & VbCrLf
			Else
				Response.Write "<script>img1.width=" & Fix((k/ForTotal) * 400) & ";" & VbCrLf
			End If
			Response.Write "column_complete.innerHTML=""<font color=green>"&i+1&"</font>"";</script>" & VbCrLf
			Response.Flush
			'response.end
		Next
		
		Response.Write "<script>img1.width=400;"& VbCrLf
		Response.Write "make_msg.innerHTML="""&str_MakeList_AllComplate&""";</script>" & VbCrLf
	End If
End Sub

Sub TagList (Keyword, ByRef PageContent)
	Dim OutStr
	Dim TempArray,i
	Dim ForTotal

	If Len(Trim(Keyword)) > 0 And Not IsNull(Keyword) Then
		TempArray= Split(Keyword,",")

		ForTotal = UBound(TempArray)

		For i=0 To ForTotal
			If Len(Trim(TempArray(i))) > 0 And Not IsNull(TempArray(i)) Then OutStr = OutStr & "<a href='" & SystemFolder & "search.asp?action=query&amp;field=1&amp;keyword=" & server.urlencode(Trim(TempArray(i))) & "' rel='external'>" & Trim(TempArray(i)) & "</a>&nbsp;"
		Next
	End If

	Call EA_Temp.ReplaceTag("TagList",OutStr,PageContent)
End Sub

Sub CorrList(Keyword,ColumnId,ArticleId,ByRef PageContent)
	Dim ConfigParameterArray
	Dim TempStr

	If Keyword = "" Or Len(Keyword) = 0 Then 
		TempStr = EA_Temp.Text_List(ConfigParameterArray,0,0,0,1,0,0,0,0,0)
	Else
		ConfigParameterArray=EA_Temp.Find_TemplateTagValues("CorrList",PageContent)

		If IsArray(ConfigParameterArray) Then 
			If UBound(ConfigParameterArray) < 8 Then 
				ReDim Preserve ConfigParameterArray(8)
				ConfigParameterArray(8) = "5"
			End If

			Dim TempArray,i,SearchKeyWord
			Dim ForTotal

			TempArray= Split(Keyword,",")
			ForTotal = UBound(TempArray)

			For i=0 To ForTotal
				Select Case iDataBaseType
				Case 0
					SearchKeyWord=SearchKeyWord&" InStr(','+keyword+',',',"&TempArray(i)&",')>0 or "
				Case 1
					SearchKeyWord=SearchKeyWord&" CharIndex(',"&TempArray(i)&",',','+keyword+',')>0 or "
				End Select
			Next
		
			TempArray=EA_DBO.Get_Article_CorrList(SearchKeyWord,ArticleId,ColumnId,CInt(ConfigParameterArray(8)))

			TempStr=EA_Temp.Text_List(TempArray,CInt(ConfigParameterArray(0)),CInt(ConfigParameterArray(1)),CInt(ConfigParameterArray(2)),CInt(ConfigParameterArray(3)),CInt(ConfigParameterArray(4)),CInt(ConfigParameterArray(5)),CInt(ConfigParameterArray(6)),CInt(ConfigParameterArray(7)),CInt(ConfigParameterArray(8)))
		End If
	End If

	Call EA_Temp.Find_TemplateTagByInput("CorrList",TempStr,PageContent)
End Sub

Sub RegExpTest(patrn, strng, ByRef regEx) 
	Dim Match, Matches			' 建立变量。 
	Dim i

	regEx.Pattern = patrn				' 设置模式。 
	Set Matches = regEx.Execute(strng)	' 执行搜索。 

	ReDim PageIndex(Matches.Count + 1)
	ReDim PageStr(Matches.Count + 1)

	i = 1
	
	PageIndex(0) = 0

	For Each Match in Matches			' 遍历匹配集合。 
		PageIndex(i) = Match.FirstIndex
		PageStr(i)	 = Match.Value

		i = i + 1
	Next

	PageIndex(i) = Len(strng)
End Sub

Function PageNav (iCount, iCurrentPage, sFilePath, ByRef re)
	Dim i
	Dim OutStr
	Dim iArticleID,sFileName,sFileExt

	re.Pattern = "(.*)\/(\w+)_(\d+).(\w+)"

	iArticleID = re.Replace(sFilePath,"$3")
	sFileName  = re.Replace(sFilePath,"$2_$3")
	sFileExt   = re.Replace(sFilePath,".$4")

	For i = 1 To iCount
		If i = iCurrentPage Then 
			OutStr = OutStr & "<span style='color: red;'>[" & i & "]</span>&nbsp;"
		ElseIf i = 1 Then
			OutStr = OutStr & "<a href='" & sFilePath & "'>[" & i & "]</a>&nbsp;"
		Else
			OutStr = OutStr & "<a href='" & sFileName & "_" & i & sFileExt & "'>[" & i & "]</a>&nbsp;"
		End If
	Next

	PageNav = OutStr
End Function
%>