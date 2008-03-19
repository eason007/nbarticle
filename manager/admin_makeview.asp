<!--#Include File="comm/inc.asp" -->
<!--#Include File="../include/page_article.asp"-->
<%
'====================================================================
'= Team Elite - EliteCMS
'= Copyright (c) 2004 - 2008 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：/Manager/Admin_MakeView.asp
'= 摘    要：后台-HTML栏目页生成文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-03-18
'====================================================================

Server.ScriptTimeout=9999999

Call EA_Manager.Chk_IsMaster

If Not EA_Manager.Chk_Power(Admin_Power,"43") Then 
	ErrMsg=str_Comm_NotAccess
	Call EA_Manager.Error(1)
End If

Dim Atcion
Dim ForTotal

Atcion=Request.Form("action")

Select Case LCase(Atcion)
Case "mark"
	Call MarkView
Case Else
	Call Main
End Select

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
	Dim TopicList,i,j
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
		StartBorder = EA_Pub.SafeRequest(2, "ColumnList", 1, "", 1)
		EndBorder = Split(StartBorder, ", ")
		If UBound(EndBorder) = -1 Then Exit Sub

		If UBound(EndBorder) >= 1 Then
			WSQL=" And ColumnId IN (" & StartBorder & ")"
		Else
			If StartBorder <> "0" Then WSQL=" And ColumnId="&StartBorder
		End If
	Case 4
		StartBorder=EA_Pub.SafeRequest(2,"sid",1,"",0)

		If Len(StartBorder)>0 Then WSQL=" And a.[id] In ("&StartBorder&")"
	End Select
		
	'0=ColumnId,1=ColumnCode,2=ColumnName,3=Title,4=Summary,5=Content,6=ViewNum,7=AuthorId,8=Author,9=CommentNum,10=IsOut
	'11=OutUrl,12=[KeyWord],13=AddDate,14=CutArticle,15=Source,16=SourceUrl,17=TColor,18=Img,19=IsTop,20=IsPass
	'21=IsDel,22=ListPower,23=IsHide,24=Article_TempId,25=TrueTime,26=SubTitle,27=SubUrl,28=Id
	SQL="Select ColumnId, ColumnCode, ColumnName, a.Title, Summary, Content, a.ViewNum, AuthorId, Author, CommentNum, a.IsOut, a.OutUrl, [KeyWord], AddDate, CutArticle, Source, SourceUrl, TColor, Img, a.IsTop, IsPass, IsDel, b.ListPower, b.IsHide, b.Article_TempId,TrueTime,a.SubTitle,a.SubUrl,a.Id"
	SQL=SQL&" FROM NB_Content AS a INNER JOIN NB_Column AS b ON a.ColumnId=b.Id Where IsPass="&EA_M_DBO.TrueValue&" And IsDel=0 "&WSQL&" Order By ColumnId Asc,a.[Id] Asc"
	'response.write sql
	'response.end
	Set Rs=Conn.Execute(SQL)
	If Not rs.eof And Not rs.bof Then 
		TopicList=Rs.GetRows()
		Rs.Close

		PageContent=EA_Temp.Load_Template_File("admin_makeview_view.htm")
		EA_Temp.P_Prefix = "{$"
		EA_Temp.P_Suffix = "$}"

		EA_Temp.SetVariable "PageTotal",Ubound(TopicList,2)+1,PageContent
		EA_Temp.SetVariable "Language_MakeList_Now",str_MakeList_Now,PageContent
		EA_Temp.SetVariable "Language_MakeList_Page",str_MakeList_Page,PageContent

		Response.Write PageContent

		PageContent = ""

		Dim IsReplace
		Dim PageContent, PageKeyword
		Dim TempStr,TempArray
		Dim Folder,sHTMLFilePath
		Dim NewFolderList
		Dim re
		Dim Tmp
		Dim clsArticle

		Set re = New RegExp
		Set clsArticle = New page_Article

		EA_Temp.P_Prefix = "<!--"
		EA_Temp.P_Suffix = "-->"

		re.IgnoreCase	= True
		re.Global		= True
		re.Pattern		= Replace(SystemFolder, "/", "\/") & "(.*)\/(\w+)_(\d+).(\w+)"

		NewFolderList = ","
		PageKeyword   = EA_Pub.SysInfo(16)
		ForTotal	  = UBound(TopicList, 2)
		EA_Pub.SysInfo(18) = "0"

		For i = 0 To ForTotal
			PageContent  = ""
			IsReplace	 = True
			EA_Pub.SysInfo(16) = PageKeyword
			sHTMLFilePath= EA_Pub.Cov_ArticlePath(TopicList(28, i), TopicList(13, i), "0")
			
			'check folder isexists
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
			If TopicList(22,i) > 0 Or TopicList(23,i) <> 0 Then
				IsReplace	= False
				TempStr		= "<meta http-equiv=""refresh"" content=""0;URL="&SystemFolder&"article.asp?articleid="&TopicList(28,i)&""">"
			End If
			
			'check is out
			If TopicList(10,i) And IsReplace Then
				IsReplace	= False
				TempStr		= "<meta http-equiv=""refresh"" content=""0;URL="&TopicList(11,i)&""">"
			End If
			
			If Not IsReplace Then
			'not replace template tag
				Call EA_Pub.Save_HtmlFile(sHTMLFilePath,TempStr)
			Else
				Call clsArticle.CutContent("\[NextPage([^\]])*\]", TopicList(5, i))

				If UBound(clsArticle.PageIndex) = 1 Then
					PageContent = clsArticle.Make(TopicList(28, i), GetOneArray(TopicList, i), 1, True)

					Call EA_Pub.Save_HtmlFile(sHTMLFilePath, PageContent)
				Else
					For j = 1 To UBound(clsArticle.PageIndex)
						PageContent = clsArticle.Make(TopicList(28, i), GetOneArray(TopicList, i), j, True)
						
						If j = 1 Then
							Call EA_Pub.Save_HtmlFile(sHTMLFilePath, PageContent)
						Else
							Call EA_Pub.Save_HtmlFile(re.Replace(sHTMLFilePath, "/$1/$2_$3_" & j & ".$4"), PageContent)
						End If
					Next
				End If
			End If
			
			If ForTotal = 0 Then 
				Response.Write "<script>img1.width=400;" & VbCrLf
			Else
				Response.Write "<script>img1.width=" & Fix(((i + 1) / ForTotal) * 100) & ";" & VbCrLf
			End If

			Response.Write "column_complete.innerHTML=""<font color=green>" & i + 1 & "</font>"";</script>" & VbCrLf
			Response.Flush
		Next
		
		Response.Write "<script>img1.width=400;"& VbCrLf
		Response.Write "make_msg.innerHTML="""&str_MakeList_AllComplate&""";</script>" & VbCrLf
	End If
End Sub

Function GetOneArray(ExArray, RowNum)
	Dim TmpArray()
	Dim i, k

	k = UBound(ExArray)
	ReDim TmpArray(k + 1, 1)

	For i = 0 To k
		TmpArray(i, 0) = ExArray(i, RowNum)
	Next

	GetOneArray = TmpArray
End Function
%>