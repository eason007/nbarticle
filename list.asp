<!--#Include File="conn.asp" -->
<!--#Include File="include/inc.asp"-->
<!--#include file="include/_cls_teamplate.asp"-->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：List.asp
'= 摘    要：列表页文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2006-11-12
'====================================================================

Dim PageContent

'check system state
If EA_Pub.SysInfo(18)="0" Then
	PageContent = EA_Pub.Cov_ColumnPath(Request("classid"),EA_Pub.SysInfo(18))

	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing

	Response.Redirect PageContent
	Response.End 
End If

Dim ColumnId,ColumnInfo
'load column data
ColumnId	= EA_Pub.SafeRequest(3,"classid",0,0,0)
ColumnInfo	= EA_DBO.Get_Column_Info(ColumnId)
If Not IsArray(ColumnInfo) Then Call EA_Pub.ShowErrMsg(9,1)

'jump to out url
If ColumnInfo(6,0) Then 
	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing

	Response.Redirect ColumnInfo(7,0)
	Response.End 
End If

Dim ArticleList,ArticleOutStr
Dim ChildColumnList
Dim FieldName(0),FieldValue(0)
Dim i
Dim PageNum,PageCount,PageSize
Dim Template
Dim Temp,ListBlock
Dim ForTotal

FieldName(0)	= "classid"
FieldValue(0)	= ColumnId

PageNum	= EA_Pub.SafeRequest(3,"page",0,1,0)
PageSize= ColumnInfo(17,0)

Set Template=New cls_NEW_TEMPLATE

PageCount=EA_Pub.Stat_Page_Total(PageSize,ColumnInfo(3,0))
If CLng(PageNum)>PageCount And PageCount>0 Then PageNum=PageCount

'load article list
If ColumnInfo(3,0)>0 Then ArticleList=EA_DBO.Get_Article_ByColumnId(ColumnId,PageNum,PageSize)

PageContent=EA_Temp.Load_Template(ColumnInfo(9,0),"list")

'make article list
ListBlock=Template.GetBlock("list",PageContent)

'--------输出---------
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

		Template.SetBlock "list",Temp,PageContent
	Next
	Template.CloseBlock "list",PageContent
End If

'replate template tag
EA_Temp.Title=ColumnInfo(0,0)&" - "&EA_Pub.SysInfo(0)
EA_Temp.Nav="<a href=""./""><b>"&EA_Pub.SysInfo(0)&"</b></a>"&EA_Pub.Get_NavByColumnCode(ColumnInfo(1,0))

PageContent=Replace(PageContent,"{$ColumnTitleNav$}","")

PageContent=Replace(PageContent,"{$ColumnId$}",ColumnId)
PageContent=Replace(PageContent,"{$ColumnName$}",ColumnInfo(0,0))
PageContent=Replace(PageContent,"{$ColumnInfo$}",ColumnInfo(2,0))
PageContent=Replace(PageContent,"{$ColumnTopicTotal$}",ColumnInfo(3,0))
PageContent=Replace(PageContent,"{$ColumnMangerTotal$}",ColumnInfo(4,0))

PageContent=Replace(PageContent,"{$ColumnTopicList$}",ArticleOutStr)
PageContent=Replace(PageContent,"{$ColumnPageNumNav$}",EA_Temp.PageList(PageCount,PageNum,FieldName,FieldValue))

PageContent=Replace(PageContent,"{$ColumnNav$}",SiteColumnNav(ColumnId,ColumnInfo(1,0)))
PageContent=Replace(PageContent,"{$ChildColumn$}",ColumnChild(ColumnInfo(1,0)))

EA_Temp.Find_TemplateTagByInput "ChildColumnNav",ChildColumnNav,PageContent

EA_Pub.SysInfo(16)=ColumnInfo(0,0)&","&EA_Pub.SysInfo(16)
If Len(ColumnInfo(2,0)) Then EA_Pub.SysInfo(17)=ColumnInfo(2,0)

PageContent=EA_Temp.Replace_PublicTag(PageContent)

Response.Write PageContent

Call EA_Pub.Close_Obj
Set EA_Pub=Nothing

Function TagList (Keyword)
	Dim TempArray,i
	Dim ForTotal
	Dim OutStr

	If Len(Keyword) > 0 Then
		TempArray= Split(Keyword,",")

		ForTotal = UBound(TempArray)

		For i=0 To ForTotal
			If Len(TempArray(i)) > 0 Then OutStr = OutStr & "<a href='" & SystemFolder & "search.asp?action=query&field=1&keyword=" & server.urlencode(Trim(TempArray(i))) & "' rel='external'>" & Trim(TempArray(i)) & "</a>&nbsp;"
		Next
	End If

	TagList = OutStr
End Function

Function ChildColumnNav()
	Dim ChilColumnConfig
	Dim Temp,OutStr,Column,j

	Temp = Split(ChildColumnList,"|")

	ChilColumnConfig = EA_Temp.Find_TemplateTagValues("ChildColumnNav",PageContent)
	If Not IsArray(ChilColumnConfig) Then Exit Function

	j = 1
	ForTotal = UBound(Temp)-1

	For i=0 To ForTotal
		Column = Split(Temp(i),",")

		OutStr = OutStr & "<a href="""&EA_Pub.Cov_ColumnPath(Column(0),EA_Pub.SysInfo(18))&""">"&Column(1)
		OutStr = OutStr & "</a>&nbsp;"

		If j = CLng(ChilColumnConfig(1)) Then Exit For
		j = j + 1
		If (i+1) Mod ChilColumnConfig(0) = 0 And (i+1) <= (UBound(Temp)-1) Then OutStr = OutStr & "<br>"
	Next

	ChildColumnNav = OutStr
End Function

Function ColumnChild(MainCode)
	Dim ChildList,Title,TopicList
	Dim j,k
	Dim OurStr
	
	ChildList=EA_DBO.Get_Column_ChildList(MainCode)
	If IsArray(ChildList) Then 
		ForTotal = Ubound(ChildList,2)
		For i=0 To ForTotal
			OurStr=OurStr&"<table>"&Chr(10)
			OurStr=OurStr&"<tr>"&Chr(10)
			OurStr=OurStr&"<td>&nbsp;<strong>"&ChildList(1,i)&"</strong>&nbsp;&nbsp;&nbsp;文章总数:"&ChildList(2,i)&"&nbsp;浏览次数:"&ChildList(3,i)&"</td>"&Chr(10)
			OurStr=OurStr&"<td><a href="""&EA_Pub.Cov_ColumnPath(ChildList(0,i),EA_Pub.SysInfo(18))&"""><img src=""Images/more.gif"" alt="""" /></a>&nbsp;</td>"&Chr(10)
			OurStr=OurStr&"</tr>"&Chr(10)
			
			TopicList=EA_DBO.Get_Article_ByColumnId(ChildList(0,i),1,10)
			If IsArray(TopicList) Then 
				OurStr=OurStr&"<tr>"&Chr(10)
				OurStr=OurStr&"<td colspan=""2"">"&Chr(10)
				OurStr=OurStr&"<table>"&Chr(10)
				k = Ubound(TopicList,2)
				For j=0 To k
					Title=Replace(TopicList(2,j),"&nbsp;","")
					OurStr=OurStr&"<tr>"&Chr(10)
					OurStr=OurStr&"<td>&nbsp;<a href="""&EA_Pub.Cov_ArticlePath(TopicList(0,j),TopicList(3,j),EA_Pub.SysInfo(18))&""" rel=""external"">"
					OurStr=OurStr&EA_Pub.Add_ArticleColor(TopicList(1,j),Title)
					OurStr=OurStr&"</a>"
					OurStr=OurStr&EA_Pub.Chk_ArticleTime(TopicList(3,j))
					OurStr=OurStr&"</td>"
					OurStr=OurStr&"<td style='color: #AAAAAA;'>&nbsp;"&FormatDateTime(TopicList(3,j),2)&"&nbsp;Browse:"&TopicList(7,j)&"</td>"&Chr(10)
					OurStr=OurStr&"</tr>"&Chr(10)
				Next
				OurStr=OurStr&"</table>"&Chr(10)
				OurStr=OurStr&"</td>"&Chr(10)
				OurStr=OurStr&"</tr>"&Chr(10)
			Else
				OurStr=OurStr&"<tr><td style=""height: 50px;"">此栏目暂时没有文章</td></tr>"&Chr(10)
			End If
			OurStr=OurStr&"</table>"&Chr(10)
		Next
	End If
	
	ColumnChild=OurStr
End Function

Function SiteColumnNav(ColumnId,ColumnCode)
	Dim TempArray,TempStr
	Dim i,StepLen
	
	TempArray=EA_DBO.Get_Column_Nav(ColumnCode)
	
	TempStr="<table>"
	If IsArray(TempArray) Then 
		ForTotal = UBound(TempArray,2)
		For i=0 To ForTotal
			StepLen=(Len(TempArray(1,i))/4)*2-2
			If Len(TempArray(1,i)) = Len(ColumnCode)+4 And ColumnCode = Left(TempArray(1,i),Len(ColumnCode)) Then ChildColumnList = ChildColumnList & TempArray(0,i) & "," & TempArray(2,i) & "|"

			TempStr=TempStr&"<tr><td>&nbsp;"
			If Len(TempArray(1,i))>4 Then 
				TempStr=TempStr&"├"
				TempStr=TempStr&String(StepLen,"-")
			End If
			If CLng(TempArray(0,i))=CLng(ColumnId) Then 
				TempStr=TempStr&"<img src=""images/public/icon2.gif"" alt="""" />"
			Else
				TempStr=TempStr&"<img src=""images/public/icon.gif"" alt="""" />"
			End If
			TempStr=TempStr&"<a href="""&EA_Pub.Cov_ColumnPath(TempArray(0,i),EA_Pub.SysInfo(18))&""">"&TempArray(2,i)
			TempStr=TempStr&"</a>"
			If TempArray(4,i) Then TempStr=TempStr&"[专]"
			TempStr=TempStr&"&nbsp;<span style=""color: #aaaaaa;"">("&TempArray(5,i)&")</span>"
			TempStr=TempStr&"</td>"
			TempStr=TempStr&"</tr>"
		Next
	End If
	TempStr=TempStr&"</table>"
	
	SiteColumnNav=TempStr
End Function
%>