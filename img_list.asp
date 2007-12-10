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
'= 文件名称：Img_List.asp
'= 摘    要：图片文章列表文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2006-07-27
'====================================================================

Dim QueryArray,QueryList
Dim PageNum,PageSize,PageCount,ReCount,Column
Dim i
Dim FieldName(0),FieldValue(0)
Dim RowSize,RowWidth
Dim PageContent
Dim Template
Dim Temp,ListBlock
Dim ForTotal

Set Template=New cls_NEW_TEMPLATE

PageNum=EA_Pub.SafeRequest(1,"page",0,1,0)

PageSize=12
RowSize=4
RowWidth=100/RowSize

PageContent=EA_Temp.Load_Template(0,"imglist")

ReCount=EA_DBO.Get_Article_ImgStat()(0,0)
PageCount=EA_Pub.Stat_Page_Total(PageSize,ReCount)
If PageNum>PageCount And PageCount>0 Then PageNum=PageCount

ListBlock=Template.GetBlock("list",PageContent)

If PageCount>0 Then 
	QueryArray=EA_DBO.Get_Article_ImgList(PageNum,PageSize)
	ForTotal = UBound(QueryArray,2)
	
	For i=0 To ForTotal
		Temp=ListBlock
  
		Template.SetVariable "Url",EA_Pub.Cov_ArticlePath(QueryArray(0,i),QueryArray(5,i),EA_Pub.SysInfo(18)),Temp
		Template.SetVariable "Img",QueryArray(2,i),Temp
		Template.SetVariable "Title",EA_Pub.Add_ArticleColor(QueryArray(3,i),QueryArray(1,i)),Temp

		Template.SetBlock "list",Temp,PageContent
	Next 
End If

Template.CloseBlock "list",PageContent
	
EA_Temp.Title=EA_Pub.SysInfo(0)&" - 图片文章"
EA_Temp.Nav="<a href=""./""><b>"&EA_Pub.SysInfo(0)&"</b></a> - 图片文章"

PageContent=EA_Temp.Replace_PublicTag(PageContent)
	
PageContent=Replace(PageContent,"{$Query_List$}",QueryList)
PageContent=Replace(PageContent,"{$PageNumNav$}",EA_Temp.PageList(PageCount,PageNum,FieldName,FieldValue))
	
Response.Write PageContent

Call EA_Pub.Close_Obj
Set EA_Pub=Nothing
%>
