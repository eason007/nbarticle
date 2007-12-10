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
'= 文件名称：Member_List.asp
'= 摘    要：会员列表文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2006-07-27
'====================================================================

Dim Action
Dim PageNum,PageSize,PageCount,ReCount
Dim FieldName(1),FieldValue(1)
Dim QueryArray
Dim PageContent,QueryList
Dim i
Dim Temp,ListBlock
Dim Template
Dim ForTotal

Set Template=New cls_NEW_TEMPLATE

Action=Request.QueryString ("action")
PageNum=EA_Pub.SafeRequest(1,"page",0,1,0)
PageSize=20

FieldName(0)="action"
FieldValue(0)=Action

PageContent=EA_Temp.Load_Template(0,"memberlist")

ReCount=EA_DBO.Get_Member_Total()(0,0)
PageCount=EA_Pub.Stat_Page_Total(PageSize,ReCount)
If PageNum>PageCount And PageCount>0 Then PageNum=PageCount

QueryArray=EA_DBO.Get_Member_List(Action,"",PageNum,PageSize)

ListBlock=Template.GetBlock("list",PageContent)

If IsArray(QueryArray) Then 
	ForTotal = UBound(QueryArray,2)

	For i = 0 To ForTotal
		Temp=ListBlock
  
		Template.SetVariable "Name",QueryArray(1,i),Temp
		Template.SetVariable "Sex",QueryArray(2,i),Temp
		Template.SetVariable "Email",QueryArray(3,i),Temp
		Template.SetVariable "QQ",QueryArray(4,i),Temp
		Template.SetVariable "GroupName",QueryArray(5,i),Temp
		Template.SetVariable "RegTime",FormatDateTime(QueryArray(6,i),2),Temp
		Template.SetVariable "HomePage",QueryArray(7,i),Temp
		Template.SetVariable "PostTotal",QueryArray(8,i),Temp

		Template.SetBlock "list",Temp,PageContent
	Next
End If
Template.CloseBlock "list",PageContent

EA_Temp.Title=EA_Pub.SysInfo(0)&" - 会员列表"
EA_Temp.Nav="<a href=""./""><b>"&EA_Pub.SysInfo(0)&"</b></a> - 会员列表"

PageContent=Replace(PageContent,"{$Query_List$}",QueryList)
PageContent=Replace(PageContent,"{$PageNumNav$}",EA_Temp.PageList(PageCount,PageNum,FieldName,FieldValue))

PageContent=EA_Temp.Replace_PublicTag(PageContent)

Response.Write PageContent

Call EA_Pub.Close_Obj
Set EA_Pub=Nothing
%>