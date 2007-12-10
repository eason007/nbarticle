<!--#Include File="conn.asp" -->
<!--#Include File="include/inc.asp"-->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：RSSFeed.asp
'= 摘    要：RSS订阅文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2006-10-21
'====================================================================

'Types:0=最新，1=热点，2=推荐，3=图片，4=作者专栏
'ColumnId:0=全站，非0=某单独栏目，指定一级栏目将自动包含下属栏目

Dim Types,ColumnId,AuthorName,TopNum
Dim OutStr,TopicList,i,TUrl
Dim ForTotal

Types=EA_Pub.SafeRequest(1,"type",0,0,0)
ColumnId=EA_Pub.SafeRequest(1,"cid",0,0,0)
AuthorName=EA_Pub.SafeRequest(1,"author",1,"",0)
TopNum=EA_Pub.SafeRequest(1,"top",0,20,0)

If Right(EA_Pub.SysInfo(11),1)<>"/" Then EA_Pub.SysInfo(11)=SysInfo(11)&"/"
If Left(EA_Pub.SysInfo(11),7)<>"http://" Then EA_Pub.SysInfo(11)="http://"& SysInfo(11)

Response.ContentType = "text/XML"
OutStr="<?xml version=""1.0"" encoding=""utf-8"" ?>"&chr(10)
OutStr=OutStr & "<rss version=""2.0"">"&chr(10)
OutStr=OutStr&"<channel>"&chr(10)
OutStr=OutStr&"	<title>"&EA_Pub.SysInfo(0)&"</title>"&chr(10)
OutStr=OutStr&"	<description>"&EA_Pub.SysInfo(17)&"</description>"&chr(10)
OutStr=OutStr&"	<link>"&EA_Pub.SysInfo(11)&"</link>"&chr(10)
OutStr=OutStr&"	<language>zh-cn</language>"&chr(10)
OutStr=OutStr&"	<generator>NBArticle</generator>"&chr(10)
OutStr=OutStr&"	<managingEditor>"&EA_Pub.SysInfo(12)&"</managingEditor>"&chr(10)

TopicList=EA_DBO.Get_Rss_List(TopNum,Types,AuthorName,ColumnId)

If IsArray(TopicList) Then
	ForTotal = UBound(TopicList,2)

	For i=0 To ForTotal
		TUrl=EA_Pub.Cov_ArticlePath(TopicList(0,i),TopicList(2,i),EA_Pub.SysInfo(18))
		If SystemFolder<>"/" Then TUrl=Replace(TUrl,SystemFolder,"")
		TUrl=EA_Pub.SysInfo(11)&TUrl
		
		OutStr=OutStr&"	<item>"&chr(10)
		OutStr=OutStr&"	<title><![CDATA["&TopicList(1,i)&"]]></title>"&chr(10)
		OutStr=OutStr&"	<link>"&TUrl&"</link>"&chr(10)
		OutStr=OutStr&"	<description><![CDATA["&TopicList(3,i)&"…… [<a href="""&TUrl&""">点击查看详细</a>] ]]></description>"&chr(10)
		OutStr=OutStr&"	<pubDate>"&RSS_FormatDateTime(TopicList(2,i))&"</pubDate>" &chr(10)
		OutStr=OutStr&"	<comments>" & EA_Pub.SysInfo(11) & "review.asp?articleid=" & TopicList(0,i) & "</comments>" &chr(10)
		OutStr=OutStr&"	</item>"&chr(10)&chr(10)
	Next
End If

OutStr=OutStr&"</channel>"&chr(10)
OutStr=OutStr&"</rss>"&chr(10)
Response.Write OutStr

Call EA_Pub.Close_Obj
Set EA_Pub=Nothing

Function RSS_FormatDateTime(sDate)
	Dim RStr
	Select Case WeekDay(sDate)
	Case 1
		RStr="Sun, "
	Case 2
		RStr="Mon, "
	Case 3
		RStr="Tue, "
	Case 4
		RStr="Wed, "
	Case 5
		RStr="Thu, "
	Case 6
		RStr="Fri, "
	Case 7
		RStr="Sat, "
	End Select
	
	RStr=RStr&Day(sDate)&" "
	
	Select Case Month(sDate)
	Case 1
		RStr=RStr&"Jan "
	Case 2
		RStr=RStr&"Feb "
	Case 3
		RStr=RStr&"Mar "
	Case 4
		RStr=RStr&"Apr "
	Case 5
		RStr=RStr&"May "
	Case 6
		RStr=RStr&"Jun "
	Case 7
		RStr=RStr&"Jul "
	Case 8
		RStr=RStr&"Aug "
	Case 9
		RStr=RStr&"Sep "
	Case 10
		RStr=RStr&"Oct "
	Case 11
		RStr=RStr&"Nov "
	Case 12
		RStr=RStr&"Dec "
	End Select
	
	RStr=RStr&Year(sDate)&" "
	RStr=RStr&Right("00"&Hour(sDate)-8,2)&":"
	RStr=RStr&Right("00"&Minute(sDate),2)&":"
	RStr=RStr&Right("00"&Second(sDate),2)&" GMT"
	
	RSS_FormatDateTime=RStr
End Function
%>