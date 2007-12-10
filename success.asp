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
'= 文件名称：Success.asp
'= 摘    要：成功提示信息文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2006-09-12
'====================================================================

Dim SusNum,SusStr,NeedLogin
Dim PageContent

SusNum=Request.QueryString ("susnum")

Select Case SusNum
Case 1
	SusStr = SusStr & "您已经成功注册了一个新帐户，现正等待管理员的审核。"
Case 2
	SusStr = SusStr & "您已经成功发布您的评论，正在等待管理员的审核。"
Case 3
	SusStr = SusStr & "您已经成功修改了您的个人资料！"
Case 4
	SusStr = SusStr & EA_Pub.Mem_Info(1)&"，欢迎您回来！"
Case 5 
	SusStr = SusStr & "成功收藏文章"
Case 6
	SusStr = SusStr & "发表文章成功，但需要管理员的审核，请耐心等待"
Case 7
	SusStr = SusStr & "发表文章成功，你现在可以马上查看你发布的文章"
Case 8
	SusStr = SusStr & "您已经成功修改了您的登陆资料！"
Case 9
	SusStr = SusStr & "您已经成功注册了一个新帐户，现在可以使用这个帐户登录本系统了！"
End Select

PageContent=EA_Temp.Load_Template(0,"success")
	
EA_Temp.Title=EA_Pub.SysInfo(0)&" - 成功信息"
EA_Temp.Nav="<a href=""./""><b>"&EA_Pub.SysInfo(0)&"</b></a> - 成功信息"

PageContent=EA_Temp.Replace_PublicTag(PageContent)

PageContent=Replace(PageContent,"{$SusStr$}",SusStr)
	
Response.Write PageContent

Call EA_Pub.Close_Obj
Set EA_Pub=Nothing
%>