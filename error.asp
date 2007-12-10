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
'= 文件名称：Error.asp
'= 摘    要：错误提示信息文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2005-04-20
'====================================================================

Dim ErrNum,ErrStr,NeedLogin
Dim PageContent

ErrNum=Request.QueryString ("errnum")

Select Case ErrNum
Case 1
	ErrStr = ErrStr & "请填写必要内容！如用户名/密码/评论内容"
Case 2
	ErrStr = ErrStr & "密码不能为空，且长度必须小于19位大于5位！"
Case 3
	ErrStr = ErrStr & "您输入的用户名不能登陆。有可能是您的帐号还没被管理员审核通过。"
	NeedLogin = True
Case 4
	ErrStr = ErrStr & "您两次输入的密码不相同！"
Case 5
	ErrStr = ErrStr & "可能是您申请注册的这个昵称已经被另一个用户使用了，或者您填写的电子信箱地址已经在本系统注册了，请返回选择其他昵称/电子信箱地址！"
Case 6
	ErrStr = ErrStr & "请不要使用非法字符，(只能使用大小写英文字符、阿拉伯数字、简/繁体中文字符以及下划线_)！"
Case 7
	ErrStr = ErrStr & "用户名必须小于14个字符(7个汉字)！"
Case 8
	ErrStr = ErrStr & "请填写正确、有效的EMail地址！"
Case 9
	ErrStr = ErrStr & "真是不可思议！您好象进错门了？！"
Case 10
	NeedLogin = True
	ErrStr = ErrStr & "必须登系统后才能进行操作！请登录！"
Case 11
	ErrStr = ErrStr & "您输入的日期格式不正确！日期格式必须为xx-xx-xx的形式！"
Case 12
	ErrStr = ErrStr & "您没有足够的权限浏览此栏目！如有任何问题，请查看系统说明，或与管理员联系！"
Case 13
	ErrStr = ErrStr & "不能对您自己进行此操作！"
Case 14
	ErrStr = ErrStr & "对不起，你没有足够的权限发表文章！"
Case 15
	ErrStr = ErrStr & "对不起，在更新过程中发生错误，没有任何文章被发布，请联系管理员！"
Case 16
	ErrStr = ErrStr & "您发表的评论内客少于5个字符，请修后再发！"
Case 17
	ErrStr = ErrStr & "您发表的评论内客超过了1000个字符，请修改后再发！"
Case 18
	ErrStr = ErrStr & "用户不存在，如有疑问，请联系社区管理员！"
Case 19
	ErrStr = ErrStr & "用户已经存在于列表中！"
Case 20
	ErrStr = ErrStr & "请不要手动修改地址参数,同时请填写好所有必须填写的项目！"
Case 21
	ErrStr = ErrStr & "发表文章成功，但需要管理员的审核，请耐心等待"
Case 22
	ErrStr = ErrStr & "发表文章成功，你现在可以马上查看你发布的文章"
Case 23
	ErrStr = ErrStr & "请求失败，很可能是您不具备执行此操作的权限"
Case 24
	ErrStr = ErrStr & "该文章你已经收藏过。"
Case 25
	ErrStr = ErrStr & "成功收藏文章！"
Case 26
	ErrStr = ErrStr & "传递错误的查询数据！"
Case 27
	ErrStr = ErrStr & "标题内容长度不符。(大于150或者等于0个字符)"
Case 28
	ErrStr = ErrStr & "作者长度不符。(大于16个字符)"
Case 29
	ErrStr = ErrStr & "关键字长度不符。(大于50个字符)"
Case 30
	ErrStr = ErrStr & "对不起，本站编辑的文章不提供此项操作！"
Case 31
	ErrStr = ErrStr & "此文件可能已经被管理人员删除！"
Case 33
	ErrStr = ErrStr & "系统管理员已经关闭了注册功能！"
Case 34
	ErrStr = ErrStr & "请输入搜索关键字！"
Case 35
	ErrStr = ErrStr & "请选择搜索的栏目！"
Case 36
	ErrStr = ErrStr & "您输入的用户名密码不能通过验证，请不要尝试登陆别人的帐号！"
Case 37
	ErrStr = ErrStr & "您所在用户组已被管理员设置为不允许登陆，请联系管理员！"
Case 38
	ErrStr = ErrStr & "您所在用户组已被管理员设置为当前时段不允许登陆，请联系管理员！"
Case 39
	ErrStr = ErrStr & "您选择发布文章的栏目已被管理员设置为不允许投稿！"
Case 40
	ErrStr = ErrStr & "您的收藏夹中的数目已到达上限，请删除其他收藏文章后再进行添加收藏操作！"
Case 41
	ErrStr = ErrStr & "您的今日的投稿数已到达上限，请稍后再进行此操作！"
End Select

PageContent=EA_Temp.Load_Template(0,"error")
	
EA_Temp.Title=EA_Pub.SysInfo(0)&" - 错误信息"
EA_Temp.Nav="<a href=""./""><b>"&EA_Pub.SysInfo(0)&"</b></a> - 错误信息"

PageContent=EA_Temp.Replace_PublicTag(PageContent)

PageContent=Replace(PageContent,"{$ErrStr$}",ErrStr)

Response.Write PageContent

Call EA_Pub.Close_Obj
Set EA_Pub=Nothing
%>