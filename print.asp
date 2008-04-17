<!--#Include File="include/inc.asp"-->
<%
'====================================================================
'= Team Elite - EliteCMS
'= Copyright (c) 2004 - 2008 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：Print.asp
'= 摘    要：文章打印版本文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-04-17
'====================================================================

Dim ArticleId, ArticleInfo
Dim FirstArticle,NextArticle
Dim i,TempStr,TempArray
Dim IsView
Dim PageContent

ArticleId=EA_Pub.SafeRequest(3,"articleid",0,0,0)

'load article info
ArticleInfo = EA_DBO.Get_Article_Info(ArticleId, 1)
If Not IsArray(ArticleInfo) Then Call EA_Pub.ShowErrMsg(2, 0)
If Not ArticleInfo(20, 0) Or ArticleInfo(21, 0) Then Call EA_Pub.ShowErrMsg(2, 0)

If ArticleInfo(22, 0) > 0 Or ArticleInfo(23, 0) <> 0 Then 
	If Not EA_Pub.IsMember Then 
		IsView = False
	Else
		If CDbl(EA_Pub.Mem_GroupSetting(2)) >= CDbl(ArticleInfo(22, 0)) Then 
			If ArticleInfo(23, 0) Then 
				If EA_Pub.Mem_GroupSetting(3) = "1" Then 
					IsView = True
				Else
					IsView = False
				End If
			Else
				IsView = True
			End If
		Else
			IsView = False
		End If
	End If
Else
	IsView = True
End If

If Not IsView Then 
	ArticleInfo(5,0)="<br><br><b>您当前的权限不允许查看该文章，请先 [<a href='member/login.asp' target='_blank'>登陆</a>] 或 [<a href='member/register.asp' target='_blank'>注册</a>]。</b>"
End If
%>
<html>
<head>
<title><%=ArticleInfo(3, 0) & " - " & ArticleInfo(2, 0) & " - " & EA_Pub.SysInfo(0)%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta http-equiv="Content-Language" content="zh-CN">
<style type="text/css">
body {margin: 5px 20px;font-size: 12px;font-family: Helvetica, Arial, sans-serif;}
</style>
</head>
<body>

<div id="print">
  <h1><%=EA_Pub.Add_ArticleColor(ArticleInfo(17, 0),ArticleInfo(3, 0))%></h1>
  <div><strong>网址：</strong><a href="<%=Left(EA_Pub.SysInfo(11), Len(EA_Pub.SysInfo(11)) - 1) & EA_Pub.Cov_ArticlePath(ArticleId, ArticleInfo(13, 0), EA_Pub.SysInfo(18))%>"><%=Left(EA_Pub.SysInfo(11), Len(EA_Pub.SysInfo(11)) - 1) & EA_Pub.Cov_ArticlePath(ArticleId, ArticleInfo(13, 0), EA_Pub.SysInfo(18))%></a></div>
  <div><strong>作者：</strong><%=ArticleInfo(8, 0)%>&nbsp;&nbsp;&nbsp;<strong>来源：</strong><%=ArticleInfo(15,0)%>&nbsp;&nbsp;&nbsp;<strong>日期：</strong><%=ArticleInfo(13, 0)%>&nbsp;&nbsp;&nbsp;<a href="javascript:vod();" onClick="window.print();"><%=SysMsg(53)%></a></div>
  <div><%=ArticleInfo(5,0)%></div>
</div>

</body>
</html>