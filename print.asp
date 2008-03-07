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
'= 文件名称：Print.asp
'= 摘    要：文章打印版本文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2005-11-30
'====================================================================

Dim ArticleId,ArticleInfo
Dim FirstArticle,NextArticle
Dim i,TempStr,TempArray
Dim IsView

ArticleId=EA_Pub.SafeRequest(3,"articleid",0,0,0)

'load article info
ArticleInfo=EA_DBO.Get_Article_Info(ArticleId,0)
If Not IsArray(ArticleInfo) Then Call EA_Pub.ShowErrMsg(9,1)

If Not ArticleInfo(20,0) Or ArticleInfo(21,0) Then Call EA_Pub.ShowErrMsg(9,1)

If ArticleInfo(10,0) Then
	Response.Redirect ArticleInfo(11,0)
	Response.End 
End If

IsView=True

If ArticleInfo(22,0)>0 Or ArticleInfo(23,0)=-1 Then 
	If Not EA_Pub.IsMember Then 
		IsView=False
	Else
		If CLng(EA_Pub.Mem_GroupSetting(2))>=CLng(ArticleInfo(22,0)) Then 
			If ArticleInfo(23,0) Then 
				If EA_Pub.Mem_GroupSetting(3)="1" Then 
					IsView=True
				Else
					IsView=False
				End If
			Else
				IsView=True
			End If
		Else
			IsView=False
		End If
	End If
End If

If Not IsView Then 
	ArticleInfo(5,0)="<br><br><b>您当前的权限不允许查看该文章，请先 [<a href='member/login.asp' target='_blank'>登陆</a>] 或 [<a href='member/register.asp' target='_blank'>注册</a>]。</b>"
End If

%>
<html>
<head>
<title><%=ArticleInfo(3,0)%>-<%=ArticleInfo(2,0)%>|<%=EA_Pub.SysInfo(0)%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta http-equiv="Content-Language" content="zh-CN">
<meta name="Keywords" content="<%=Replace(EA_Pub.SysInfo(16),"|",",")&","&ArticleInfo(12,0)%>">
<meta name="Description" content="<%=ArticleInfo(4,0)%>">
</head>
<body> 
<table width="778" border="0" cellspacing="0" cellpadding="0"> 
<tr> 
  <td><!--PrintBegin--> 
    <table width="100%" border="0" cellspacing="0" cellpadding="0"> 
      <tr> 
        <td> <h3><%=EA_Pub.Add_ArticleColor(ArticleInfo(17,0),ArticleInfo(3,0))%></h3></td> 
      </tr> 
      <tr> 
        <td>作者:<%=ArticleInfo(8,0)%>　来源:<%=ArticleInfo(16,0)%>　最后修改于：<i><%=ArticleInfo(13,0)%></i>　<a href="javascript:vod();" onClick="window.print();">点击开始打印</a></td> 
      </tr> 
      <tr> 
        <td> <hr size="1" noshade color="#999999"> 
          页面地址是:<a href="<%=EA_Pub.Cov_ArticlePath(ArticleId,ArticleInfo(13,0),EA_Pub.SysInfo(18))%>">http://<%=Request.ServerVariables ("SERVER_NAME")%><%=EA_Pub.Cov_ArticlePath(ArticleId,ArticleInfo(13,0),EA_Pub.SysInfo(18))%></a> 
          <hr size="1" noshade color="#999999"> </td> 
      </tr> 
      <tr> 
        <td> <%=ArticleInfo(5,0)%> </td> 
      </tr> 
    </table> 
</body>
</html>
<%
Call EA_Pub.Close_Obj
Set EA_Pub=Nothing
%>
