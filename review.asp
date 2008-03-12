<!--#Include File="init.asp" -->
<!--#Include File="include/inc.asp"-->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：Review.asp
'= 摘    要：评论文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2007-01-20
'====================================================================

Response.Buffer = True 
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache"

Dim Action
Action=Request.QueryString ("action")

Select Case LCase(Action)
Case "review_min"
	Call Min_ReviewList
Case "save"
	Call Save_Review
Case Else
	Call Main
End Select
Call EA_Pub.Close_Obj
Set EA_Pub=Nothing

Sub Main
	Dim ArticleId
	ArticleId=EA_Pub.SafeRequest(3,"articleid",0,0,3)
	
	EA_Temp.Title=EA_Pub.SysInfo(0)&" - 评论系统"
	EA_Temp.Nav="<a href=""./""><b>"&EA_Pub.SysInfo(0)&"</b></a> - <a href=""article.asp?articleid="&ArticleId&""" target=""_blank"">文章正文</a> - 评论列表"
%>
<html>
<head>
<link title="RSS" type="application/rss+xml" rel="alternate" href='<%=EA_Pub.SysInfo(11)%>rssfeed.asp'/>
<meta http-equiv="Content-Type" content="text/html;charset=utf-8">
<meta http-equiv="Content-Language" content="zh-CN">
<meta name="keywords" content="<%=Replace(EA_Pub.SysInfo(16),"|",",")%>">
<meta name="Description" content="<%=EA_Pub.SysInfo(17)%>">
<meta name="generator" content="NB文章系统(NBArticle)">
<title><%=EA_Temp.Title%></title>
<link href="member/style.css" rel="stylesheet" type="text/css" />
<script language="JavaScript" src="js/public.js"></script> 
<table width="750" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="efefef"> 
<form name="form1" method="post" action="?action=save&articleid=<%=ArticleId%>">  
  <tr> 
    <td bgcolor="ffffff" height="25">&nbsp;<%=EA_Temp.Nav%></td>
  </tr>
  <tr> 
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0"> 
          <tr> 
            <td align="center" width="60%"><a name="review"></a> 
              <table width="100%" border="0" cellpadding="0" cellspacing="0"> 
                <tr> 
                  <td width="20%"><img src="images/public/topic_author.gif" align="absmiddle"> 笔名：</td> 
                  <td><input type="text" name="name" class="LoginInput" size="42"></td> 
                </tr> 
                <tr> 
                  <td><img src="images/public/topic_ping.gif" align="absmiddle"> 评论：</td> 
                  <td><textarea name="Review" cols="40" rows="10" id="Review" onkeydown="ctlent()"></textarea></td> 
                </tr> 
                <tr> 
                  <td colspan="2" align="center"><input type="submit" name="Submit" value="发表评论" class="LoginInput"> 
                    <input type="reset" name="Submit2" value="重写评论" class="LoginInput"></td> 
                </tr> 
                <tr> 
                  <td colspan="2" align="center">[评论将在5分钟内被审核，请耐心等待]</td> 
                </tr> 
              </table></td> 
            <td valign="top"> 【注】 发表评论必需遵守以下条例：
              <ul style="list-style-type:square;margin-left:1em;line-height:150%"> 
                <li>尊重网上道德，遵守中华人民共和国的各项有关法律法规
                <li>承担一切因您的行为而直接或间接导致的民事或刑事法律责任
                <li>本站管理人员有权保留或删除其管辖留言中的任意内容
                <li>本站有权在网站内转载或引用您的评论
                <li>参与本评论即表明您已经阅读并接受上述条款
              </ul></td> 
          </tr> 
        </table></td> 
  </tr> 
</form> 
</table>
<br>
<table width="750" border="0" align="center" cellpadding="0" cellspacing="0" style="word-break: break-all">
	<tr>
		<td width="100%" bgcolor="ffffff"><SPAN ID=CommentList>正在读取...</SPAN></td>
	</tr>
</table>
<script language=javascript>
<%
	Dim TopicList

	TopicList=EA_DBO.Get_Review_List(ArticleId)

	If IsArray(TopicList) Then 
		OutPutJsList TopicList,15,1
		Response.Write vbcrlf&"ShowContentList(1,0);"
	Else
		Response.Write "document.getElementById('CommentList').innerHTML=""·暂无评论"";"
	End If
%>
</script>
<%
End Sub 

Sub Min_ReviewList
	Dim ArticleId
	Dim TopicList
	ArticleId=EA_Pub.SafeRequest(3,"articleid",0,0,3)
	
	EA_DBO.Set_Article_CommentNum_UpDate(ArticleId)
	TopicList=EA_DBO.Get_Review_List(ArticleId)

	If IsArray(TopicList) Then 
		OutPutJsList TopicList,5,0
		Response.Write vbcrlf&"ShowContentList(1,1);"
	Else
		Response.Write "document.getElementById('CommentList').innerHTML='·暂无评论';"
	End If
End Sub

Sub OutPutJsList(TopicList,PageSize,ShowStyle)
	Dim i,RCount
	Dim ForTotal

	If IsArray(TopicList) Then 
		RCount=UBound(TopicList,2)+1

		Response.Write "var RCount="&RCount&";"&vbcrlf
		Response.Write "var PUB_PageSize="&PageSize&";"&vbcrlf
		If (RCount Mod PageSize)=0 Then
			Response.Write "var PUB_PageCount="&(RCount\PageSize)&";"&vbcrlf
		Else
			Response.Write "var PUB_PageCount="&(RCount\PageSize)+1&";"&vbcrlf
		End If
		Response.Write "var ReviewList=[];"&vbcrlf
		If IsArray(TopicList) Then 
			ForTotal = UBound(TopicList,2)

			For i=0 To ForTotal
				If ShowStyle=0 Then 
					Response.Write "ReviewList["&i&"]="""&TopicList(0,i)&"||"&TopicList(1,i)&"||"&EA_Pub.Un_Base_HTMLFilter(TopicList(2,i))&""";"&vbcrlf
				Else
					Response.Write "ReviewList["&i&"]="""&TopicList(0,i)&"||"&TopicList(1,i)&"||"&TopicList(2,i)&""";"&vbcrlf
				End If
			Next
		End If
	End If
End Sub
%>