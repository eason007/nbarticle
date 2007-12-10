<!--#Include File="../conn.asp" -->
<!--#Include File="../include/inc.asp"-->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：Member/MyAppear.asp
'= 摘    要：会员-我的投稿列表文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2005-05-16
'====================================================================

If Not EA_Pub.IsMember Then Call EA_Pub.ShowErrMsg(10,1)

Dim Action
Dim Member_Id
Action=Request.QueryString ("action")
Member_Id=EA_Pub.Mem_Info(0)

Select Case LCase(Action)
Case "del"
	Call Del_Fav
Case Else
	Call Main
End Select
Call EA_Pub.Close_Obj
Set EA_Pub=Nothing

Sub Main
	Dim PageContent
	Dim PageNum,PageSize,PageCount
	Dim ReCount
	Dim AppearArray,AppearList
	Dim i
	Dim FieldName(0),FieldValue(0)
	
	PageNum=EA_Pub.SafeRequest(1,"page",0,1,3)
	PageSize=15
	
	ReCount=EA_DBO.Get_Member_AppearTotal(Member_Id)(0,0)
	PageCount=EA_Pub.Stat_Page_Total(PageSize,ReCount)
	If PageNum>PageCount And PageCount>0 Then PageNum=PageCount
	
	If ReCount>0 Then 
		AppearArray=EA_DBO.Get_MemberAppearList(Member_Id,PageNum,PageSize)
		
		For i=0 To UBound(AppearArray,2)
			AppearList=AppearList&"<tr height=""22"">"
			AppearList=AppearList&"<td><font color=""#FF6600"" face=""Webdings"">4</font>&nbsp;<a href="""&EA_Pub.Cov_ArticlePath(AppearArray(0,i),AppearArray(6,i),EA_Pub.SysInfo(18))&""" target=""_blank"">"
			AppearList=AppearList&EA_Pub.Add_ArticleColor(AppearArray(1,i),AppearArray(2,i))&"</a></td>"
			AppearList=AppearList&"<td align=center>"&AppearArray(3,i)&"</td>"
			AppearList=AppearList&"<td align=center>"&AppearArray(4,i)&"/"&AppearArray(5,i)&"</td>"
			AppearList=AppearList&"<td align=center>"&FormatDateTime(AppearArray(6,i),2)&"</td>"
			AppearList=AppearList&"<td align=center>"
			If AppearArray(7,i) Then
				AppearList=AppearList&"<font color=green>Y</font>"
			Else
				AppearList=AppearList&"<font color=red>N</font>"
			End If
			AppearList=AppearList&"</td>"
			AppearList=AppearList&"<td align=center><a href=""appear.asp?postid="&AppearArray(0,i)&""" target=""_blank"">修改</a> <a href='?action=del&postid="&AppearArray(0,i)&"&columnid="&AppearArray(8,i)&"&ispass="&CInt(AppearArray(7,i))&"' onclick=""{if(confirm('确认删除？')){return true;}return false;}"">删除</a></td>"
			AppearList=AppearList&"</tr>"
		Next
	Else
		AppearList=AppearList&"<tr height=""22"">"
		AppearList=AppearList&"<td colspan='6'><font color=""#FF6600"" face=""Webdings"">4</font>&nbsp;该页为空</td>"
		AppearList=AppearList&"</tr>"
	End If
%>
<html>
<head>
<title>我的文集</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta http-equiv="Content-Language" content="zh-CN">
<link href="style.css" rel="stylesheet" type="text/css" />
</head>
<body bgcolor="#FFFFFF" text="#000000">
<table width="99%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td bgcolor="#FFFFFF">
      <table border=0 cellpadding=3 cellspacing=2 width="100%" align=center>
        <tr> 
          <td bgcolor="#dddddd" height="30">&nbsp;<b>我的文集</b></td>
        </tr>
      </table>
      <table width="100%" align=center cellpadding="1" cellspacing="2" bgcolor="#FFFFFF">
          <tr height="25" bgcolor="#e6f0ff" align=center> 
            <td>文章标题</td>
            <td width="10%">栏目</td>
            <td width="10%">浏览/评论</td>
            <td width="10%">投稿日期</td>
            <td width="5%">状态</td>
            <td width="10%">操 作</td>
          </tr>
          <%=AppearList%>
          <tr height="25" bgcolor="#efefef" align=right> 
            <td colspan="6"><%=EA_Temp.PageList(PageCount,PageNum,FieldName,FieldValue)%>&nbsp;</td>
          </tr>
      </table></td>
  </tr>
</table>
<%
End Sub

Sub Del_Fav
	Call EA_Pub.Chk_Post
	
	Dim Appear_Id,Column_Id,IsPass
	
	Appear_Id=EA_Pub.SafeRequest(3,"postid",0,0,3)
	Column_Id=EA_Pub.SafeRequest(3,"columnid",0,0,3)
	IsPass=EA_Pub.SafeRequest(3,"ispass",0,0,3)
	
	Call EA_DBO.Del_MemberAppear(Appear_Id,Member_Id,Column_Id,IsPass)

	Application.Lock 
	Application(sCacheName&"IsFlush")=1
	Application.UnLock 
	
	Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub
%>