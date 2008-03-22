<!--#Include File="../include/inc.asp"-->
<!--#Include File="cls_db.asp"-->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：Member/FavList.asp
'= 摘    要：会员-收藏夹列表文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-03-22
'====================================================================

If Not EA_Pub.IsMember Then Call EA_Pub.ShowErrMsg(41, 2)

Dim EA_Mem_DBO
Set EA_Mem_DBO = New cls_Member_DBOperation

Dim Action
Dim Member_Id
Action=Request.QueryString ("action")
Member_Id=EA_Pub.Mem_Info(0)

Select Case LCase(Action)
Case "add"
	Call Add_Fav
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
	Dim FavArray,FavList
	Dim i
	Dim Url
	
	PageNum=EA_Pub.SafeRequest(1,"page",0,1,3)
	PageSize=15
	
	ReCount=EA_Mem_DBO.Get_MemberFavTotalByAccountId(Member_Id)(0,0)
	PageCount=EA_Pub.Stat_Page_Total(PageSize,ReCount)
	If PageNum>PageCount And PageCount>0 Then PageNum=PageCount
	
	If ReCount>0 Then 
		FavArray=EA_DBO.Get_MemberFavListByAccountId(Member_Id,PageNum,PageSize)
		
		For i=0 To UBound(FavArray,2)
			FavList=FavList&"<tr height=""22"">"
			FavList=FavList&"<td><font color=""#FF6600"" face=""Webdings"">4</font>&nbsp;<a href="""&EA_Pub.Cov_ArticlePath(FavArray(0,i),FavArray(1,i),EA_Pub.SysInfo(18))&""">"&FavArray(2,i)&"</a></td>"
			FavList=FavList&"<td align=center>"&FavArray(5,i)&"</td>"
			FavList=FavList&"<td align=center>"&FavArray(4,i)&"</td>"
			FavList=FavList&"<td align=center><a href='?action=del&postid="&FavArray(3,i)&"' onclick=""{if(confirm('确认删除？')){return true;}return false;}"">删除</a></td>"
			FavList=FavList&"</tr>"
		Next
	Else
		FavList=FavList&"<tr height=""22"">"
		FavList=FavList&"<td colspan='4'><font color=""#FF6600"" face=""Webdings"">4</font>&nbsp;该页为空</td>"
		FavList=FavList&"</tr>"
	End If

	Url = "?&page=$page"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="zh-cn">
<head>
<title>我的收藏夹 - <%=EA_Pub.SysInfo(0)%></title>
<meta name="generator" content="NB文章系统(NBArticle) - <%=SysVersion%>" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="Content-Language" content="zh-cn" />
<link href="style.css" rel="stylesheet" type="text/css" />
</head>
<body id="center">

<table width="750" style="border: #A9D5F4 1px solid;" align="center">
	<tr>
		<td height="25" colspan="3" bgcolor="#DBF2FF">&nbsp;<strong>我的收藏夹</strong></td>
	</tr>
	<tr>
		<td>
			<table width="100%">
				<tr height="22" bgcolor="#e6f0ff" align="center"> 
					<td>文章标题</td>
					<td width="10%">文章作者</td>
					<td width="20%">收藏日期</td>
					<td width="10%">操 作</td>
				</tr>
				<%=FavList%>
				<tr align="right"> 
					<td colspan="4"><%=EA_Temp.PageList(PageCount,PageNum,Url)%>&nbsp;</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<%
End Sub

Sub Add_Fav
	Call EA_Pub.Chk_Post
	
	Dim Fav_Id
	Dim Feedback
	Fav_Id=EA_Pub.SafeRequest(3,"postid",0,0,3)
	
	Feedback=EA_DBO.Set_AddFav(Fav_Id,Member_Id)
	Select Case Feedback
	Case 0
		ErrMsg="成功收藏文章！"
	Case 1
		ErrMsg="该文章你已收藏过。"
	Case -1
		ErrMsg="您的收藏夹中的数目已到达上限，请删除其他收藏文章后再进行添加收藏操作！"&Member_Id
	End Select
	Call EA_Pub.ShowErrMsg(0,2)
End Sub

Sub Del_Fav
	Call EA_Pub.Chk_Post
	
	Dim Fav_Id
	Fav_Id=EA_Pub.SafeRequest(3,"postid",0,0,3)
	
	Call EA_DBO.Del_MemberFav(Fav_Id,Member_Id)
	
	Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub
%>