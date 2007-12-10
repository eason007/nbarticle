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
'= 文件名称：Florilegium.asp
'= 摘    要：作品集锦文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2007-07-27
'====================================================================

	Dim Author_Name,Author_Id
	Dim FieldName(1),FieldValue(1)
	Dim PageNum,PageSize,PageCount,ReCount,Column
	Dim i,WSQL
	Dim FlorilegiumArray,FlorilegiumList
	Dim PageContent
	Dim ForTotal
	
	PageNum=EA_Pub.SafeRequest(1,"page",0,1,0)
	
	Author_Name=EA_Pub.SafeRequest(1,"a_name",1,"",0)
	Author_Id=EA_Pub.SafeRequest(1,"a_id",1,"",0)
	PageSize=20
	
	FieldName(0)="a_name"
	FieldName(1)="a_id"
	
	FieldValue(0)=Author_Name
	FieldValue(1)=Author_Id
	
	ReCount=EA_DBO.Get_FlorilegiumStat(Author_Name,Author_Id)(0,0)
	PageCount=EA_Pub.Stat_Page_Total(PageSize,ReCount)
	If PageNum>PageCount And PageCount>0 Then PageNum=PageCount

	If ReCount>0 Then 
		FlorilegiumArray=EA_DBO.Get_FlorilegiumStatList(Author_Name,Author_Id,PageNum,PageSize)
		ForTotal = UBound(FlorilegiumArray,2)

		For i=0 To ForTotal
			FlorilegiumList=FlorilegiumList&"<tr height=""22"">"
			FlorilegiumList=FlorilegiumList&"<td><font color=""#FF6600"" face=""Webdings"">4</font>&nbsp;<a href="""&EA_Pub.Cov_ArticlePath(FlorilegiumArray(0,i),FlorilegiumArray(6,i),EA_Pub.SysInfo(18))&""">"&FlorilegiumArray(1,i)&"</a></td>"
			FlorilegiumList=FlorilegiumList&"<td align=center><a href="""&EA_Pub.Cov_ColumnPath(FlorilegiumArray(2,i),EA_Pub.SysInfo(18))&""">"&FlorilegiumArray(3,i)&"</a></td>"
			FlorilegiumList=FlorilegiumList&"<td align=center>"&FlorilegiumArray(4,i)&"/"&FlorilegiumArray(5,i)&"</td>"
			FlorilegiumList=FlorilegiumList&"<td align=center>"&FormatDateTime(FlorilegiumArray(6,i),2)&"</td>"
			FlorilegiumList=FlorilegiumList&"</tr>"
		Next
	Else
		FlorilegiumList=FlorilegiumList&"<tr height=""22"">"
		FlorilegiumList=FlorilegiumList&"<td colspan='3'><font color=""#FF6600"" face=""Webdings"">4</font>&nbsp;该页为空</td>"
		FlorilegiumList=FlorilegiumList&"</tr>"
	End If
%>
<html>
<head>
<title>作品集锦</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta http-equiv="Content-Language" content="zh-CN">
<style type="text/css">
body
{
	background-attachment: fixed;
	background-repeat: repeat-y;
	background-position: center;
	background-color: #FFFFFF;
	font-family: "宋体";
	font-size: 12pt; 
}
TD {
	font-family: "宋体";
	font-size: 12px; 
}
</style>
</head>
<body bgcolor="#FFFFFF" text="#000000">
<table width="99%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td bgcolor="#FFFFFF">
      <table border=0 cellpadding=3 cellspacing=2 width="100%" align=center>
        <tr> 
          <td bgcolor="#dddddd" height="30">&nbsp;<b><%=Author_Name%>的作品集锦</b></td>
        </tr>
      </table>
      <table width="100%" align=center cellpadding="1" cellspacing="2" bgcolor="#FFFFFF">
          <tr height="25" bgcolor="#e6f0ff" align=center> 
            <td>文章标题</td>
            <td width="20%">所属栏目</td>
            <td width="10%">浏览/评论</td>
            <td width="15%">投稿日期</td>
          </tr>
          <%=FlorilegiumList%>
          <tr height="25" bgcolor="#efefef" align=right> 
            <td colspan="4"><%=EA_Temp.PageList(PageCount,PageNum,FieldName,FieldValue)%>&nbsp;</td>
          </tr>
      </table></td>
  </tr>
</table>