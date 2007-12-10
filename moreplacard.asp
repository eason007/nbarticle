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
'= 文件名称：MorePlacard.asp
'= 摘    要：公告列表文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2006-07-27
'====================================================================

Dim PageNum,PageSize,RCount,PageCount
Dim FieldName(1),FieldValue(1)
Dim PlacardArray,PlacardList
Dim i
Dim ForTotal

PageNum=EA_Pub.SafeRequest(1,"page",0,1,0)
PageSize=20

RCount=EA_DBO.Get_PlacardStat()(0,0)
PageCount=EA_Pub.Stat_Page_Total(PageSize,RCount)
If PageNum>PageCount And PageCount>0 Then PageNum=PageCount

	If RCount>0 Then 
		PlacardArray=EA_DBO.Get_PlacardList(PageNum,PageSize)
		ForTotal = UBound(PlacardArray,2)

		For i=0 To ForTotal
			PlacardList=PlacardList&"<tr height=""22"">"
			PlacardList=PlacardList&"<td><font color=""#FF6600"" face=""Webdings"">4</font>&nbsp;<a href=""#"" onclick=""javascript:window.open('viewplacard.asp?postid="&PlacardArray(0,i)&"','','scrollbars=yes,height=350,width=550')"">"&PlacardArray(1,i)&"</a></td>"
			PlacardList=PlacardList&"<td align=center>"&PlacardArray(2,i)&"</td>"
			PlacardList=PlacardList&"<td align=center>"&PlacardArray(3,i)&"</td>"
			PlacardList=PlacardList&"</tr>"
		Next
	Else
		PlacardList=PlacardList&"<tr height=""22"">"
		PlacardList=PlacardList&"<td colspan='3'><font color=""#FF6600"" face=""Webdings"">4</font>&nbsp;该页为空</td>"
		PlacardList=PlacardList&"</tr>"
	End If
%>
<html>
<head>
<title>站点公告</title>
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
          <td bgcolor="#dddddd" height="30">&nbsp;<b>站点公告列表</b></td>
        </tr>
      </table>
      <table width="100%" align=center cellpadding="1" cellspacing="2" bgcolor="#FFFFFF">
          <tr height="25" bgcolor="#e6f0ff" align=center> 
            <td>标题</td>
            <td width="20%">发布日期</td>
            <td width="20%">结束日期</td>
          </tr>
          <%=PlacardList%>
          <tr height="25" bgcolor="#efefef" align=right> 
            <td colspan="3"><%=EA_Temp.PageList(PageCount,PageNum,FieldName,FieldValue)%>&nbsp;</td>
          </tr>
      </table></td>
  </tr>
</table>