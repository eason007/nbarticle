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
'= 文件名称：MoreLink.asp
'= 摘    要：列表页文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2005-11-30
'====================================================================

Dim TopicList
Dim i,j,TagId
Dim OutStr
Dim ForTotal

TagId=-1
j=0

TopicList=EA_DBO.Get_FriendList_All()

If IsArray(TopicList) Then
	OutStr=OutStr&"<table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" bgcolor=""efefef"">"
	ForTotal = UBound(TopicList,2)

	For i=0 To ForTotal
		If TagId<>TopicList(0,i) Then 
			TagId=TopicList(0,i)
			OutStr=OutStr&"<tr>"
			OutStr=OutStr&"<td colspan=""5"" height=""22"" bgcolor=""efefef"">&nbsp;<img src=""Images/bullet1.gif"" width=""16"" height=""16"" align=""absmiddle"">[<b>"&TopicList(1,i)&"</b>]下的友情连接</td></tr><tr bgcolor=""ffffff"">"
			i=i-1
			j=0
		Else
			If TopicList(6,i)=1 Then 
				OutStr=OutStr&"<td width=""20%"" align=""center"">&nbsp;<a href="""&TopicList(2,i)&""" title="""&TopicList(3,i)&""" target=""_blank""><img src="""&TopicList(4,i)&""" border=0 align=""absmiddle"" alt="""&TopicList(3,i)&""" width=""88"" height=""31""></a></td>"
			Else
				OutStr=OutStr&"<td width=""20%"" align=""center"" height=""22"">&nbsp;<a href="""&TopicList(2,i)&""" title="""&TopicList(3,i)&""" target=""_blank"">"&TopicList(5,i)&"</a></td>"
			End If
			If i+1<=UBound(TopicList,2) Then 
				If TopicList(6,i)<>TopicList(6,i+1) Then OutStr=OutStr&"</tr><tr bgcolor=""ffffff"">"
			End If			
			If (j+1) Mod 5 =0 Then 
				OutStr=OutStr&"</tr><tr bgcolor='ffffff'>"
				j=0
			Else
				j=j+1
			End If
		End If
	Next
	OutStr=OutStr&"</tr></table>"
Else
	OutStr=OutStr&"·暂无友情连接"
End If

Call EA_Pub.Close_Obj
Set EA_Pub=Nothing
%>
<html>
<head>
<title>友情连接列表</title>
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
          <td bgcolor="#dddddd" height="30">&nbsp;<b>友情连接列表</b></td>
        </tr>
      </table>
      <table width="100%" align=center cellpadding="1" cellspacing="2" bgcolor="#FFFFFF">
          <tr height="25" bgcolor="#e6f0ff" align=center> 
            <td><%=OutStr%></td>
          </tr>
      </table></td>
  </tr>
</table>