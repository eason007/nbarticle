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
'= 文件名称：ViewPlacard.asp
'= 摘    要：浏览公告文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2005-11-30
'====================================================================

Dim PostId
Dim PlacardInfo
Dim Title,Content,OverTime,AddTime

PostId=EA_Pub.SafeRequest(3,"postid",0,0,0)

PlacardInfo=EA_DBO.Get_PlacardInfo(PostId)
If IsArray(PlacardInfo) Then 
	Title=PlacardInfo(0,0)
	Content=PlacardInfo(1,0)
	OverTime=FormatDateTime(PlacardInfo(2,0),2)
	If DateDiff("d",OverTime,Now())>0 Then OverTime=OverTime&" [已过期]"
	AddTime=FormatDateTime(PlacardInfo(3,0),2)
End If
%>
<html>
<head>
<title><%=EA_Pub.SysInfo(0)%> - 站点公告</title>
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
<SCRIPT language=JavaScript>
function open_window(){
	window.opener.location.href='moreplacard.asp';
	window.close();
}
</SCRIPT>
</head>
<body bgcolor="#FFFFFF" text="#000000">
<table width="100%" border="0" cellpadding="3" cellspacing="1" align="center" style="border: 1 solid #000000" bgcolor="dddddd">
    <tr bgcolor="#FFFFFF" valign="middle" height="25"> 
      <td align="center" colspan="2"><%=Title%></td>
    </tr>
    <tr bgcolor="#efefef" height="22"> 
      <td align="left" width="50%">&nbsp;加入时间：<%=AddTime%></td>
      <td align="left">&nbsp;到期时间：<%=OverTime%></td>
    </tr>
    <tr height="100"> 
      <td bgcolor="#FFFFFF" colspan="2" valign="top"><%=EA_Pub.Full_HTMLFilter(Content)%></td>
    </tr>
	<tr height="25"> 
      <td bgcolor="#efefef" colspan="2" align="center"><a href="javascript:window.close();">关闭窗口</a> | <a href="#" onclick="open_window()">更多公告</a></td>
    </tr>
</table>
</body>
</html>
<%
Call EA_Pub.Close_Obj
Set EA_Pub=Nothing
%>