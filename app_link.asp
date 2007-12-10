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
'= 文件名称：App_Link.asp
'= 摘    要：申请友情连接文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2006-07-27
'====================================================================

Dim Action
Action=Request.QueryString ("action")

Select Case LCase(Action)
Case "save"
	Call Save_App
Case Else
	Call Main
End Select
Call EA_Pub.Close_Obj
Set EA_Pub=Nothing

Sub Main
%>
<html>
<head>
<title><%=EA_Pub.SysInfo(0)%> - 申请友情连接</title>
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
<table width="100%" border="0" cellpadding="5" cellspacing="0" align="center" style="border: 1 solid #000000">
<form name="form1" method="post" action="?action=save">
    <tr bgcolor="#FFFFFF" valign="middle"> 
      <td align="left">站点名称：<input type="text" name="name" size="15"></td>
    </tr>
    <tr bgcolor="#FFFFFF" valign="middle"> 
      <td align="left">站点Logo：<input type="text" name="logo" value="http://"></td>
    </tr>
    <tr bgcolor="#FFFFFF" valign="middle"> 
      <td align="left">站点URL：&nbsp;<input type="text" name="url" value="http://"></td>
    </tr>
    <tr bgcolor="#FFFFFF" valign="top"> 
      <td align="left">站点简介：<textarea name="info" wrap="VIRTUAL" cols="50" rows="5"></textarea> </td>
    </tr>
    <tr bgcolor="#FFFFFF" valign="top"> 
      <td align="left">申请位置：<select name="column">
	  <option value="0">--首  页--</option>
	  <%
		Dim Level,i
		Dim TempArray
		Dim ForTotal

		TempArray=EA_DBO.Get_Column_List()
		If IsArray(TempArray) Then
			ForTotal = UBound(TempArray,2)
			For i=0 To ForTotal
				Level=(Len(TempArray(2,i))/4-1)*3
				Response.Write "<option value="""&TempArray(0,i)&""">"
				Response.Write "├"
				Response.Write String(Level,"-")
				Response.Write TempArray(1,i)&"</option>"
			Next
		End If
	  %></select></td>
    </tr>
	<tr bgcolor="#FFFFFF" valign="top"> 
      <td align="left">显示风格：<select name="style"><option value="0">文本</option><option value="1">图片</option></select></td>
    </tr>
    <tr> 
      <td valign="middle" align="center" bgcolor="#FFFFFF">&nbsp; <input type="submit" name="Submit" value="提交"> 
      </td>
    </tr>
</form>
</table>
</body>
</html>
<%
End Sub

Sub Save_App()
	Call EA_Pub.Chk_Post

	If EA_Pub.Chk_PostTime(30,"s",Session("lastpost")) Then
		ErrMsg="本系统限制数据提交间隔时间为 30 秒，请稍后再发。！"
		Call EA_Pub.ShowErrMsg(0,2)
	End If
		
	Dim LinkName,LinkImg,LinkUrl,LinkInfo,ColumnId,Style

	LinkName=EA_Pub.SafeRequest(2,"name",1,"",1)
	LinkImg=EA_Pub.SafeRequest(2,"logo",1,"",1)
	LinkUrl=EA_Pub.SafeRequest(2,"url",1,"",1)
	LinkInfo=EA_Pub.SafeRequest(2,"info",1,"",1)
	ColumnId=EA_Pub.SafeRequest(2,"column",0,0,0)
	Style=EA_Pub.SafeRequest(2,"style",0,0,0)
	FoundErr=False
	
	If LinkName="" Or Len(LinkName)>50 Then FoundErr=True
	If Len(LinkImg)>150 Then FoundErr=True
	If LinkUrl="" Or Len(LinkUrl)>150 Then FoundErr=True
	
	If Not FoundErr Then
		Call EA_DBO.Set_FriendList_Insert(LinkName,LinkImg,LinkUrl,LinkInfo,ColumnId,Style,0,0)
		
		ErrMsg="申请已记录，请等待站长开通。谢谢您的支持。"
		Session("lastpost")=Now()
	Else
		ErrMsg="您填写的资料不符合要求，请重新填写。"
	End If

	Call EA_Pub.ShowErrMsg(0,2)
End Sub
%>