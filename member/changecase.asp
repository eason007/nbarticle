<!--#Include File="../include/inc.asp"-->
<!--#Include File="../include/cls_xml_rpc.asp"-->
<!--#Include File="cls_db.asp"-->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：Member/ChangeCase.asp
'= 摘    要：会员-修改个人资料文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-03-22
'====================================================================

If Not EA_Pub.IsMember Then Call EA_Pub.ShowErrMsg(41, 2)

Dim MemberInfo

If LCase(Request.QueryString ("action"))="savedata" Then
	Call EA_Pub.Chk_Post

	Dim Feedback
	Dim Password
	Dim XML_RPC
	ReDim MemberInfo(10)
	
	Password	  = Request.Form("Password")
	MemberInfo(1) = EA_Pub.SafeRequest(2,"Email",1,"",0)
	MemberInfo(2) = EA_Pub.SafeRequest(2,"Sex",0,0,0)
	MemberInfo(3) = EA_Pub.SafeRequest(2,"HomePage",1,"",0)
	MemberInfo(4) = EA_Pub.SafeRequest(2,"QQ",0,0,0)
	MemberInfo(5) = EA_Pub.SafeRequest(2,"ICQ",0,0,0)
	MemberInfo(6) = EA_Pub.SafeRequest(2,"MSN",1,"",0)
	MemberInfo(7) = EA_Pub.SafeRequest(2,"UserName",1,"",0)
	MemberInfo(8) = EA_Pub.SafeRequest(2,"Birthday",2,Now(),0)
	MemberInfo(9) = EA_Pub.SafeRequest(2,"Comefrom",1,"",0)
	
	MemberInfo(0) = MD5(Password)

	Feedback=EA_DBO.Set_Member_Info(EA_Pub.Mem_Info(0),MemberInfo,"")

	Select Case Feedback
	Case 1
		Call EA_Pub.ShowErrMsg(38, 2)
	Case 0
		MemberInfo(0)	= Password
		MemberInfo(10)	= EA_Pub.Mem_Info(1)

		Set XML_RPC = New cls_XML_RPC

		XML_RPC.OutInterfaceType = 3
		XML_RPC.StructData = MemberInfo
		XML_RPC.Start_OutInterface

		XML_RPC.Close_Obj

		Call EA_Pub.ShowErrMsg(51, 2)
	Case -1
		Call EA_Pub.ShowErrMsg(29, 2)
	Case 2
		Call EA_Pub.ShowErrMsg(2, 2)
	End Select
Else
	MemberInfo=EA_DBO.Get_MemberInfo(EA_Pub.Mem_Info(0))
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="zh-cn">
<head>
<title>修改个人资料 - <%=EA_Pub.SysInfo(0)%></title>
<meta name="generator" content="NB文章系统(NBArticle) - <%=SysVersion%>" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="Content-Language" content="zh-cn" />
<link href="style.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="<%=SystemFolder%>js/jsdate.js"></script>
</head>
<body id="center">

<table width="750" style="border: #A9D5F4 1px solid;" align="center">
	<form action="<%=SystemFolder%>member/changecase.asp?action=savedata" method="post" name="reg" id="reg">
	<tr>
		<td height="25" colspan="3" bgcolor="#DBF2FF" align="center"><strong>修改个人资料</strong></td>
	</tr>
	<tr>
		<td align="right">请先输入您的密码&nbsp;</td>
		<td>&nbsp;<input name="Password" type="password" id="Password" class="LoginInput" style="width: 150px;" /></td> 
		<td><font color="#FF0000">*</font>[<font color="#999999">为了保护您的个人信息，您必须输入你的密码以通过身份验证</font>]</td>
	</tr>
	<tr>
		<td align="right">电子信箱&nbsp;</td>
		<td colspan="2">&nbsp;<input name="Email" type="text" class="LoginInput" id="Email" value="<%=MemberInfo(3,0)%>" style="width: 150px;" /></td>
	</tr> 
	<tr> 
		<td height="25" align="right">您的性别&nbsp;</td>
		<td colspan="2">&nbsp;男
		<input name="sex" type="radio" value="1" <%If MemberInfo(2,0) Then Response.Write "checked"%> />
		女
		<input name="sex" type="radio" value="0" <%If Not MemberInfo(2,0) Then Response.Write "checked"%> /></td>
	</tr> 
	<tr>
		<td height="24" align="right">个人网站地址&nbsp;</td>
		<td colspan="2">&nbsp;<input name="HomePage" type="text" class="LoginInput" id="HomePage" value="<%=MemberInfo(10,0)%>" style="width: 150px;" /></td>
	</tr>
	<tr>
		<td height="25" align="right">腾讯QQ号码&nbsp;</td>
		<td colspan="2">&nbsp;<input name="QQ" type="text" class="LoginInput" id="QQ" value="<%=MemberInfo(11,0)%>" style="width: 150px;" /></td>
	</tr> 
	<tr>
		<td height="25" align="right">ICQ号码&nbsp;</td>
		<td colspan="2">&nbsp;<input name="ICQ" type="text" class="LoginInput" id="ICQ" value="<%=MemberInfo(12,0)%>" style="width: 150px;" /></td>
	</tr>
	<tr>
		<td height="25" align="right">MSN帐户&nbsp;</td>
		<td colspan="2">&nbsp;<input name="MSN" type="text" class="LoginInput" id="MSN" value="<%=MemberInfo(13,0)%>" style="width: 150px;" /></td>
	</tr>
	<tr>
		<td height="25" align="right">真实姓名&nbsp;</td>
		<td colspan="2">&nbsp;<input name="UserName" type="text" class="LoginInput" id="UserName" value="<%=MemberInfo(6,0)%>" style="width: 150px;" /></td>
	</tr>
	<tr>
		<td height="25" align="right">出生日期&nbsp;</td>
		<td colspan="2">&nbsp;<input type="text" name="Birthday" id="Birthday" size="10" readonly onclick="SD(this,'Birthday');" class="LoginInput" value="<%=FormatDateTime(MemberInfo(7,0),2)%>" /></td>
	</tr> 
	<tr>
		<td height="25" align="right">来自&nbsp;</td>
		<td colspan="2">&nbsp;<textarea name="ComeFrom" cols="30" id="ComeFrom"><%=MemberInfo(14,0)%></textarea></td>
	</tr>
	<tr>
		<td align="center" colspan="3" height="30"><input type="submit" value="提交修改" name="acc" id="acc" class="LoginInput" />&nbsp;<input type="reset" value="清除重来" name="noacc" id="noacc" class="LoginInput" /></td>
	</tr>
  </form>
</table> 
