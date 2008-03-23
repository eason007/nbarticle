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
'= 文件名称：Member/ChangePwd.asp
'= 摘    要：会员-修改密码文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-03-22
'====================================================================

If Not EA_Pub.IsMember Then Call EA_Pub.ShowErrMsg(41, 2)

Dim EA_Mem_DBO
Set EA_Mem_DBO = New cls_Member_DBOperation

Dim Member_Id
Member_Id=EA_Pub.Mem_Info(0)
	
If LCase(Request.QueryString("action"))="savepwd" Then
	Call EA_Pub.Chk_Post

	Dim MemberInfo(4)
	Dim ReNewPassword,Password,Answer
	Dim Feedback
	Dim XML_RPC
	
	Password	  = Request.Form("OldPassword")
	MemberInfo(1) = Request.Form("Password")
	ReNewPassword = Request.Form ("RePassword")
	MemberInfo(2) = EA_Pub.SafeRequest(2,"Question",1,"",0)
	Answer		  = Request.Form("Answer")
		
	If MemberInfo(1)<>ReNewPassword Then
		ErrMsg="二次输入的新密码不相符！"
		Call EA_Pub.ShowErrMsg(52, 2)
	End If
	
	MemberInfo(0) = MD5(Password)
	MemberInfo(1) = MD5(ReNewPassword)
	MemberInfo(3) = MD5(Answer)

	Feedback=EA_DBO.Set_Member_SafetyInfo(Member_Id,MemberInfo,"")

	Select Case Feedback
	Case -1
		Call EA_Pub.ShowErrMsg(2, 2)
	Case 1
		ErrMsg="密码错误！"
		Call EA_Pub.ShowErrMsg(29, 2)
	Case 0
		MemberInfo(0) = Password
		MemberInfo(1) = ReNewPassword
		MemberInfo(3) = Answer
		MemberInfo(4) = EA_Pub.Mem_Info(1)

		Set XML_RPC = New cls_XML_RPC

		XML_RPC.OutInterfaceType = 2
		XML_RPC.StructData = MemberInfo
		XML_RPC.Start_OutInterface

		XML_RPC.Close_Obj

		Call EA_Pub.ShowErrMsg(51, 2)
	End Select
Else
	Dim Temp,Question
	Member_Id=EA_Pub.Mem_Info(0)
	
	Temp=EA_Mem_DBO.Get_MemberQuestionByAccountId(Member_Id)
	If IsArray(Temp) Then 
		Question=Temp(1,0)
	Else
		Call EA_Pub.ShowErrMsg(2, 2)
	End If
End if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="zh-cn">
<head>
<title>修改密码 - <%=EA_Pub.SysInfo(0)%></title>
<meta name="generator" content="NB文章系统(NBArticle) - <%=SysVersion%>" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="Content-Language" content="zh-cn" />
<link href="style.css" rel="stylesheet" type="text/css" />
<script src="<%=SystemFolder%>js/jsdate.js"></script>
</head>
<body id="center">

<table width="750" style="border: #A9D5F4 1px solid;" align="center">
	<form action="<%=SystemFolder%>member/changepwd.asp?action=SavePwd" method="post" name="reg">
	<tr>
		<td height="25" colspan="3" bgcolor="#DBF2FF" align="center"><strong>修改密码</strong></td>
	</tr>
	<tr>
		<td align="right">请输入您现在的密码&nbsp;</td>
		<td>&nbsp;<input name="OldPassword" type="password" id="OldPassword2" class="LoginInput" style="width: 150px;" /></td> 
		<td><font color="#FF0000">*</font>[<font color="#999999">为了系统安全要求必须验证您的身份合法性</font>]</td>
	</tr>
	<tr>
		<td align="right">请输入新的密码&nbsp;</td>
		<td>&nbsp;<input name="Password" type="password" class="LoginInput" id="Password" style="width: 150px;" /></td> 
		<td><font color="#FF0000">*</font>[<font color="#999999">大于6小于14个字符.</font><font color="#999999">不能使用特殊字符</font>]</td>
	</tr>
	<tr>
		<td align="right">请再输入一遍新的密码&nbsp;</td>
		<td>&nbsp;<input name="RePassword" type="password" class="LoginInput" id="RePassword" style="width: 150px;" /></td> 
		<td><font color="#FF0000">*</font>[<font color="#999999">确认一遍您输入的密码</font>]</td>
	</tr>
	<tr>
		<td align="right">您的密码提示问题&nbsp;</td>
		<td>&nbsp;<input name="Question" type="text" class="LoginInput" id="Question" value="<%=Question%>" style="width: 150px;" /></td>
		<td></td>
	</tr>
	<tr>
		<td align="right">请输入答案&nbsp;</td>
		<td>&nbsp;<input name="Answer" type="text" class="LoginInput" id="Answer" style="width: 150px;" /></td>
		<td></td>
	</tr>
	<tr>
	<td align="center" height="30" colspan="3"><input type="submit" value="提交修改" name="acc2" id="acc3" class="LoginInput">&nbsp;<input type="reset" value="清除重来" name="noacc2" id="noacc3" class="LoginInput"></td>
	</tr>
	</form>
</table>
