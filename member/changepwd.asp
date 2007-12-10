<!--#Include File="../conn.asp" -->
<!--#Include File="../include/inc.asp"-->
<!--#Include File="../include/cls_xml_rpc.asp"-->
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
'= 最后日期：2005-10-19
'====================================================================

If Not EA_Pub.IsMember Then Call EA_Pub.ShowErrMsg(10,1)

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
		Call EA_Pub.ShowErrMsg(0,2)
	End If
	
	MemberInfo(0) = MD5(Password)
	MemberInfo(1) = MD5(ReNewPassword)
	MemberInfo(3) = MD5(Answer)

	Feedback=EA_DBO.Set_Member_SafetyInfo(Member_Id,MemberInfo,"")

	Select Case Feedback
	Case -1
		ErrMsg="用户不存在！"
		Call EA_Pub.ShowErrMsg(0,2)
	Case 1
		ErrMsg="密码错误！"
		Call EA_Pub.ShowErrMsg(0,2)
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

		Call EA_Pub.ShowSusMsg(3,0)
	End Select
Else
	Dim Temp,Question
	Member_Id=EA_Pub.Mem_Info(0)
	
	Temp=EA_DBO.Get_MemberQuestionByAccountId(Member_Id)
	If IsArray(Temp) Then 
		Question=Temp(1,0)
	Else
		ErrMsg="用户不存在！"
		Call EA_Pub.ShowErrMsg(0,2)
	End If
End if
%>
<html>
<head>
<title>修改密码</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta http-equiv="Content-Language" content="zh-CN">
<link href="style.css" rel="stylesheet" type="text/css" />
<script src="../js/jsdate.js"></script>
</head>
<body bgcolor="#FFFFFF" text="#000000"> 
<table width="762" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC"> 
  <form action="?action=SavePwd" method="post" name="reg" target="_self"> 
    <tr> 
      <td align="center" valign="top" bgcolor="#FFFFFF"><table border=0 cellpadding=3 cellspacing=2 width="100%" align=center> 
          <tr> 
            <td bgcolor="#dddddd" height="30">&nbsp;<b>修改密码</b></td> 
          </tr> 
        </table> 
        <table width="100%" cellspacing="1" cellpadding="2"> 
          <tr> 
            <td width="30%" height="30" bgcolor="#e6f0ff" align="right"><font color="#FF6600" face="Webdings">4</font>请输入您现在的密码&nbsp;</td> 
            <td bgcolor="#FFFFFF">&nbsp;<input name="OldPassword" type="password" id="OldPassword2" class="LoginInput" size="15" maxlength="20"> 
              <font color="#FF0000">*</font>[<font color="#999999">为了系统安全要求必须验证您的身份合法性</font>] </td> 
          </tr> 
          <tr> 
            <td height="30" bgcolor="#e6f0ff" align="right"><font color="#FF6600" face="Webdings">4</font>请输入新的密码&nbsp;</td> 
            <td>&nbsp;<input name="Password" type="password" class="LoginInput" id="Password" size="15" maxlength="20"> 
              <font color="#FF0000">*</font>[<font color="#999999">大于6小于14个字符.</font><font color="#999999">不能使用特殊字符</font>]</td> 
          </tr> 
          <tr> 
            <td height="30" bgcolor="#e6f0ff" align="right"><font color="#FF6600" face="Webdings">4</font>请再输入一遍新的密码&nbsp;</td> 
            <td bgcolor="#FFFFFF">&nbsp;<input name="RePassword" type="password" class="LoginInput" id="RePassword" size="15" maxlength="20"> 
              <font color="#FF0000">*</font>[<font color="#999999">确认一遍您输入的密码</font>]</td> 
          </tr> 
          <tr> 
            <td height="30" bgcolor="#e6f0ff" align="right"><font color="#FF6600" face="Webdings">4</font>您的密码提示问题&nbsp;</td> 
            <td>&nbsp;<input name="Question" type="text" class="LoginInput" id="Question" size="20" maxlength="40" value="<%=Question%>"></td> 
          </tr> 
          <tr> 
            <td height="30" bgcolor="#e6f0ff" align="right"><font color="#FF6600" face="Webdings">4</font>请输入答案&nbsp;</td> 
            <td>&nbsp;<input name="Answer" type="text" class="LoginInput" id="Answer" size="20" maxlength="40"></td> 
          </tr> 
          <tr> 
            <td align="center" bgcolor="#efefef" height="30" colspan="2"><input type="submit" value="提交修改" name="acc2" id="acc3" class="LoginInput"> 
&nbsp; 
              <input type="reset" value="清除重来" name="noacc2" id="noacc3" class="LoginInput"></td> 
          </tr> 
        </table></td> 
    </tr> 
  </form> 
</table> 
