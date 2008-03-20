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
'= 最后日期：2006-08-11
'====================================================================

If Not EA_Pub.IsMember Then Call EA_Pub.ShowErrMsg(10,1)

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
		ErrMsg="该E-Mail已被人注册，请重新填写。"
		Call EA_Pub.ShowErrMsg(0,2)
	Case 0
		MemberInfo(0)	= Password
		MemberInfo(10)	= EA_Pub.Mem_Info(1)

		Set XML_RPC = New cls_XML_RPC

		XML_RPC.OutInterfaceType = 3
		XML_RPC.StructData = MemberInfo
		XML_RPC.Start_OutInterface

		XML_RPC.Close_Obj

		Call EA_Pub.ShowSusMsg(3,0)
	Case -1
		ErrMsg="密码错误。"
		Call EA_Pub.ShowErrMsg(0,2)
	Case 2
		Call EA_Pub.ShowErrMsg(18,1)
	End Select
Else
	MemberInfo=EA_DBO.Get_MemberInfo(EA_Pub.Mem_Info(0))
End If
%>
<html>
<head>
<title>修改个人资料</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta http-equiv="Content-Language" content="zh-CN">
<link href="style.css" rel="stylesheet" type="text/css" />
<script src="../js/public.js"></script> 
<script src="../js/jsdate.js"></script> 
</head>
<body bgcolor="#FFFFFF" text="#000000">
<table width="760" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
<form action="?action=savedata" method="post" name="reg" id="reg"> 
  <tr> 
    <td align="center" valign="top" bgcolor="#FFFFFF">
        <table border=0 cellpadding=3 cellspacing=2 width="100%" align=center> 
          <tr> 
            <td bgcolor="#dddddd" height="30">&nbsp;<b>修改个人资料</b></td> 
          </tr> 
        </table> 
        <table align=center cellpadding="1" cellspacing="2" bgcolor="#FFFFFF" width="100%"> 
          <tbody> 
            <tr> 
              <td height="30" align="right" bgcolor="#e6f0ff" width="30%"><font color="#FF6600" face="Webdings">4</font>请先输入您的密码&nbsp;</td> 
              <td height="30" bgcolor="#FFFFFF">&nbsp;<input name="Password" type="password" id="Password" class="LoginInput" size="15" maxlength="20"> 
                <font color="#FF0000">*</font>[<font color="#999999">为了保护您的个人信息，您必须输入你的密码以通过身份验证</font>] </td> 
            </tr> 
            <tr> 
              <td height="25" align="right" bgcolor="#e6f0ff">电子信箱&nbsp;</td> 
              <td>&nbsp;<input name="Email" type="text" class="LoginInput" id="Email" value="<%=MemberInfo(3,0)%>" size="30" maxlength="30"> </td> 
            </tr> 
            <tr> 
              <td height="25" align="right" bgcolor="#e6f0ff">您的性别&nbsp;</td> 
              <td>&nbsp;男
                <input name="sex" type="radio" value="1" <%If MemberInfo(2,0) Then Response.Write "checked"%>> 
                女
                <input name="sex" type="radio" value="0" <%If Not MemberInfo(2,0) Then Response.Write "checked"%>> </td> 
            </tr> 
            <tr> 
              <td height="24" align="right" bgcolor="#e6f0ff">个人网站地址&nbsp;</td> 
              <td>&nbsp;<input name="HomePage" type="text" class="LoginInput" id="HomePage" value="<%=MemberInfo(10,0)%>" size="30" maxlength="50"></td> 
            </tr> 
            <tr> 
              <td height="25" align="right" bgcolor="#e6f0ff">腾讯QQ号码&nbsp;</td> 
              <td>&nbsp;<input name="QQ" type="text" class="LoginInput" id="QQ" value="<%=MemberInfo(11,0)%>" size="15" maxlength="20"></td> 
            </tr> 
            <tr> 
              <td height="25" align="right" bgcolor="#e6f0ff">ICQ号码&nbsp;</td> 
              <td>&nbsp;<input name="ICQ" type="text" class="LoginInput" id="ICQ" value="<%=MemberInfo(12,0)%>" size="15" maxlength="20"></td> 
            </tr> 
            <tr> 
              <td height="25" align="right" bgcolor="#e6f0ff">MSN帐户&nbsp;</td> 
              <td>&nbsp;<input name="MSN" type="text" class="LoginInput" id="MSN" value="<%=MemberInfo(13,0)%>" size="20" maxlength="40"></td> 
            </tr> 
            <tr> 
              <td height="25" align="right" bgcolor="#e6f0ff">真实姓名&nbsp;</td> 
              <td>&nbsp;<input name="UserName" type="text" class="LoginInput" id="UserName" value="<%=MemberInfo(6,0)%>" size="15" maxlength="20"> </td> 
            </tr> 
            <tr> 
              <td height="25" align="right" bgcolor="#e6f0ff">出生日期&nbsp;</td> 
              <td>&nbsp;<input type="text" name="Birthday" id="Birthday" size="10" readonly class="LoginInput" value="<%=FormatDateTime(MemberInfo(7,0),2)%>">&nbsp;<a href="javascript:vod()" onclick="SD(this,'Birthday');"><img border="0" src="../images/public/date_picker.gif" width="30" height="19"></a></td> 
            </tr> 
            <tr> 
              <td height="25" align="right" bgcolor="#e6f0ff">来自&nbsp;</td> 
              <td>&nbsp;<textarea name="ComeFrom" cols="30" id="ComeFrom"><%=MemberInfo(14,0)%></textarea></td> 
            </tr> 
          </tbody> 
        </table> 
        <table width="100%"> 
          <tr> 
            <td align="center" bgcolor="#efefef" height="30"><input type="submit" value="提交修改" name="acc" id="acc" class="LoginInput">&nbsp;<input type="reset" value="清除重来" name="noacc" id="noacc" class="LoginInput"></td> 
          </tr> 
        </table> 
      </td> 
  </tr> 
  </form>
</table> 
