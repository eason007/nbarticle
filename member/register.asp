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
'= 文件名称：Member/Register.asp
'= 摘    要：会员-注册文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-03-22
'====================================================================

If EA_Pub.SysInfo(7)="0" Then Call EA_Pub.ShowErrMsg(34, 0)

If LCase(Request.Querystring("action"))="savedata" Then
	'0=account,1=password,2=email,3=question,4=answer
	'5=sex,6=homepage,7=qq,8=icq,9=msn
	'10=rename,11=date,12=uform,13=ip
	Dim MemberInfo(14)
	Dim Password,Answer
	Dim Feedback
	Dim XML_RPC
	
	Call EA_Pub.Chk_Post

	MemberInfo(0)	= EA_Pub.SafeRequest(2,"UserName",1,"",0)
	Password		= Request.Form("Password")
	MemberInfo(2)	= EA_Pub.SafeRequest(2,"Email",1,"",0)
	MemberInfo(3)	= EA_Pub.SafeRequest(2,"Question",1,"",0)
	Answer			= Request.Form("Ans")

	MemberInfo(5) = EA_Pub.SafeRequest(2,"sex",0,0,0)
	MemberInfo(6) = EA_Pub.SafeRequest(2,"HomePage",1,"",0)
	MemberInfo(7) = EA_Pub.SafeRequest(2,"QQ",0,0,0)
	MemberInfo(8) = EA_Pub.SafeRequest(2,"ICQ",0,0,0)
	MemberInfo(9) = EA_Pub.SafeRequest(2,"MSN",1,"",0)

	MemberInfo(10) = EA_Pub.SafeRequest(2,"ReName",1,"",0)
	MemberInfo(11) = EA_Pub.SafeRequest(2,"date",2,Now(),0)
	MemberInfo(12) = EA_Pub.SafeRequest(2,"UForm",1,"",0)
	MemberInfo(13) = EA_Pub.Get_UserIp

	MemberInfo(1) = MD5(Password)
	MemberInfo(4) = MD5(Answer)

	Feedback=EA_DBO.Set_RegistrationMember(MemberInfo)

	Select Case Feedback
	Case -1
		Call EA_Pub.ShowErrMsg(38, 2)
	Case 0, 1
		MemberInfo(1) = Password
		MemberInfo(4) = Answer

		Set XML_RPC = New cls_XML_RPC

		XML_RPC.OutInterfaceType = 1
		XML_RPC.StructData = MemberInfo
		XML_RPC.Start_OutInterface

		XML_RPC.Close_Obj

		Application.Lock 
		Application(sCacheName&"IsFlush")=1
		Application.UnLock 
		Call EA_Pub.ShowErrMsg(40, 2)
	Case 2
		Call EA_Pub.ShowErrMsg(35, 2)
	End Select	
End if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="zh-cn">
<title>注册会员 - <%=EA_Pub.SysInfo(0)%></title>
<meta name="generator" content="NB文章系统(NBArticle) - <%=SysVersion%>" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="Content-Language" content="zh-cn" />
<meta name="robots" content="nofollow" />
<link href="style.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="../js/jsdate.js"></script>
<script type="text/javascript">
var Step = 1;
var goreg = true;

function accaction(){
	if(Step==1){
		xieyi.style.display = "none";
		acc.value = "确认申请";
		noacc.value = "放弃申请";
		mustinfo.style.display = "block";
		Step = 2
	}
	else{
		checkerr("\n请输入用户名\n\n10个字符以下",document.reg.UserName,10,1);
		if(goreg==true){checkerr("\n请输入密码\n\n6-14个字符内",document.reg.Password,14,6);}
		if(goreg==true){checkerr("\n请再输入一次密码\n\n6-14个字符内",document.reg.RePassWord,14,6);}
		if(goreg==true){
			if(document.reg.Password.value!=document.reg.RePassWord.value){
			alert("\n您两次输入的密码不一致</b>");
			document.reg.RePassWord.focus();
			goreg = false;
		}
		else{
			goreg=true
		}
	}

	if(goreg==true){checkerr("\n请输入正确的电子信箱\n",document.reg.Email,30,5);}
	if(goreg==true){
		if(document.reg.Email.value.indexOf("@",0) < 1||document.reg.Email.value.indexOf(".",0) < 1||document.reg.Email.value.length-document.reg.Email.value.indexOf("@",0) < 5||document.reg.Email.value.length-document.reg.Email.value.indexOf(".",0) < 3){
			alert("\n请输入正确的电子信箱\n");
			document.reg.Email.focus();
			goreg = false;
		}
		else{
			goreg=true
		}
	}
	if(goreg==true){checkerr("\n请输入密码找回问题\n\n20个字符以下",document.reg.Question,20,1);}
	if(goreg==true){checkerr("\n请输入密码找回答案\n\n20个字符以下",document.reg.Ans,20,1);}

	if(goreg == true){
		acc.value = "正在提交";
		acc.disabled = true;
		reg.submit();}
	}
}

function noaccaction(){
	if(Step==2){
		xieyi.style.display = "block";
		acc.value = "同意协议";
		noacc.value = "我不同意";
		mustinfo.style.display = "none";
		others.style.display = "none";
		Step = 1
	}
	else{
		window.location.href="../";}
}

function checkerr(errmsg,obj,maxlen,minlen){
	if(obj.value==''||obj.value.length < minlen||obj.value.length > maxlen){
		alert(errmsg)
		obj.focus();
		goreg = false;
	}
	else{
		goreg=true;
	}
}

function viewnone(e){
	e.style.display=(e.style.display=="none")?"":"none";
}
</script>
</head>
<body id="center">

<div id="xieyi">
	<div style="text-align: center; background: #DBF2FF; border: #A9D5F4 1px solid; line-height: 25px;"><strong>用户协议</strong></div>
	<div style="border-left: #A9D5F4 1px solid; border-right: #A9D5F4 1px solid; border-bottom: #A9D5F4 1px solid;"><!--#include file="../language/zh-cn_reg.asp" --></div>
</div>

<div id="mustinfo" style="display:none">
	<form action="?action=SaveData" method="post" name="reg" id="reg">
	<table width="100%" style="border: #A9D5F4 1px solid;">
		<tr>
			<td height="25" colspan="3" bgcolor="#DBF2FF" align="center"><strong>基本资料</strong></td>
		</tr>
		<tr>
			<td height="25" align="right">请输入您想要申请的昵称&nbsp;</td>
			<td align="left"><input name="UserName" type="text" id="UserName" class="LoginInput" style="width: 150px;" /></td>
			<td><font color="#FF0000">*</font>[<font color="#999999">最大10个字符.不能使用特殊字符</font>]</td>
		</tr> 
		<tr>
			<td height="25" align="right">请为新帐户设定一个密码&nbsp;</td>
			<td align="left"><input name="Password" type="password" id="Password" class="LoginInput" style="width: 150px;" /></td>
			<td><font color="#FF0000">*</font>[<font color="#999999">大于6小于14个字符.</font><font color="#999999">不能使用特殊字符</font>]</td>
		</tr>
		<tr>
			<td height="25" align="right">请再输入一遍设定的密码&nbsp;</td>
			<td align="left"><input name="RePassWord" type="password" id="RePassWord" class="LoginInput" style="width: 150px;" /></td>
			<td><font color="#FF0000">*</font>[<font color="#999999">确认一遍您输入的密码</font>]</td>
		</tr>
		<tr>
			<td height="25" align="right">请输入您常用的电子信箱&nbsp;</td>
			<td align="left"><input name="Email" type="text" id="Email" class="LoginInput" style="width: 150px;" /></td>
			<td><font color="#FF0000">*</font>[<font color="#999999">必须是有效的电子信箱,当你忘记密码时要用到它</font>]</td>
		</tr>
		<tr>
			<td height="25" align="right">请输入你的密码找回问题&nbsp;</td>
			<td align="left"><input name="Question" type="text" id="Question" class="LoginInput" style="width: 150px;" /></td>
			<td><font color="#FF0000">*</font>[<font color="#999999">不能使用空格等特殊字符</font>]</td>
		</tr>
		<tr>
			<td height="25" align="right">请输入你的密码找回答案&nbsp;</td>
			<td align="left"><input name="Ans" type="text" id="Ans" class="LoginInput" style="width: 150px;" /></td>
			<td><font color="#FF0000">*</font>[<font color="#999999">不能使用空格等特殊字符</font>]</td>
		</tr>
		<tr>
			<td height="25" align="right" name=up>您的性别&nbsp;</td>
			<td align="left">男<input name="sex" type="radio" value="1" />女<input name="sex" type="radio" value="0" /></td>
			<td>&nbsp;</td>
		</tr>
		<tr>
			<td height="25" align="right">开启高级选项&nbsp;</td>
			<td align="left"><input name="otherselect" type="checkbox" value="1" onclick="viewnone(others)" /></td>
			<td>&nbsp;</td>
		</tr>
	</table>
</div>
<p />
<div id="others" style="display:none">
	<table width="100%" style="border: #A9D5F4 1px solid;">
		<tr style="color:#000000">
			<td height="25" colspan="2" bgcolor="#DBF2FF" align="center"><strong>详细资料</strong></td>
		</tr>
		<tr>
			<td height="25" align="right">个人网站地址&nbsp;</td>
			<td><input name="HomePage" type="text" id="HomePage" class="LoginInput" style="width: 150px;" /></td>
		</tr>
		<tr>
			<td height="25" align="right">腾讯QQ号码&nbsp;</td>
			<td><input name="QQ" type="text" id="QQ" class="LoginInput" style="width: 150px;" /></td>
		</tr>
		<tr>
			<td height="25" align="right">ICQ号码&nbsp;</td>
			<td><input name="ICQ" type="text" id="ICQ" class="LoginInput" style="width: 150px;" /></td>
		</tr>
		<tr>
			<td height="25" align="right">MSN帐户&nbsp;</td>
			<td><input name="MSN" type="text" id="MSN" class="LoginInput" style="width: 150px;" /></td>
		</tr>
		<tr>
			<td height="25" align="right">真实姓名&nbsp;</td>
			<td><input name="ReName" type="text" id="ReName" class="LoginInput" style="width: 150px;" /></td>
		</tr>
		<tr>
			<td height="25" align="right">出生日期&nbsp;</td>
			<td><input type="text" name="date" id="date" size="10" readonly class="LoginInput" onclick="SD(this,'date');"></td>
		</tr>
		<tr>
			<td height="25" align="right">来自&nbsp;</td>
			<td><input name="UForm" type="text" id="UForm" class="LoginInput" style="width: 150px;" /></td>
		</tr>
	</table>
	</form>
	<p />
</div> 

<div style="text-align: center;">
	<input type="button" value="同意协议" name="acc" id="acc" onClick="accaction()" />&nbsp;<input type="button" value="我不同意" name="noacc" id="noacc" onClick="noaccaction()" />
</div>
