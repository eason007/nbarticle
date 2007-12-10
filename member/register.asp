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
'= 文件名称：Member/Register.asp
'= 摘    要：会员-注册文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2006-08-20
'====================================================================

If EA_Pub.SysInfo(7)="0" Then
	ErrMsg="系统已暂停新会员注册"
	Call EA_Pub.ShowErrMsg(33,1)
End If

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
		ErrMsg="您填写的E-Mail地址已被注册，请重新选择其他E-Mail地址！"
		Call EA_Pub.ShowErrMsg(0,2)
	Case 0
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
		Call EA_Pub.ShowSusMsg(1,0)
	Case 1
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
		Call EA_Pub.ShowSusMsg(9,0)
	Case 2
		ErrMsg="您填写的用户名已被注册，请重新选择其他用户名！"
		Call EA_Pub.ShowErrMsg(0,2)
	End Select	
End if
%>
<html>
<head>
<title>注册会员</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta http-equiv="Content-Language" content="zh-CN">
<link href="style.css" rel="stylesheet" type="text/css" />
</head>
<body bgcolor="#FFFFFF" text="#000000"> 
<script language="JavaScript" type="text/javascript">
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
<script src="../js/jsdate.js"></script> 
<table width="762" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#dddddd"> 
  <tr> 
    <td bgcolor="#FFFFFF"><table width="760" border="0" align="center" cellpadding="0" cellspacing="0"> 
        <tr> 
          <td align="center" valign="top"><div id="xieyi"> 
              <table border=0 cellpadding=3 cellspacing=2 width="760" align=center> 
                <tr style="color:#000000"> 
                  <td bgcolor="#e6f0ff" height="30">&nbsp;<b>用户协议</b></td>
                </tr> 
              </table> 
              <table width="760" align=center border="0" cellspacing="3" cellpadding="3"> 
                <tr> 
                  <td bgcolor="#FFFFFF"><!--#include file="../language_files/registration.asp" --></td> 
                </tr> 
              </table> 
            </div> 
            <form action="?action=SaveData" method="post" name="reg" id="reg"> 
              <div id="mustinfo" style="display:none"> 
                <table border=0 cellpadding=3 cellspacing=2 width="760" align=center> 
                  <tr style="color:#000000"> 
                    <td bgcolor="#e6f0ff" height="30">&nbsp;<b>基本资料</b></td>
                  </tr> 
                </table> 
                <table width="760" align=center cellpadding="3" cellspacing="3" bgcolor="#FFFFFF"> 
                  <tbody> 
                    <tr> 
                      <td width="250" height="25" align="right" bgcolor="#efefef">请输入您想要申请的昵称&nbsp;</td> 
                      <td align="left"><input name="UserName" type="text" id="UserName" class="LoginInput" maxlength="20"> </td> 
                      <td><font color="#FF0000">*</font>[<font color="#999999">最大10个字符.不能使用特殊字符</font>]</td> 
                    </tr> 
                    <tr> 
                      <td height="25" align="right" bgcolor="#efefef">请为新帐户设定一个密码&nbsp;</td> 
                      <td align="left"><input name="Password" type="password" id="Password" class="LoginInput" maxlength="20"></td> 
                      <td><font color="#FF0000">*</font>[<font color="#999999">大于6小于14个字符.</font><font color="#999999">不能使用特殊字符</font>]</td> 
                    </tr> 
                    <tr> 
                      <td height="25" align="right" bgcolor="#efefef">请再输入一遍设定的密码&nbsp;</td> 
                      <td align="left"><input name="RePassWord" type="password" id="RePassWord" class="LoginInput" maxlength="20"></td> 
                      <td><font color="#FF0000">*</font>[<font color="#999999">确认一遍您输入的密码</font>]</td> 
                    </tr> 
                    <tr> 
                      <td height="25" align="right" bgcolor="#efefef">请输入您常用的电子信箱&nbsp;</td> 
                      <td align="left"><input name="Email" type="text" id="Email" class="LoginInput" maxlength="30"> </td> 
                      <td><font color="#FF0000">*</font>[<font color="#999999">必须是有效的电子信箱,当你忘记密码时要用到它</font>]</td> 
                    </tr> 
                    <tr> 
                      <td height="25" align="right" bgcolor="#efefef">请输入你的密码找回问题&nbsp;</td> 
                      <td align="left"><input name="Question" type="text" id="Question" class="LoginInput" maxlength="20"></td> 
                      <td><font color="#FF0000">*</font>[<font color="#999999">不能使用空格等特殊字符</font>]</td> 
                    </tr> 
                    <tr> 
                      <td height="25" align="right" bgcolor="#efefef">请输入你的密码找回答案&nbsp;</td> 
                      <td align="left"><input name="Ans" type="text" id="Ans" class="LoginInput" maxlength="20"></td> 
                      <td><font color="#FF0000">*</font>[<font color="#999999">不能使用空格等特殊字符</font>]</td> 
                    </tr> 
                    <tr> 
                      <td height="25" align="right" name=up bgcolor="#efefef">您的性别&nbsp;</td> 
                      <td align="left">男
                        <input name="sex" type="radio" value="1" checked> 
                        女
                        <input name="sex" type="radio" value="0"></td> 
                      <td>&nbsp;</td> 
                    </tr> 
                    <tr> 
                      <td height="25" align="right" bgcolor="#efefef">开启高级选项&nbsp;</td> 
                      <td align="left"><input name="otherselect" type="checkbox" value="1" onclick="viewnone(others)"> </td> 
                      <td>&nbsp;</td> 
                    </tr> 
                  </tbody> 
                </table>
              </div> 
              <div id="others" style="display:none"> 
                <table border=0 cellpadding=3 cellspacing=2 width="760" align=center> 
                  <tr style="color:#000000"> 
                    <td bgcolor="#e6f0ff" height="30">&nbsp;<b>详细资料</b></td>
                  </tr> 
                </table> 
                <table width="760" align=center border="0" cellpadding="3" cellspacing="3" bgcolor="#FFFFFF"> 
                  <tbody> 
                    <tr> 
                      <td width="250" height="24" align="right" bgcolor="#efefef">个人网站地址&nbsp;</td> 
                      <td colspan="2"><input name="HomePage" type="text" id="HomePage" class="LoginInput" size="35" maxlength="50"></td> 
                    </tr> 
                    <tr> 
                      <td height="25" align="right" bgcolor="#efefef">腾讯QQ号码&nbsp;</td> 
                      <td colspan="2"><input name="QQ" type="text" id="QQ" class="LoginInput" size="15" maxlength="20"></td> 
                    </tr> 
                    <tr> 
                      <td height="25" align="right" bgcolor="#efefef">ICQ号码&nbsp;</td> 
                      <td colspan="2"><input name="ICQ" type="text" id="ICQ" class="LoginInput" size="15" maxlength="20"></td> 
                    </tr> 
                    <tr> 
                      <td height="25" align="right" bgcolor="#efefef">MSN帐户&nbsp;</td> 
                      <td colspan="2"><input name="MSN" type="text" id="MSN" class="LoginInput" size="20" maxlength="40"></td> 
                    </tr> 
                    <tr> 
                      <td height="25" align="right" bgcolor="#efefef">真实姓名&nbsp;</td> 
                      <td colspan="2"><input name="ReName" type="text" id="ReName" class="LoginInput" size="15" maxlength="20"> </td> 
                    </tr> 
                    <tr> 
                      <td height="25" align="right" bgcolor="#efefef">出生日期&nbsp;</td> 
                      <td colspan="2"><input type="text" name="date" id="date" size="10" readonly class="LoginInput">&nbsp;<a href="javascript:vod()" onclick="SD(this,'date');"><img border="0" src="../images/public/date_picker.gif" width="30" height="19"></a></td> 
                    </tr> 
                    <tr> 
                      <td height="25" align="right" bgcolor="#efefef">来自&nbsp;</td> 
                      <td colspan="2"><input name="UForm" type="text" id="UForm" class="LoginInput" size="50" maxlength="30"></td> 
                    </tr> 
                  </tbody> 
                </table> 
              </div> 
            </form>
            <div bgcolor="efefef">
            <input type="button" value="同意协议" name="acc" id="acc" onClick="accaction()">&nbsp;
            <input type="button" value="我不同意" name="noacc" id="noacc" onClick="noaccaction()">
            <p></p>
            </div>
            </td> 
        </tr> 
      </table></td> 
  </tr> 
</table> 
