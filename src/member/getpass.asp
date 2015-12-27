<!--#Include File="../include/inc.asp"-->
<!--#Include File="../include/md5.asp"-->
<!--#Include File="cls_db.asp"-->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：Member/GetPass.asp
'= 摘    要：会员-取回密码文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2005-06-03
'====================================================================

Dim Action
Action=Request.QueryString ("action")

Select Case LCase(Action)
Case "savepwd"
	Call SavePwd()
Case "change"
	Call Change()
Case "view"
	Call ViewQuestion()
Case Else
	Call Main()
End Select

'===========Step 1=========
Sub Main()
	Call Top
	Call NormalJS("document.reg.UName","请输入您的用户名") 
%>
<table width="99%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr>
    <td bgcolor="#FFFFFF">
      <table border=0 cellpadding=3 cellspacing=2 width="100%" align=center>
        <tr style="color:#000000"> 
          <td bgcolor="#dddddd" height="30">&nbsp;<b>第一步：输入用户名</b></td>
        </tr>
      </table>
      <table width="100%" align=center cellpadding="1" cellspacing="2" bgcolor="#FFFFFF">
        <form action="?action=View" method="post" name="reg" target="_self" onsubmit="return Noac;">
          <tr height="25"> 
            <td width="20%" bgcolor="#e6f0ff">&nbsp;<font color="#FF6600" face="Webdings">4</font>用户名</td>
            <td>&nbsp;<input name="UName" type="text" onkeydown="readkey()" id="UName" size="30" maxlength="40"></td>
          </tr>
          <tr> 
            <td colspan="2" bgcolor="#efefef" height="30"><input type="button" value="提交" onclick="checkkey()" name="acc" id="acc">&nbsp;<input type="reset" value="清除重来" name="noacc" id="noacc"></td>
          </tr>
        </form>
      </table></td>
  </tr>
</table>
<%
End Sub

'==========Step 2==============
Sub ViewQuestion()
	Call EA_Pub.Chk_Post
	
	Dim Question,UserName
	Dim Temp
	
	UserName=EA_Pub.SafeRequest(2,"UName",1,"",0)
	
	Temp=EA_DBO.Get_MemberQuestionByAccountName(UserName)
	If IsArray(Temp) Then 
		Question=Temp(1,0)
	Else
		If Temp=0 Then
			ErrMsg="用户不存在！"
			Call EA_Pub.ShowErrMsg(0,2)
		End If
	End If
	
	Call Top
	Call NormalJS("document.reg.Answer","请输入您的答案")
%>
<table width="99%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr> 
    <td bgcolor="#FFFFFF">
      <table border=0 cellpadding=3 cellspacing=2 width="100%" align=center>
        <tr style="color:#000000"> 
          <td bgcolor="#dddddd" height="30">&nbsp;<b>第二步：回答问题</b></td>
        </tr>
      </table>
      <table width="100%" align=center cellpadding="1" cellspacing="2" bgcolor="#FFFFFF">
        <form action="?action=Change" method="post" name="reg" target="_self" onsubmit="return Noac;">
          <tr> 
            <td width="20%" bgcolor="#e6f0ff">&nbsp;<font color="#FF6600" face="Webdings">4</font>您的密码提示问题</td>
            <td>&nbsp;<font color=800000><b><%=Question%></b></font></td>
          </tr>
          <tr> 
            <td width="20%" bgcolor="#e6f0ff">&nbsp;<font color="#FF6600" face="Webdings">4</font>请输入答案</td>
            <td>&nbsp;<input name="Answer" onkeydown="readkey()" type="text" id="Answer" size="20" maxlength="40"></td>
          </tr>
          <tr> 
            <td colspan="2" height="30" bgcolor="#efefef"><input name="UName" type="hidden" value="<%=UserName%>"><input type="submit" value="提交" onclick="checkkey()" name="acc" id="acc">&nbsp;<input type="reset" value="清除重来" name="noacc" id="noacc"></td>
          </tr>
        </form>
      </table></td>
  </tr>
</table>
<%
End Sub
		
'=======Step 3=======
Sub Change()
	Call EA_Pub.Chk_Post
		
	Dim UserName,Answer
	Dim Temp,TempAnswer
	
	UserName=EA_Pub.SafeRequest(2,"UName",1,"",0)
	Answer=Request.Form ("Answer")
	TempAnswer=MD5(Answer)
	
	Temp=EA_DBO.Get_MemberQuestionByAccountName(UserName)
	If IsArray(Temp) Then 
		If Temp(2,0)<>TempAnswer Then 
			ErrMsg="问题答案错误。"
			Call EA_Pub.ShowErrMsg(0,2)
		End If
	Else
		If Temp=0 Then
			ErrMsg="用户不存在！"
			Call EA_Pub.ShowErrMsg(0,2)
		End If
	End If
	
	Call Top
%>
<script language="JavaScript" type="text/javascript">
var ie=(document.all)?true:false;
var Noac = false;
function readkey(){
	if(ie){	if(window.event.keyCode==13){checkkey();}
	}
}

function checkkey(){
	if(document.reg.Password.value==''||document.reg.Password.value.length<6||document.reg.Password.value.length>14){
		alert("请填写您的新密码\n6-14个字符以内");
		document.reg.Password.select();
	}
	else{
		if(document.reg.RePassword.value!=document.reg.Password.value||document.reg.RePassword.value==''||document.reg.RePassword.value.length<6||document.reg.RePassword.value.length>14){
			alert("请再一次填写新密码\n必须与上面的密码相同");
			document.reg.RePassword.select();
		}
		else{
			Noac = true;
			document.reg.submit();
		}
	}
}
</script>
<table width="99%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr> 
    <td bgcolor="#FFFFFF">
      <table border=0 cellpadding=3 cellspacing=2 width="100%" align=center>
        <tr style="color:#000000"> 
          <td bgcolor="#dddddd" height="30">&nbsp;<b>第三步：修改密码</b></td>
        </tr>
      </table>
      <table width="100%" align=center cellpadding="1" cellspacing="2" bgcolor="#FFFFFF">
        <form action="?action=SavePwd" method="post" name="reg" target="_self" onsubmit="return Noac;">
          <tr> 
            <td width="20%" bgcolor="#e6f0ff">&nbsp;<font color="#FF6600" face="Webdings">4</font>新的密码</td>
            <td>&nbsp;<input name="Password" type="password" onkeydown="readkey()" id="Password" size="15" maxlength="20">&nbsp;[<font color="#999999">6-14个字符.</font><font color="#999999">不能使用特殊字符</font>]</td>
          </tr>
          <tr> 
            <td width="20%" bgcolor="#e6f0ff">&nbsp;<font color="#FF6600" face="Webdings">4</font>确认新的密码</td>
            <td>&nbsp;<input name="RePassword" type="password" onkeydown="readkey()" id="RePassword" size="15" maxlength="20">
              <input name="Answer" type="hidden" id="Answer" value="<%=Answer%>">
              <input name="UName" type="hidden" value="<%=UserName%>"></td>
          </tr>
          <tr>
            <td colspan="2" height="30" bgcolor="#efefef"><input type="button" value="提交修改" onclick="checkkey()" name="acc" id="acc">&nbsp;<input type="reset" value="清除重来" name="noacc" id="noacc"></td>
          </tr>
        </form>
      </table></td>
  </tr>
</table>
<%
End Sub

Sub SavePwd()
	Call EA_Pub.Chk_Post
	
	Dim Question,UserName,Answer,TempAnswer,NewPassword
	Dim Temp
	
	UserName=EA_Pub.SafeRequest(2,"UName",1,"",0)
	Question=EA_Pub.SafeRequest(2,"Question",1,"",0)
	Answer=Request.Form ("Answer")
	TempAnswer=MD5(Answer)
	NewPassword=Request.Form("Password")
	NewPassword=MD5(NewPassword)
	
	Temp=EA_DBO.Get_MemberQuestionByAccountName(UserName)
	If IsArray(Temp) Then 
		If Temp(2,0)<>TempAnswer Then 
			ErrMsg="问题答案错误。"
			Call EA_Pub.ShowErrMsg(0,2)
		Else
			Temp=EA_DBO.Set_MemberPasswordByAccountName(UserName,NewPassword)
			If Temp=0 Then 
				ErrMsg="密码修改成功！"
				Call EA_Pub.ShowErrMsg(0,0)
			End If
		End If
	Else
		If Temp=0 Then
			ErrMsg="用户不存在！"
			Call EA_Pub.ShowErrMsg(0,2)
		End If
	End If
End Sub

Sub Top
%>
<html>
<head>
<title>取回密码</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta http-equiv="Content-Language" content="zh-CN">
<link href="style.css" rel="stylesheet" type="text/css" />
</head>
<body bgcolor="#FFFFFF" text="#000000"> 
<%
End Sub

Sub NormalJS(Control,Msg)
%>
<script language="JavaScript" type="text/javascript">
var ie=(document.all)?true:false;
var Noac = false;
function readkey(){
	if(ie){	if(window.event.keyCode==13){checkkey();}
	}
}

function checkkey(){
	if(<%=Control%>.value==''){
		alert("<%=Msg%>");
		<%=Control%>.select();
	}
	else{
		Noac = true;
		document.reg.submit();
	}
}
</script>
<%End Sub%>