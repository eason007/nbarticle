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
'= 文件名称：Eail.asp
'= 摘    要：发送邮件文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2005-11-30
'====================================================================

Dim Action
Action=Request.QueryString ("action")

Select Case LCase(Action)
Case "send"
	Call Send_Mail
Case Else
	Call Main
End Select
Call EA_Pub.Close_Obj
Set EA_Pub=Nothing

Sub Main
%>
<html>
<head>
<title>推荐文章给好友 - <%=EA_Pub.SysInfo(0)%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta http-equiv="Content-Language" content="zh-CN">
<style type="text/css">
body
{
	font-family: "宋体";
	font-size: 12px;
}
TD 
{
	FONT-SIZE:12px; 
	FONT-FAMILY: "宋体";
}
TR 
{
	FONT-SIZE:12px;
	FONT-FAMILY: "宋体";
}
</style>
</head>
<body bgcolor="#FFFFFF" text="#000000" background="Images/kabg.gif">
<table border="0" cellpadding="5" cellspacing="0" align="center" style="border: 1 solid #000000">
  <form name="form1" method="post" action="?action=send">
    <tr> 
      <td align="center" valign="middle" bgcolor="#F0FAFF">：：：将以下链接推荐给好友：：：</td>
    </tr>
    <tr>
      <td height="1" align="center" valign="middle" bgcolor="#000000"> </td>
    </tr>
    <tr> 
      <td valign="middle" align="left" bgcolor="#FFFFFF">&nbsp;您朋友的名字： 
        <input type="text" name="name" class="LoginInput"size="10">
        E-MAIL地址： 
        <input type="text" name="email" class="LoginInput" size="20"> </td>
    </tr>
    <tr> 
      <td valign="middle" bgcolor="#FFFFFF">&nbsp;您的名字： &nbsp;&nbsp;&nbsp; <input type="text" name="name2" class="LoginInput" size="10">
        E-MAIL地址： 
        <input type="text" name="email2" class="LoginInput" size="20"> </td>
    </tr>
    <tr> 
      <td bgcolor="#FFFFFF">&nbsp;正文： 
        <textarea name="text" wrap="VIRTUAL" cols="50" rows="5"></textarea> 
      </td>
    </tr>
    <tr> 
      <td valign="middle" align="center" bgcolor="#FFFFFF">&nbsp; <input type="submit" name="Submit" value="发 送"> 
        <input type="hidden" name="ArticleId" value="<%=Request.QueryString("ArticleId")%>"> 
      </td>
    </tr>
  </form>
</table>
</body>
</html>
<%
End Sub

Sub Send_Mail
	Dim MailBody,Url
	Dim SendFlag
	Dim ArticleId,ArticleTitle,ArticleTime,ArticleInfo

	ArticleId=EA_Pub.SafeRequest(2,"articleid",0,0,0)
	
	ArticleInfo=EA_DBO.Get_Article_Info(ArticleId,0)
	If IsArray(ArticleInfo) Then
		ArticleTitle=ArticleInfo(3,0)
		ArticleTime=ArticleInfo(13,0)
	Else
		ErrMsg="提交参数错误"
		Call EA_Pub.ShowErrMsg(0,0)	
	End If	
	
	If Right(EA_Pub.SysInfo(11),1)<>"/" Then EA_Pub.SysInfo(11)=SysInfo(11)&"/"
	Url=EA_Pub.Cov_ArticlePath(ArticleId,ArticleTime,EA_Pub.SysInfo(18))
	Url=Replace(Url,SystemFolder,"")
	Url=EA_Pub.SysInfo(11)&Url
	
	MailBody="<html>"
	MailBody=MailBody & "<title>好友推荐</title>"
	MailBody=MailBody & "<body>"
	MailBody=MailBody & "<TABLE border=0 width='95%' align=center><TBODY><TR>"
	MailBody=MailBody & "<TD valign=middle align=top>"
	MailBody=MailBody & Request.Form ("name")&"，您好：<br><br>"
	MailBody=MailBody & "以下是你的好友"&Request.Form ("name2")&"特地向您推荐的文章：<br>"
	MailBody=MailBody & "文章标题："&ArticleTitle&"<br>"
	MailBody=MailBody & "它的连接是：<br>"
	MailBody=MailBody & "<a href='"&Url&"' target=_blank>" & Url &"</a><br>"	
	MailBody=MailBody & "并附言：<br>"
	MailBody=MailBody & Request.Form ("text")
	MailBody=MailBody & "<br><br>"
	MailBody=MailBody & "<center><font color=red>感谢您的支持，我们将提供给您最好的服务！</font>"
	MailBody=MailBody & "</TD></TR></TBODY></TABLE><br><hr width=95% size=1>"
	MailBody=MailBody & "<p align=center>Copyright 2005 <a href="&EA_Pub.SysInfo(11)&" target=_blank>"&EA_Pub.SysInfo(0)&"</a>. All Rights Reserved.</p>"
	MailBody=MailBody & "</body>"
	MailBody=MailBody & "</html>"
	
	If IsObjInstalled("JMail.SMTPMail") Then
		Call jmail(Request.Form ("email"),Request.Form ("title"),MailBody)
	Else
		Call Cdonts(Request.Form ("email"),Request.Form ("title"),MailBody)
	End If
	
	If SendFlag Then 
		ErrMsg="推荐信发送成功！......"
	Else
		ErrMsg="推荐信发送失败"
	End If 
	ErrMsg=ErrMsg&"<br><li>此服务由<A href="&EA_Pub.SysInfo(11)&" target=_blank>"&EA_Pub.SysInfo(0)&"</A>提供，感谢您的使用和支持! "
	'ErrMsg=MailBody
	Call EA_Pub.ShowErrMsg(0,0)	
End Sub

Sub Jmail(email,topic,MailBody)
	on error resume next

	Dim vTemp,MailPassword

	vTemp=EA_DBO.Get_System_Info()
	If IsArray(vTemp) Then 
		vTemp = Split(vTemp(5,0),",")
		
		MailPassword = vTemp(14)

		dim JMail
		Set JMail=Server.CreateObject("JMail.Message")
		JMail.silent=true
		JMail.Logging=True
		JMail.Charset="gb2312"
		JMail.MailServerUserName = EA_Pub.SysInfo(13)		'您的邮件服务器登录名
		JMail.MailServerPassword = MailPassword				'登录密码
		JMail.ContentType = "text/html"
		JMail.Priority = 1
		JMail.From= EA_Pub.SysInfo(12)
		JMail.FromName = EA_Pub.SysInfo(0)
		JMail.AddRecipient email
		JMail.Subject=topic
		JMail.Body=MailBody
		JMail.Send (EA_Pub.SysInfo(15))
		Set JMail=nothing
		SendFlag=True
		If err then SendFlag=False 
	Else
		SendFlag=False 
	End If
End Sub

Sub Cdonts(email,topic,MailBody)
	on error resume next
	dim objCDOMail
	Set objCDOMail = Server.CreateObject("CDONTS.NewMail")
	objCDOMail.From = EA_Pub.SysInfo(12)
	objCDOMail.To =email
	objCDOMail.Subject =topic
	objCDOMail.BodyFormat = 0 
	objCDOMail.MailFormat = 0 
	objCDOMail.Body =MailBody
	objCDOMail.Send
	Set objCDOMail = Nothing
	SendFlag=True
	If err then SendFlag=False 
End Sub
%>