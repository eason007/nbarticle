<!--#Include File="../conn.asp" -->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2004 - 2008 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：Install.asp
'= 摘    要：安装文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-02-14
'====================================================================

Dim dis
Dim Gocode
Dim StepNum
Dim Url

If Request("action") = "testDBConnection" Then
	Call testDBConnection()
	Response.End
End If

If Request("Step") = "" Or Not IsNumeric(Request("Step")) Then
	StepNum = 1
Else
	StepNum = CInt(Request("Step"))
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="zh-CN">
<head>
<title>欢迎您使用NB文章系统(NBArticle)安装向导 - 免费、开源、高效、安全的ASP内容管理系统</title>
<meta http-equiv="content-type" content="text/html; charset=utf-8" />
<meta http-equiv="content-language" content="zh-CN" />
<meta name="generator" content="NB文章系统(NBArticle) EliteArticle System" />
<meta name="author" content="eason007<eason007#163.com>" />
<meta name="copyright" content="www.nbarticle.com" />
<script src="../js/public.js" type="text/javascript"></script>
<script src="../js/objAjax.js" type="text/javascript"></script>
<script type="text/javascript">
function testDBConnection () {
	$('testBtn').disabled = true;
	$('message').innerHTML = "";

	var iDatabaseType = $('DataType').value;

	var postDal = "action=testDBConnection&DataType=" + iDatabaseType;
	postDal+= "&Key=" + $('Key').value;

	if (iDatabaseType == "0")
	{
		postDal+= "&Folder=" + $('Folder').value;
		postDal+= "&DataName=" + $('Folder').value + $('DataName').value;
	}
	else{
		postDal+= "&DataHost=" + $('DataHost').value;
		postDal+= "&DataName=" + $('DataName').value;
		postDal+= "&DataUser=" + $('DataUser').value;
		postDal+= "&DataPass=" + $('DataPass').value;
	}

	ajaxPostDate ("./install.asp", postDal, reportTest, "TEXT");
}

function reportTest (obj) {
	$('testBtn').disabled = false;

	var message = obj.returnMessage;

	switch (message.substring(0, 1))
	{
	case "0":
		$('message').innerHTML = "配置文件已被创建";
		$('message').style.color = "green";
		$('N').disabled = false;
		break;

	case "1":
		$('message').innerHTML = "您的数据驱动存在问题,或者您设定的条件不足以连接指定的数据服务器";
		$('message').style.color = "red";
		break;

	case "2":
		$('message').innerHTML = "因为您的服务器FSO权限设置可能有点问题,配置文件未被写入,请自行修改系统根目录的[Connection.asp]文件";
		$('message').style.color = "red";
		break;
	}

	$('message').style.display = "";
}
</script>
<style type="text/css">
body {
	padding: 0;
	margin: 5px 10px;
	font-family: "Microsoft JhengHei","微軟正黑體","Microsoft YaHei","微软雅黑","华文细黑", "Trebuchet MS", Verdana, Arial, sans-serif;
	font-size: 12px;
}
#center {
	WIDTH: 990px;
	MARGIN-RIGHT: auto;
	MARGIN-LEFT: auto;
}

a {
	color: #0082FF;
}

.left {
	float: left;
}
.right {
	float: right;
}
.hidden {
	display: none;
}
.separator{
	height: 5px;
	clear: both;
}
img {
	padding: 0;
	margin: 0;
	border: 0;
}

#head {
	border:1px #c5d7e2 solid;
	text-align: center;
	height: 90px;
}

#head #logo {
	padding:20px 15px;
}

#head #banner {
	padding:5px 40px;
}

#nav {
	border:1px #c5d7e2 solid;
	background:#CCEBF7;
	color:#666;
	padding:1px 5px;
}

#content {
	padding: 5px 100px;
}

#content #title {
	border-top:1px #AAA solid;
	border-left:1px #AAA solid;
	border-right:1px #AAA solid;
	text-align: center;
	font-size: 15px;
	font-weight:bold;
	padding: 5px;
	background:#F8F8F8;
}

#content #text {
	border:1px #AAA solid;
	padding: 0 10px;
	clear: both; 
	word-wrap: break-word; 
	word-break: break-all;
	overflow: auto;
}

#content #text .italic {
	color: #0082FF;
	font-style:italic;
}

#content #text strong {
	font-style:italic;
	font-weight:normal;
}

#message {
	text-align: center;
}

#step {
	padding: 5px 100px;
}
</style>
<body id="center">

<div id="head">
	<span class="left" id="logo"><img src="nbarticle.gif" alt="NB文章系统(NBArticle)"></span>
	<span class="left" id="banner"><img src="banner.jpg" alt="NB文章系统(NBArticle) - 免费、开源、高效的ASP内容管理系统"></span>
</div>

<div class="separator"></div>

<div id="nav">欢迎使用NB文章系统(NBArticle)安装向导 - 第<%=StepNum%>步</div>

<div class="separator"></div>

<div id="content">
	<%
	Select Case StepNum
	Case 1
		If iDataBaseType = 0 Then
			Call ShowLicense
		Else
			Call Step1
		End If
		GoCode = "location.href = '?Step="& StepNum + 1 &"'"
	Case 2
		If iDataBaseType = 0 Then
			Call Step1
		Else
			Call SQL_Step2
		End If
		GoCode = "location.href = '?Step="& StepNum + 1 &"'"
	Case 3
		Dim RndNum
		Dim AllPath,Folder,Index
		
		Randomize Timer
		RndNum = "NB" & (1+Int(rnd*1000000000))
		
		Folder=Replace(Request.ServerVariables ("URL"), "install/install.asp", "")

		If iDataBaseType=0 Then 
			Call Access_Step2

			dis = " disabled"
		Else
			StepNum = 4

			Call EndStr
		End If
		
		GoCode = "location.href = '?Step="& StepNum + 1 &"'"
	Case 4
		Call EndStr
	End Select
	%>
</div>

<div id="message" style="display:none;"></div>

<div id="step" class="right">
	<% If StepNum <= 3 Then %>
		<input type="button" name="C" value="取消安装" onclick="javascript: window.close();" />
		<% If StepNum > 1 Then %>
		<input type="button" name="P" value="&lt;&lt;上一步" onclick="javascript: location.href = '?Step=<%= StepNum - 1 %>';" />
		<% End If %>
		<input type="button" name="N" id="N" value="下一步&gt;&gt;"<%=dis%> onclick="javascript: <%=GoCode%>;" />
	<% End If %>
</div>

</body>
</html>

<%Sub ShowLicense%>
<div id="title">最终用户授权协议</div>

<div id="text">
	<p>版权所有&nbsp;(&copy)&nbsp;2004&nbsp;-&nbsp;2008&nbsp;[Team Elite,Eason Chan]</p>
	<p>感谢您选择&nbsp;<a href="http://www.nbarticle.com/" target="_blank">NB文章系统(NBArticle)</a>。本文章系统是由&nbsp;<span class="italic">Team Elite</span>&nbsp;自主开发的一个免费、开源、高效、安全的ASP内容管理系统。</p>
	<p><a href="http://www.nbarticle.com/" target="_blank">NB文章系统(NBArticle)</a>&nbsp;英文全称为&nbsp;NetBuilder Article，中文全称为&nbsp;NB文章系统，以下统称&nbsp;<strong>NB文章系统(NBArticle)</strong>。</p>
	<p><strong>NB文章系统(NBArticle)</strong>&nbsp;官方技术支持论坛为&nbsp;<a href="http://forum.nbarticle.com/" target="_blank">http://forum.nbarticle.com/</a>；<strong>NB文章系统(NBArticle)</strong>&nbsp;官方产品网站为&nbsp;<a href="http://www.nbarticle.com/" target="_blank">http://www.nbarticle.com/</a>。</p>
	<p>在开始安装&nbsp;<strong>NB文章系统(NBArticle)</strong>&nbsp;之前，请务必仔细阅读本授权文档，在您确定符合授权协议的全部条件后，即可继续&nbsp;<strong>NB文章系统(NBArticle)</strong>&nbsp;的安装。即：您一旦开始安装&nbsp;<strong>NB文章系统(NBArticle)</strong>&nbsp;，即被视为完全同意本授权协议的全部内容，如果出现纠纷，我们将根据相关法律和协议条款追究责任。</p>
	<p><strong>NB文章系统(NBArticle)</strong>&nbsp;著作权受到法律及国际公约保护。本次安装版本为免费版，用户可以在完全遵守本最终用户授权协议的基础上，免费获得、安装及在非商业用途使用本程序，而不必支付费用。</p>
	<p>您可以查看&nbsp;<strong>NB文章系统(NBArticle)</strong>&nbsp;的全部源代码，也可以根据自己的需要对其进行修改，但无论如何，即无论用途如何、是否改动、改动程度如何，只要&nbsp;<strong>NB文章系统(NBArticle)</strong>&nbsp;程序的任何部分被包含在您修改后的系统中，都必须保留页脚处的&nbsp;<strong>NB文章系统(NBArticle)</strong>&nbsp;名称和&nbsp;<a href="http://www.nbarticle.com/" target="_blank">http://www.nbarticle.com/</a>&nbsp;的链接。您修改后的代码，在没有获得我们书面许可的情况下，严禁公开发布或发售。</p>
	<p>用户出于自愿而使用本软件，我们不承诺提供任何形式的技术支持、使用担保，也不承担任何因使用本软件而产生问题的相关责任。</p>
	<p>对于且仅对于非营利性个人用户，本版本&nbsp;<strong>NB文章系统(NBArticle)</strong>&nbsp;是开放源代码的免费软件，欢迎您在原样完整保留全部版权信息和说明文档的前提下，传播和转载本程序。在未购买商业授权前，严禁将本软件用于商业用途或盈利性站点，即：企业、营利性社会团体、从事经营性互联网业务的公司与个人、对最终用户收取费用的收费网站，均需要购买商业版本授权才能使用本软件，有关商业版本购买事宜请访问官方产品网站（<a href="http://www.nbarticle.com/" target="_blank">http://www.nbarticle.com/</a>）。同时欢迎对&nbsp;<strong>NB文章系统(NBArticle)</strong>&nbsp;感兴趣并有实力的团体或个人在开发过程中提供支持。</p>
	<p>安装&nbsp;<strong>NB文章系统(NBArticle)</strong>&nbsp;建立在完全同意本授权协议的基础之上，因此而产生的纠纷，违反本协议的一方将承担全部民事与刑事责任。</p>
</div>
<%
End Sub

Sub Step1
%>
<div id="title">检测服务器环境</div>

<div id="text">
	<input type="hidden" name="Key" id="Key" value="<%=RndNum%>" />
	<input type="hidden" name="DataType" id="DataType" value="<%=iDataBaseType%>" />
	<input type="hidden" name="Folder" id="Folder" value="<%=Folder%>" />

	<p><strong>系统运行需要服务器有以下的支持：</strong>
	<br />>>IIS版本：<span style="color:#800000;">5.1或以上</span>
	<br />>>脚本解译引擎：<span style="color:#800000;">VBScript 5.6.8820或以上</span>
	<br />>>Scripting.FileSystemObject(FSO)：<span style="color:#800000;">启用状态</span>
	<br />>>ADODB.Stream：<span style="color:#800000;">启用状态</span>
	<br />>>ADODB.Connection(ADO)：<span style="color:#800000;">启用状态</span>
	<br />>>CDONTS或Jmail：<span style="color:#800000;">启用状态</span></p>
	<p><strong>当前状态：</strong>
	<br />>IIS版本：<span style="color:#800000;"><%=Request.ServerVariables("SERVER_SOFTWARE")%></span>
	<br />>脚本解译引擎：<span style="color:#800000;"><%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %></span>
	<br />>Scripting.FileSystemObject(FSO)：<span style="color:#800000;"><%=ObjTest("Scripting.FileSystemObject")%></span>
	<br />>ADODB.Stream：<span style="color:#800000;"><%=ObjTest("ADODB.Stream")%></span>
	<br />>ADODB.Connection(ADO)：<span style="color:#800000;"><%=ObjTest("ADODB.Connection")%></span>
	<br />>CDONTS或Jmail：<span style="color:#800000;"><%=ObjTest("CDONTS.NewMail")%>/<%=ObjTest("JMail.SmtpMail ")%></span></p>
</div>
<%
End Sub

Sub Access_Step2
%>
<div id="title">连接数据库</div>

<div id="text">
	<input type="hidden" name="Key" id="Key" value="<%=RndNum%>" />
	<input type="hidden" name="DataType" id="DataType" value="<%=iDataBaseType%>" />
	<input type="hidden" name="Folder" id="Folder" value="<%=Folder%>" />

	<p>>>请输入你的Access数据库路径：
	<br /><%=Folder%><input name="DataName" id="DataName" type="text" size="30" value="db/NBArticle.asp" />&nbsp;<input type="button" id="testBtn" onclick="javascript:testDBConnection();" value="测试连接" />
	<br />>><span style="color:#800000;">如没有修改数据库目录名及文件名，使用默认的即可；否则请改为正确的目录名及文件名。数据库文件已进行防下载处理</span></p>
</div>
<%
End Sub

Sub SQL_Step2
%>
<div id="title">连接数据库</div>

<div id="text">
	<input type="hidden" name="Key" id="Key" value="<%=RndNum%>" />
	<input type="hidden" name="DataType" id="DataType" value="<%=iDataBaseType%>" />

	<p>>>请输入你的SQL Server 2000数据库信息：
	<br />数据服务器：<input name="DataHost" id="DataHost" type="text" size="30" value="(local)" />
	<br />数据库名称：<input name="DataName" id="DataName" type="text" />
	<br />数据库用户：<input name="DataUser" id="DataUser" type="text" />
	<br />数据库密码：<input name="DataPass" id="DataPass" type="text" />

	<br /><input type="button" id="testBtn" onclick="javascript:testDBConnection();" value="测试连接"></p>
</div>
<%
End Sub

Sub EndStr
	Dim HostUrl
	
	HostUrl = Request.ServerVariables ("SERVER_NAME")
	If Request.ServerVariables ("SERVER_PORT") <> "80" Then
		HostUrl = HostUrl & ":" & Request.ServerVariables ("SERVER_PORT")
	End If
	HostUrl = HostUrl & Replace(Request.ServerVariables ("URL"), "/install/install.asp", "") & "/"
%>
<div id="title">安装完成</div>

<div id="text">
	<p>恭喜！<br />安装顺利完成，现在您可以通过&nbsp;<a href="../">http://<%=HostUrl%></a>&nbsp;来访问NB文章系统(NBArticle)了。</p>
</div>
<%
End Sub

Sub testDBConnection()
	On Error Resume Next
	Err.Clear
	Set Conn = Server.CreateObject("ADODB.Connection")

	If Request("DataType")="0" Then 
		ConnStr="Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(Request("DataName"))
		Conn.Open ConnStr
	Else
		ConnStr="Provider=Sqloledb;server="& Request("DataHost") &";uid="& Request("DataUser") &";pwd="& Request("DataPass") & ";database="& Request("DataName")
		Conn.Open ConnStr
	End If
	
	If Err.Number <> 0 Then
		Response.Write "1"
		Response.End
	End If

	Dim Content
	If Request("DataType")<>"0" Then 
		Content = "<" & CHR(37) & VBCrlf
		Content = Content & "Dim ConnStr" & VBCrlf
		Content = Content & "ConnStr = ""Provider=Sqloledb;server="& Request("DataHost") &";uid="& Request("DataUser") &";pwd="& Request("DataPass") &";database="& Request("DataName") &"""" & VBCrlf
		Content = Content & "Const sCacheName=""" & Request("key") & """" & VBCrlf
		Content = Content & "Const SystemFolder=""" & Request("folder") & """" & VBCrlf
		Content = Content & CHR(37) & ">" & VBCrlf
	Else
		Content = "<" & CHR(37) & VBCrlf
		Content = Content & "Dim ConnStr" & VBCrlf
		Content = Content & "Dim DataBaseFilePath" & VBCrlf
		Content = Content & "DataBaseFilePath=""" &Request("DataName") & """" & VBCrlf
		Content = Content & "ConnStr=""Provider = Microsoft.Jet.OLEDB.4.0;Data Source ="""
		Content = Content & " & Server.MapPath(DataBaseFilePath)"& VBCrlf
		Content = Content & "Const sCacheName=""" & Request("key") & """" & VBCrlf
		Content = Content & "Const SystemFolder=""" & Request("folder") & """" & VBCrlf
		Content = Content & CHR(37) & ">" & VBCrlf
	End If
	
	Response.Write WriteFile(Content, Server.MapPath ("../Connection.asp"))
	Response.Flush

	Conn.Close
	Set Conn = Nothing
End Sub

Function WriteFile(Content,Path)
	On Error Resume Next
	Dim MyFile

	Err.Clear
	Set MyFile = Server.CreateObject("ADOD" & "B.S" & "TREAM")
	With MyFile
		.Open
		.Charset = "utf-8"
		.WriteText Content
		.SaveToFile Path,2
		.Close
	End With
	If Err.Number <> 0 Then 
		WriteFile = "2"
	Else
		WriteFile = "0"
	End If
End Function

'检查组件是否被支持及组件版本的子程序
Function ObjTest(strObj)
	On Error Resume Next
	
	IsObj=False
	VerObj=""

	Dim TestObj
	Set TestObj = Server.CreateObject (strObj)

	If -2147221005 <> Err Then
		IsObj = True
		
		ObjTest = "<span style='color:green;'>√</span>"
	Else
		ObjTest = "<span style='color:red;'>×</span>"
	End If

	ObjTest = ObjTest & TestObj.version
End Function
%>
