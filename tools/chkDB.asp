<%@ LANGUAGE = VBScript CodePage = 65001%>
<%
'====================================================================
'= Team Elite - EliteCMS
'= Copyright (c) 2004 - 2008 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 文件名称：Coon.asp
'= 摘    要：数据库连接文件
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-03-26
'====================================================================
Option Explicit

Response.Charset= "UTF-8"
Response.Buffer	= True

Dim Action
Dim Conn
Dim TableList
Action = LCase(Request.QueryString("action"))

Select Case Action
Case "step2"
	Call Step2()
Case Else
	Call Step1()
End Select

Sub Top ()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="zh-cn">
<title>数据库维护工具 - EliteCMS Ver4.00 Beta1</title>
<meta name="generator" content="EliteCMS Ver 4.00 Beta1" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="Content-Language" content="zh-cn" />
<meta name="robots" content="nofollow" />
<style type="text/css">
body
{
	font-family: 'Lucida Grande','Lucida Sans Unicode','宋体','新宋体',arial,verdana,sans-serif;
	font-size: 12px;

}
#center {
	width: 750px;
	margin-right: auto;
	margin-left: auto;
}
#foot {
	text-align: center;
}

.border {
	border: #A9D5F4 1px solid;
}
.title {
	background: #DBF2FF;
	line-height: 30px;
	border-bottom: #A9D5F4 1px solid;
}

a {
	color: #0365BF;
}
p {
	padding: 0 10px;
	line-height: 180%;
}
</style>
</head>
<body id="center">
<%
End Sub

Sub Foot ()
%>
<div id="foot">Copyright &copy; 2004 - 2008 <a href="http://www.nbarticle.com" target="_blank">Www.NbArticle.Com</a></div>

</body>
</html>
<%
End Sub

Sub Step1 ()
	Call Top()
%>
<div class="border">
	<div class="title">&nbsp;<strong>数据库维护工具 - EliteCMS Ver4.00 Beta1</strong></div>
	<p>
		版本: Ver 1.00<br />
		时间: 2008-3-26<br />
		作者: eason007<br />
		网址: <a href="http://www.nbarticle.com" target="_blank">Www.NbArticle.Com</a>
	</p>
	<p>
		<strong>操作说明，请仔细阅读操作说明后进行操作</strong>：<br />
		1、强烈建议您在本地计算机系统完成升级操作并做好<font color="red">备份</font>工作，如不能在本地进行升级操作也请升级前<font color="red">备份</font>好您的所有文件和数据。<br />
		2、<strong>本地升级</strong><br />
		　　A.系统要求：使用W2k(Pro&Server)+IIS5.0或者Win2003+IIS6，请不要在Win9x下进行操作。<br />
		　　B.在WEB目录中新建一目录，将本文件放入，同时将网站数据库下载回本地也放入相同目录。<br />
		　　C.在浏览器中执行http://localhost/新建目录名/chkDb.asp文件，然后按照步骤和说明进行操作。<br />
		　　D.如果出现<font color="blue">错误信息</font>，请仔细阅读，这时或许需要您手动进行调整；如果是出现异常的没有解决的问题，请到官方讨论区发言。<br />
		　　E.完成升级，请把升级好的数据库文件传到网站数据库目录即可。在做这步操作之前，<font color="blue">请对您的原版本数据库和文件做好备份。</font>
	</p>
	<hr></hr>
	<p>
		<form method="post" action="?action=step2">
		<strong>第一步</strong><br />
		数据库文件地址：<input type="text" name="dbpath" value="NBArticle.asp" />&nbsp;[相对路径，如data目录的Depot/NBArticle.asp]<br />
		数据库结构模版：<input type="text" name="xmlpath" value="EliteCMS-400.xml" />&nbsp;[相对路径]<br />
		<input name="submit" type="submit" value="开始升级" />
		
		</form>
	</p>
</div>
<p />
<%
	Call Foot()
End Sub

Sub Step2 ()
	If Len(Request.Form("dbpath")) = 0 Or Len(Request.Form("xmlpath")) = 0 Then Response.Redirect "?":Exit Sub

	Dim tmpFsoObj, sFilePath
	Dim Xml, objNode, TBCount, FCount, objAtr
 	Dim IsTableExists, TbName, FName, FType
	Dim TbIndex, TbIndexField
	Dim i, j

	Set tmpFsoObj = CreateObject("Scripting.FileSystemObject")

	Call Top()

	Response.Write "<div class=""border"">" & VBCrlf
	Response.Write "<div class=""title"">&nbsp;<strong>第二步 - 数据库结构检查</strong></div>" & VBCrlf
	Response.Write "<p>" & VBCrlf

	sFilePath = Server.MapPath(Request.Form("dbpath"))
	If Not tmpFsoObj.FileExists (sFilePath) Then 
		Response.Write "<span style=""color: red;"">数据库文件地址不正确。</span><br />"
		Err = True
	End If

	sFilePath = Server.MapPath(Request.Form("xmlpath"))
	If Not tmpFsoObj.FileExists (sFilePath) Then 
		Response.Write "<span style=""color: red;"">数据库结构模版不正确。</span><br />"
		Err = True
	End If

	If Not Err Then
		ConnectionDataBase(Request.Form("dbpath"))

		Call LoadTableName()

		Set Xml  = Server.CreateObject("Microsoft.XMLDOM")
		Xml.Async= False
		Xml.Load(Server.MapPath(Request.Form("xmlpath")))

		Set objNode = Xml.documentElement

		TBCount = objNode.ChildNodes.Length - 1
		For i = 0 To TBCount
			Set objAtr = objNode.ChildNodes.item(i)

			TbName = objAtr.Attributes(0).Text

			If InStr(TableList, "," & TbName & ",") Then
				IsTableExists = True
			Else
				IsTableExists = False

				Call CreatTable(TbName)
			End If

			Call ClearIndex(TbName)
			
			FCount = objAtr.ChildNodes.Length - 1
			For j = 0 To FCount
				FName = objAtr.ChildNodes.item(j).tagName
				FType = objAtr.ChildNodes.item(j).Text

				If FName = "TableIndex" Then
					TbIndex = objAtr.ChildNodes.item(j).Attributes(0).Text
					
					If TbIndex = "PrimaryKey" Then
						Call AddIndex(TbName, TbIndex, FType, 1)
					Else
						Call AddIndex(TbName, TbIndex, FType, 0)
					End If
				Else
					If Not ChkField(TbName, FName) Then
						Call AddColumn(TbName, FName, FType)
					ElseIf LCase(FName) <> "id" Then
						Call ModColumn(TbName, FName, FType)
					End If
				End If
			Next
		Next

		Response.Write "<br /><span style=""color: blue;"">所有数据库结构检查工作已完成。</span>" & VBCrlf
	Else
		Response.Write "<br /><a href=""javascript: history.go(-1);"">返回上一步</a>" & VBCrlf
	End If

	Response.Write "</p>" & VBCrlf
	Response.Write "</div>" & VBCrlf
	Response.Write "<p />" & VBCrlf

	Call Foot()
End Sub

'添加索引
Sub AddIndex(TableName,IndexName,ColumnName,IndexType)
	On Error Resume Next
	Dim SQL

	SQL= "CREATE INDEX "& IndexName &" ON "& TableName &"("& ColumnName &")"

	If IndexType = 1 Then
		SQL=SQL& " WITH PRIMARY "
	End If

	Conn.Execute(SQL)
	If Err Then
		Response.Write "添加 "&TableName&" 表中索引<font color=blue>错误</font>，请手动添加，原因：" & Err.Description & "<BR>"
		Err.Clear
		Response.Flush
	End If
End Sub

'删除字段通用函数
Sub DelColumn(TableName,ColumnName)
	On Error Resume Next
	Conn.Execute("Alter Table "&TableName&" Drop "&ColumnName&"")
	If Err Then
		Response.Write "删除 "&TableName&" 表中字段<font color=blue>错误</font>，请手动将数据库中 <B>"&ColumnName&"</B> 字段删除，原因：" & Err.Description & "<BR>"
		Err.Clear
		Response.Flush
	Else
		Response.Write "删除 <font color=""#4455aa"">"&TableName&"</font> 表中字段 "&ColumnName&" 成功 <BR>"
		Response.Flush
	End If
End Sub

'更改字段通用函数
Sub ModColumn(TableName,ColumnName,ColumnType)
	On Error Resume Next
	Dim SQL
	SQL="Alter Table ["&TableName&"] Alter Column ["&ColumnName&"] "&ColumnType&""
	'Response.Write SQL&"<br>"
	Conn.Execute(SQL)

	If Err Then
		Response.Write "更改 "&TableName&" 表中字段属性<font color=blue>错误</font>，请手动将数据库中 <B>"&ColumnName&"</B> 字段更改为 <B>"&ColumnType&"</B> 属性，原因：" & Err.Description & "<BR>"
		Err.Clear
		Response.Flush
	End If
End Sub

'添加字段通用函数
Sub AddColumn(TableName,ColumnName,ColumnType)
	On Error Resume Next
	Conn.Execute("Alter Table "&TableName&" Add ["&ColumnName&"] "&ColumnType&"")
	If Err Then
		Response.Write "新建 "&TableName&" 表中字段<font color=blue>错误</font>，请手动将数据库中 <B>"&ColumnName&"</B> 字段建立，属性为 <B>"&ColumnType&"</B>，原因：" & Err.Description & "<BR>"
		Err.Clear
		Response.Flush
	Else
		Response.Write "新建 <font color=""#4455aa"">"&TableName&"</font> 表中字段 "&ColumnName&" 成功 <BR>"
		Response.Flush
	End If
End Sub

Sub CreatTable (TableName)
	On Error Resume Next

	Conn.Execute("CREATE TABLE ["&TableName&"] (Tmp int)")
	If Err Then
		Response.Write "添加 "&TableName&" 表<font color=blue>错误</font>，请手动在数据库中建立 <B>"&TableName&"</B> 表，原因：" & Err.Description & "<BR>"
		Err.Clear
		Response.Flush
	Else
		Response.Write "添加 <font color=""#4455aa"">"&TableName&"</font> 表成功 <BR>"
		Response.Flush

		Call DelColumn(TableName, "Tmp")
	End If
End Sub

Function ChkField (TableName, FieldName)
	On Error Resume Next

	Err.Clear

	Conn.Execute("SELECT " & FieldName & " FROM [" & TableName & "]")
	If Err Then
		ChkField = False
	Else
		ChkField = True
	End If
End Function

Sub LoadTableName ()
	Dim rsSchema
	Dim i

	Set rsSchema = Conn.openSchema(20)

	rsSchema.MoveFirst
	TableList = ","

	Do Until rsSchema.EOF
		If rsSchema("TABLE_TYPE")="TABLE" Then TableList = TableList & rsSchema("TABLE_NAME") & ","

		rsSchema.MoveNext
	Loop

	Set rsSchema = Nothing
End Sub

Sub ClearIndex (TableName)
	On Error Resume Next

	Dim rsSchema
	Dim i

	Set rsSchema = Conn.openSchema(12, Array(Empty, Empty, Empty, Empty, TableName))
	If rsSchema.EOF Then Exit Sub

	rsSchema.MoveFirst

	Do Until rsSchema.EOF
		Conn.Execute("DROP INDEX " & rsSchema("INDEX_NAME") & " ON " & TableName)

		rsSchema.MoveNext
	Loop

	Set rsSchema = Nothing
End Sub

Function ConnectionDataBase(dbPath)
	On Error Resume Next
	Err.Clear
	Dim ConnStr

	Set Conn = Server.CreateObject("ADODB.Connection")
	ConnStr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(dbPath)
	Conn.Open ConnStr
End Function
%>
