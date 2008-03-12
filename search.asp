<!--#Include File="init.asp" -->
<!--#Include File="include/inc.asp"-->
<!--#include file="include/_cls_teamplate.asp"-->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：Search.asp
'= 摘    要：搜索文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2006-11-13
'====================================================================

Dim Action
Dim ForTotal
Action=Request.QueryString ("action")

Select Case LCase(Action)
Case "query"
	Call Request_Query()
Case Else
	Call Main()
End Select
Call EA_Pub.Close_Obj
Set EA_Pub=Nothing

Sub Request_Query()
	Dim PageNum,KeyWord,Column,Types,StartDate,EndTime,IsInclude,PageSize
	
	PageNum	= EA_Pub.SafeRequest(1,"page",0,1,0)
	KeyWord	= EA_Pub.SafeRequest(1,"keyword",1,"",0)
	Column	= EA_Pub.SafeRequest(1,"column",1,"0|",0)
	Column	= Split(Column,"|")
	If UBound(Column)<>1 Then Column=Split("0|","|")
	Column(0)= EA_Pub.SafeRequest(0,Column(0),0,0,0)
	Column(1)= EA_Pub.SafeRequest(0,Column(1),1,"",0)
	Types	 = EA_Pub.SafeRequest(1,"field",0,0,0)
	StartDate= EA_Pub.SafeRequest(1,"stime",2,"1900-1-1",0)
	EndTime	 = EA_Pub.SafeRequest(1,"etime",2,FormatDateTime(Now()+1,2),0)
	IsInclude= EA_Pub.SafeRequest(1,"isinclude",0,0,0)
	PageSize = 15

	If KeyWord="" And Column(0)=0 And StartDate="1900-1-1" And CStr(EndTime)=CStr(FormatDateTime(Now()+1,2)) Then 
		Call Main()
		Exit Sub
	End If

	Dim FieldName(6),FieldValue(6)
	Dim PageCount,ReCount
	Dim i,WSQL
	Dim QueryArray,QueryList
	Dim PageContent
	Dim SQL
	Dim Temp,ListBlock
	Dim Template
	
	FieldName(0)="keyword"
	FieldName(1)="column"
	FieldName(2)="field"
	FieldName(3)="stime"
	FieldName(4)="etime"
	FieldName(5)="isinclude"
	FieldName(6)="action"
	
	FieldValue(0)=KeyWord
	FieldValue(1)=Join(Column,"|")
	FieldValue(2)=Types
	FieldValue(3)=StartDate
	FieldValue(4)=EndTime
	FieldValue(5)=IsInclude
	FieldValue(6)="query"
	
	If iDataBaseType=0 Then
		WSQL="Where IsPass="&EA_DBO.TrueValue&" And IsDel=0 And AddDate Between #"&StartDate&"# And #"&EndTime&"# "
	Else
		WSQL="Where IsPass="&EA_DBO.TrueValue&" And IsDel=0 And AddDate Between '"&StartDate&"' And '"&EndTime&"' "
	End If
	Select Case Types
	Case 1
		WSQL=WSQL&MakeSQLQuery("keyword",KeyWord)
	Case 2
		WSQL=WSQL&MakeSQLQuery("Author",KeyWord)
	Case 3
		WSQL=WSQL&MakeSQLQuery("Summary",KeyWord)
	Case Else
		WSQL=WSQL&MakeSQLQuery("Title",KeyWord)
	End Select
	
	If Column(0)>0 Then 
		If IsInclude=1 Then 
			WSQL=WSQL&" And Left(ColumnCode,"&Len(Column(1))&")='"&Column(1)&"'"
		Else
			WSQL=WSQL&" And ColumnId="&Column(0)
		End If
	End If

	Set Template=New cls_NEW_TEMPLATE

	PageContent=EA_Temp.Load_Template(0,"search")

	SQL="Select Count([Id]) From [NB_Content] "&WSQL
	ReCount=EA_DBO.DB_Query(SQL)(0,0)
	PageCount=EA_Pub.Stat_Page_Total(PageSize,ReCount)
	If PageNum>PageCount And PageCount>0 Then PageNum=PageCount
	
	ListBlock=Template.GetBlock("list",PageContent)

	If PageCount>0 Then 
		WSQL = WSQL & " ORDER BY TrueTime DESC"
		SQL="Select [Id],ColumnId,ColumnName,IsImg,IsTop,TColor,Title,AddDate,Author,ViewNum,CommentNum,Summary From [NB_Content] "&WSQL
		QueryArray=EA_DBO.DB_CutPageQuery(SQL,PageNum,PageSize)

		If IsArray(QueryArray) Then
			ForTotal = UBound(QueryArray,2)

			For i=0 To ForTotal
				Temp=ListBlock
  
				Template.SetVariable "ColumnUrl",EA_Pub.Cov_ColumnPath(QueryArray(1,i),EA_Pub.SysInfo(18)),Temp
				Template.SetVariable "Column",QueryArray(2,i),Temp
				Template.SetVariable "ArticleTitle",EA_Pub.Add_ArticleColor(QueryArray(5,i),QueryArray(6,i)),Temp
				Template.SetVariable "ArticleUrl",EA_Pub.Cov_ArticlePath(QueryArray(0,i),QueryArray(7,i),EA_Pub.SysInfo(18)),Temp
				Template.SetVariable "Icon",EA_Pub.Chk_ArticleType(QueryArray(3,i),QueryArray(4,i)),Temp
				Template.SetVariable "Time",QueryArray(7,i),Temp
				Template.SetVariable "View",QueryArray(9,i),Temp
				Template.SetVariable "Author",QueryArray(8,i),Temp
				Template.SetVariable "Comment",QueryArray(10,i),Temp
				Template.SetVariable "Summary",EA_Pub.Full_HTMLFilter(QueryArray(11,i)),Temp

				Template.SetBlock "list",Temp,PageContent
			Next 
		End If
	End If
	Template.CloseBlock "list",PageContent
	
	EA_Temp.Title=EA_Pub.SysInfo(0)&" - 站点搜索"
	EA_Temp.Nav="<a href=""./""><b>"&EA_Pub.SysInfo(0)&"</b></a> - 搜索文章"

	PageContent=EA_Temp.Replace_PublicTag(PageContent)
	PageContent=Replace(PageContent,"{$PageNumNav$}",EA_Temp.PageList(PageCount,PageNum,FieldName,FieldValue))

	Call EA_Temp.Find_TemplateTagByInput("Config","",PageContent)
	
	PageContent=Replace(PageContent,"{$Query_List$}",QueryList)
	
	Response.Write PageContent
End Sub

Function MakeSQLQuery(QueryField,QueryStr)
	Dim TagStart,TagEnd
	Dim TempStr,TempArray
	Dim FullQueryStr
	Dim i,Way
	
	'先找引号定界符
	Do
		TagStart=InStr(QueryStr,"\")
		If TagStart>0 Then 
			TagEnd=InStr(TagStart+1,QueryStr,"\")
			
			TempStr=Mid(QueryStr,TagStart+1,TagEnd-TagStart-1)
			TempStr=Replace(TempStr," ","#")
			
			QueryStr=Left(QueryStr,TagStart-1)&TempStr&Right(QueryStr,Len(QueryStr)-TagEnd)
		End If
	Loop While TagStart>0
	
	
	QueryStr = Replace(QueryStr,"|"," @")	'处理or定界符
	TempArray= Split(QueryStr," ")			'分隔关键字
	ForTotal = UBound(TempArray)

	For i=0 To ForTotal
		If Left(TempArray(i),1)="@" Then
			FullQueryStr=FullQueryStr&" Or "&QueryField
			TempArray(i)=Right(TempArray(i),Len(TempArray(i))-1)
		Else
			FullQueryStr=FullQueryStr&" And "&QueryField
		End If
		
		If Left(TempArray(i),1)="-" Then 
			FullQueryStr=FullQueryStr&" Not "
			TempArray(i)=Right(TempArray(i),Len(TempArray(i))-1)
		End If
		
		FullQueryStr=FullQueryStr&" Like '%"&TempArray(i)&"%'"
		
		FullQueryStr=Replace(FullQueryStr,"%$","")
		FullQueryStr=Replace(FullQueryStr,"$%","")
		FullQueryStr=Replace(FullQueryStr,"#"," ")
	Next
	
	MakeSQLQuery=FullQueryStr
End Function

Sub Main
	Dim ColumnArray,ColumnList
	Dim i,Level

	ColumnArray=EA_DBO.Get_Column_List()
	If IsArray(ColumnArray) Then 
		ForTotal = UBound(ColumnArray,2)

		For i=0 To ForTotal
			Level=(Len(ColumnArray(2,i))/4-1)
			ColumnList=ColumnList&"<option value="""&ColumnArray(0,i)&"|"&ColumnArray(2,i)&""">"
			If Level>0 Then ColumnList=ColumnList&"├"
			ColumnList=ColumnList&String(Level,"-")
			ColumnList=ColumnList&ColumnArray(1,i)&"</option>"
		Next
	End If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html;charset=utf-8">
<meta http-equiv="Content-Language" content="zh-CN">
<meta name="keywords" content="<%=Replace(EA_Pub.SysInfo(16),"|",",")%>">
<meta name="Description" content="<%=EA_Pub.SysInfo(17)%>">
<meta name="generator" content="NB文章系统(NBArticle)">
<title><%=EA_Pub.SysInfo(0)%> - 站点搜索</title>
<script src="js/jsdate.js"></script>
<script src="js/public.js"></script>
<table width="762" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <form name="form1" method="post" action="?action=query">
    <tr> 
      <td bgcolor="#FFFFFF"> <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr> 
            <td height="22" align="center" bgcolor="efefef">高 级 查 找</td>
          </tr>
          <tr> 
            <td height="20" align="center" bgcolor="#FFFFFF">搜索范围： 
              <input name="field" type="radio" value="0" checked>
              标题 
              <input type="radio" name="field" value="1">
              关键字 
              <input type="radio" name="field" value="2">
              作者 
              <input type="radio" name="field" value="3">
              摘要 </td>
          </tr>
          <tr> 
            <td height="20" align="center" bgcolor="#FFFFFF">关键字： 
              <input name="keyword" type="text" id="keyword"></td>
          </tr>
          <tr> 
            <td height="20" align="center" bgcolor="#FFFFFF">栏目分类： 
              <select name="column" class="iptA">
                <option value="0">--栏 目--</option>
                <%=ColumnList%>
              </select><input type="checkbox" name="isinclude" value="1">包含子栏目</td>
          </tr>
		  <tr> 
            <td height="20" align="center" bgcolor="#FFFFFF">开始时间：<input type="text" name="stime" id="stime" size="10" readonly>&nbsp;<a href="javascript:vod();" onclick="SD(this,'stime');"><img border="0" src="images/public/date_picker.gif" width="30" height="19" align="absmiddle"></a>&nbsp;&nbsp;结束时间：<input type="text" name="etime" id="etime" size="10" readonly>&nbsp;<a href="javascript:vod();" onclick="SD(this,'etime');"><img border="0" src="images/public/date_picker.gif" width="30" height="19" align="absmiddle"></a></td>
          </tr>
          <tr> 
            <td height="22" align="center" bgcolor="#FFFFFF"> <input type="submit" name="Submit" value="开始"> 
              <input type="reset" name="Submit2" value="重置"> </td>
          </tr>
        </table></td>
    </tr>
  </form>
</table>
<br>
<table width="762" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC">
  <tr> 
    <td height="22" align="center" bgcolor="efefef">帮 助 说 明</td>
  </tr>
  <tr> 
    <td bgcolor="ffffff"> <table width="90%" height="100" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td bgcolor="ffffff"> <li>1、 空格连接=and，如 <font color="#006600">你好 我要</font>=<font color="#0033FF">%你好% 
              and %我要%</font> 
            <li>2、 避免内容包含字符=-，如 <font color="#006600">你好 -我要</font>=<font color="#0033FF">%你好% 
              and not like %我要%</font> 
            <li>3、 |=or，如 <font color="#006600">你好|我要</font>=<font color="#0033FF">%你好% 
              or %我要%</font> 
            <li>4、 词组搜索用双引号包含，如 <font color="#006600">\i love this game\</font>=<font color="#0033FF">%i 
              love this game%</font>，而非=<font color="#CC3300">i and love and this 
              and game</font> 
            <li>5、 $为定界符，如 <font color="#006600">$你好</font>=<font color="#0033FF">以 
              你好 开头</font>的字符，<font color="#006600">你好$</font>=<font color="#0033FF">以 
              你好 结尾</font>的字符 <br>
              <br>
              组合查询 
            <li>如 <font color="#006600">\i love this game\|-你好</font>=<font color="#0033FF">%i 
              love this game% and not like %你好%</font> 
            <li>如 <font color="#006600">我要$|-$你好</font>=<font color="#0033FF">%我要 
              or not like 你好%</font> 
            <li>如 <font color="#006600">$\i love this game\ $你好$</font>=<font color="#0033FF">i 
              love this game% and like 你好</font> </td>
        </tr>
      </table></td>
  </tr>
</table>
<%
End Sub
%>
