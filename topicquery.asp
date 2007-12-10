<!--#Include File="conn.asp" -->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：TopicQuery.asp
'= 摘    要：首页文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2006-07-27
'====================================================================

'参数一(Order)：筛选条件，1=最新文章，2=专题文章，3=热点文章
'参数二(SortId)：筛选栏目id，0则为全系统筛选
'参数三(TopNum)：调用的记录条数
'参数四(Style)：调用样式,1=滚动横式，2=滚动竖式，3=静态竖式单列，4=静态竖式双列(该参数只在Order=1下有效)
'参数五(IsSort)：是否显示分类,0=不显示，1=显示

Const SysUrl="http://www.nbarticle.com/"		'文章系统地址，切记最后一定要有/号

Dim Order,SortId,TopNum,Style,IsSort
Dim i,WSQL,SQL,Rs
Dim Topic
Dim iTrue
Dim ForTotal

ConnectionDatabase

Order=SafeRequest("order",1)
If Order="" Or CLng(Order)<=0 Then Order=1
SortId=SafeRequest("sortId",1)
If SortId="" Or CLng(SortId)<0 Then SortId=0
TopNum=SafeRequest("topnum",1)
If TopNum="" Or CLng(TopNum)<0 Then TopNum=0
Style=SafeRequest("style",1)
If Style="" Or CLng(Style)<0 Then Style=1
IsSort=SafeRequest("issort",1)
If IsSort="" Or CLng(IsSort)<0 Then IsSort=0

If SortId<>0 Then WSQL=" and ColumnId="&SortId

If iDataBaseType = 0 Then
	iTrue = -1
Else
	iTrue = 1
End If

Select Case CInt(Order)
Case 1
	'参数说明：
	'0=文章id，1=栏目id，2=栏目名，3=文章标题，4=添加时间，5=标题颜色号，6=浏览人数
	SQL="Select Top "&TopNum&" Id,ColumnId,ColumnName,Title,AddDate,TColor,ViewNum From [NB_Content] Where IsDel=0 And IsPass="&iTrue&WSQL&" Order By TrueTime Desc"
Case 2
	SQL="Select Top "&TopNum&" Id,ColumnId,ColumnName,Title,AddDate,TColor,ViewNum From [NB_Content] Where IsDel=0 And IsPass="&iTrue&" And IsDis="&iTrue&WSQL&" Order By TrueTime Desc"
Case 3
	SQL="Select Top "&TopNum&" Id,ColumnId,ColumnName,Title,AddDate,TColor,ViewNum From [NB_Content] Where IsDel=0 And IsPass="&iTrue&" And ViewNum>0"&WSQL&" Order By ViewNum Desc"
End Select
'Response.Write sql
Set Rs=Conn.Execute(SQL)
If Not Rs.eof And Not Rs.bof Then Topic=Rs.GetRows()

If IsArray(Topic) Then 
	Select Case CInt(Order)
	Case 1
		Select Case CInt(Style)
		Case 1
			Response.Write "document.write ('<marquee scrollamount=2 width=""100%"" border=1 valign=middle scrolldelay=5 height=25 onmouseover=""this.stop()"" onmouseout=""this.start()"" direction=left behavior=loop>"
			ForTotal = Ubound(Topic,2)

			For i=0 To ForTotal
				If CBool(IsSort) Then Response.Write "[<a href="""&SysUrl&"list.asp?classid="&Topic(1,i)&""">"&Topic(2,i)&"</a>]"
				Response.Write "·<a href="""&SysUrl&"Article.asp?ArticleId="&Topic(0,i)&""" target=""_blank"">"
				Response.Write TitleAddColor(Topic(5,i),CutTitle(15,replace(Topic(3,i),"&nbsp;"," ")))
				Response.Write "</a>"
				Response.Write CheckNow(Topic(4,i))
				Response.Write "&nbsp;("
				Response.Write Month(Topic(4,i))&"/"&Day(Topic(4,i))
				Response.Write ")&nbsp;&nbsp;&nbsp;&nbsp;"
			Next
			Response.Write "</marquee>');"&chr(10)
		Case 2
			Response.Write "document.write ('<marquee scrollamount=2 width=""100%"" border=1 valign=middle scrolldelay=10 height=100 onmouseover=""this.stop()"" onmouseout=""this.start()"" direction=up behavior=loop>"
			ForTotal = Ubound(Topic,2)

			For i=0 To ForTotal
				If CBool(IsSort) Then Response.Write "[<a href="""&SysUrl&"list.asp?classid="&Topic(1,i)&""">"&Topic(2,i)&"</a>]"
				Response.Write "·<a href="""&SysUrl&"Article.asp?ArticleId="&Topic(0,i)&""" target=""_blank"">"
				Response.Write TitleAddColor(Topic(5,i),CutTitle(15,replace(Topic(3,i),"&nbsp;"," ")))
				Response.Write "</a>"
				Response.Write CheckNow(Topic(4,i))
				Response.Write "&nbsp;("
				Response.Write Month(Topic(4,i))&"/"&Day(Topic(4,i))
				Response.Write ")<br>"
			Next
			Response.Write "</marquee>');"&chr(10)
		Case 3
			Response.Write "document.write ('<table width=""100%"" align=""center"" cellpadding=""0"" cellspacing=""0"" border=""0"">');"&chr(10)
			ForTotal = Ubound(Topic,2)

			For i=0 To ForTotal
				Response.Write "document.write ('<tr>');"&chr(10)
				Response.Write "document.write ('<td align=""left"">"
				If CBool(IsSort) Then Response.Write "[<a href="""&SysUrl&"list.asp?classid="&Topic(1,i)&""">"&Topic(2,i)&"</a>]"
				Response.Write "·<a href="""&SysUrl&"Article.asp?ArticleId="&Topic(0,i)&""" target=""_blank"">"
				Response.Write TitleAddColor(Topic(5,i),CutTitle(15,replace(Topic(3,i),"&nbsp;"," ")))
				Response.Write "</a>"
				Response.Write CheckNow(Topic(4,i))
				Response.Write "</td>');"&chr(10)
				Response.Write "document.write ('<td align=""center"">"
				Response.Write Month(Topic(4,i))&"/"&Day(Topic(4,i))
				Response.Write "</td>');"&chr(10)
				Response.Write "document.write ('</tr>');"&chr(10)
			Next
			Response.Write "document.write ('</table>');"&chr(10)
		Case 4
			Response.Write "document.write ('<table width=""100%"" align=""center"" cellpadding=""0"" cellspacing=""0"" border=""0"">');"&chr(10)
			Response.Write "document.write ('<tr>');"&chr(10)
			ForTotal = Ubound(Topic,2)

			For i=0 To ForTotal			
				Response.Write "document.write ('<td align=""left"" width=""50%"">"
				If CBool(IsSort) Then Response.Write "[<a href="""&SysUrl&"list.asp?classid="&Topic(1,i)&""">"&Topic(2,i)&"</a>]"
				Response.Write "·<a href="""&SysUrl&"Article.asp?ArticleId="&Topic(0,i)&""" target=""_blank"">"
				Response.Write TitleAddColor(Topic(5,i),CutTitle(15,replace(Topic(3,i),"&nbsp;"," ")))
				Response.Write "</a>"
				Response.Write CheckNow(Topic(4,i))
				Response.Write "&nbsp;("
				Response.Write Month(Topic(4,i))&"/"&Day(Topic(4,i))
				Response.Write ")</td>');"&chr(10)
				If (i+1) Mod 2=0 Then Response.Write "document.write ('</tr><tr>');"&chr(10)
			Next
			Response.Write "document.write ('</tr>');"&chr(10)
			Response.Write "document.write ('</table>');"&chr(10)
		End Select
	Case 2
		Response.Write "document.write ('<table width=""100%"" align=""center"" cellpadding=""0"" cellspacing=""0"" border=""0"">');"&chr(10)
		ForTotal = Ubound(Topic,2)

		For i=0 To ForTotal
			Response.Write "document.write ('<tr>');"&chr(10)
			Response.Write "document.write ('<td align=""center"">"
			Response.Write "<a href="""&SysUrl&"list.asp?ClassId="&Topic(1,i)&""" target=""_blank"">["
			Response.Write Topic(2,i)
			Response.Write "]</a>"
			Response.Write CheckNow(Topic(4,i))
			Response.Write "</td>');"&chr(10)
			Response.Write "document.write ('<td align=""left"">"
			Response.Write "·<a href="""&SysUrl&"Article.asp?ArticleId="&Topic(0,i)&""" target=""_blank"">"
			Response.Write TitleAddColor(Topic(5,i),CutTitle(15,replace(Topic(3,i),"&nbsp;"," ")))
			Response.Write "</a>"
			Response.Write "</td>');"&chr(10)
			Response.Write "document.write ('</tr>');"&chr(10)
		Next
		Response.Write "document.write ('</table>');"&chr(10)
	Case 3
		Response.Write "document.write ('<table width=""100%"" align=""center"" cellpadding=""0"" cellspacing=""0"" border=""0"">');"&chr(10)
		ForTotal = Ubound(Topic,2)

		For i=0 To ForTotal
			Response.Write "document.write ('<tr>');"&chr(10)
			Response.Write "document.write ('<td align=""left"">"
			Response.Write "·<a href="""&SysUrl&"Article.asp?ArticleId="&Topic(0,i)&""" target=""_blank"">"
			Response.Write TitleAddColor(Topic(5,i),CutTitle(15,replace(Topic(3,i),"&nbsp;"," ")))
			Response.Write "</a>"
			Response.Write CheckNow(Topic(4,i))
			Response.Write "</td>');"&chr(10)
			Response.Write "document.write ('<td align=""center"">"
			Response.Write "<font color=ff00000>"
			Response.Write Topic(6,i)
			Response.Write "</font>"
			Response.Write "</td>');"&chr(10)
			Response.Write "document.write ('</tr>');"&chr(10)
		Next
		Response.Write "document.write ('</table>');"&chr(10)
	End Select
End If

Function TitleAddColor(TColor,Title)
	Dim TempStr
	If Not IsNumeric(TColor) Then
		TempStr=""
	Else
		Select Case CLng(TColor)
		Case 1
		'红色
			TempStr="#FF0000"
		Case 2
		'绿色
			TempStr="#37a61c"
		Case 3
		'兰色
			TempStr="#0066CC"
		End Select
	End If
	If TempStr="" Or IsNull(TempStr) Then 
		TitleAddColor=Title
	Else
		TitleAddColor="<font color="""&TempStr&""">"&Title&"</font>"
	End If
End Function

Function CutTitle(TLen,Title)
	If Len(Title)>TLen Then
		CutTitle=Left(Title,TLen)
	Else
		CutTitle=Title
	End If
End Function

Function SafeRequest(ParaName,ParaType)
	Dim ParaValue
	ParaValue=Request(ParaName)
	If ParaType=1 Then
		If Not isNumeric(ParaValue) Then
			sysErr(9)
		End If
	Else
		ParaValue=Replace(ParaValue,"'","''")
	End if
	SafeRequest=ParaValue
End function

Function CheckNow(PostDate)
	If DateDiff("d",PostDate,Now())=0 Then 
		CheckNow=" <img src=""images/new.gif"" border=0 align=absmiddle>"
	End If
End Function
%>