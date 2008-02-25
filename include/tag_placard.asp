<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：cls_template.asp
'= 摘    要：模版类文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-02-24
'====================================================================

	Function Load_Placard(Parameter)
		Dim TempArray,TempStr
		Dim i
		Dim CutStr
		Dim ForTotal
		
		If Parameter(1)=0 Then 
			CutStr="&nbsp;&nbsp;&nbsp;"
		Else
			CutStr="<br />"
		End If
		
		If CLng(Parameter(0))=0 Then Parameter(0)=10
		
		TempArray=EA_DBO.Get_PlacardTopList(Parameter(0))
		If IsArray(TempArray) Then 
			ForTotal = UBound(TempArray,2)

			For i=0 To ForTotal
				TempStr=TempStr&"<img src="""&SystemFolder&"images/public/gb.gif"" alt="""" />&nbsp;<a href=""#"" onclick=""javascript:window.open('"&SystemFolder&"viewplacard.asp?postid="&TempArray(0,i)&"','','scrollbars=yes,height=350,width=550')"">"&TempArray(1,i)&"</a>&nbsp;<font color=""#999999"">("&FormatDateTime(TempArray(2,i),2)&")</font>"&CutStr
			Next
		Else
			TempStr=TempStr&"<font color=""#800000"">欢迎光临"&EA_Pub.SysInfo(0)&"。</font>"
		End If
		
		Load_Placard=TempStr
	End Function
%>