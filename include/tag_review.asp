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
Function Load_NewReview(Parameter)
	Dim TempStr,TempArray
	Dim i
	Dim ForTotal
	
	TempArray=EA_DBO.Get_Review_NewList(Parameter(0),Parameter(1))
	If IsArray(TempArray) Then 
		TempStr="<table>"
		ForTotal = UBound(TempArray,2)

		For i=0 To ForTotal
			TempStr=TempStr&"<tr>"
			TempStr=TempStr&"<td style=""TEXT-ALIGN: left;""><a href=""review.asp?articleid="&TempArray(0,i)&""">"&EA_Pub.Un_Full_HTMLFilter(TempArray(1,i))&"</a></td>"
			TempStr=TempStr&"<td style=""TEXT-ALIGN: center;"">"&TempArray(2,i)&"</td>"
			TempStr=TempStr&"<td style=""COLOR: #800000;TEXT-ALIGN: center;"">"&TempArray(3,i)&"</font></td>"
			TempStr=TempStr&"</tr>"
		Next
		TempStr=TempStr&"</table>"
	Else
		TempStr=""
	End If
	
	Load_NewReview=TempStr
End Function
%>