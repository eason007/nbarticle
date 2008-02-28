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
'= 最后日期：2008-02-28
'====================================================================

Function Load_Friend(Parameter)
	Dim FriendList,WidthPercent
	Dim TempStr,i
	Dim ForTotal
	
	WidthPercent=100/CLng(Parameter(3))

	FriendList=EA_DBO.Get_Friend_List(Parameter(1),Parameter(0),Parameter(2))
	
	If IsArray(FriendList) Then 
		TempStr="<table>"
		TempStr=TempStr&"<tr>"
		ForTotal = UBound(FriendList,2)

		For i=0 To ForTotal
			If Parameter(2)="1" Then 
				TempStr=TempStr&"<td style""WIDTH: "&WidthPercent&"%;""><a href="""&FriendList(1,i)&"""><img src="""&FriendList(2,i)&""" alt="""&FriendList(0,i)&""" /></a></td>"
			Else
				TempStr=TempStr&"<td style""WIDTH: "&WidthPercent&"%;""><a href="""&FriendList(1,i)&""" title="""&FriendList(0,i)&""">"&FriendList(0,i)&"</a></td>"
			End If

			If (i+1) Mod CLng(Parameter(3))=0 Then TempStr=TempStr&"</tr><tr>"
		Next
		If i Mod CLng(Parameter(3))=0 Then TempStr=TempStr&"</tr>"

		TempStr=TempStr&"</table>"
	Else
		TempStr = ""
	End If

	Load_Friend=TempStr
End Function
%>