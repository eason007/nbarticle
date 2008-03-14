<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2004 - 2008 Eason Chan All Rights Reserved.
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

	Function Load_Vote(Parameter)
		Dim Id,Title,VoteText,Mtype
		Dim Result,Content,j
		Dim TempArray
		Dim ForTotal
		
		TempArray=EA_DBO.Get_Vote_Info(Parameter(0))
		If IsArray(TempArray) Then 
			If TempArray(5,0)=0 Then
				Id=TempArray(0,0)
				Title=TempArray(1,0)
				VoteText=TempArray(2,0)
				Mtype=TempArray(4,0)
				
				Content=split(VoteText,"|")
				
				Result=Result&"<form action="""&SystemFolder&"vote.asp"" method=""post"" id=""vote_"&Id&"""><table><tr><td>"&Title&"</td></tr>"
		
				IF Mtype=False Then
					ForTotal = Ubound(Content)

					For i=0 To ForTotal
						Result=Result&"<tr><td>"
						Result=Result&"<input type=""radio"" name=""vote"" value="""&i&""" />&nbsp;&nbsp;&nbsp;"&Content(i)
						Result=Result&"</td></tr>"
					Next
				End if

				If Mtype=True Then
					ForTotal = Ubound(Content)

					For i=0 To ForTotal
						Result=Result&"<tr><td>"
						Result=Result&"<input type=""checkbox"" name=""vote"" value="""&i&""" />&nbsp;&nbsp;&nbsp;"&Content(i)
						Result=Result&"</td></tr>"
					Next
				End if
		
				Result=Result&"<tr><td><input name=""votetype"" id=""votetype"" type=""hidden"" value="""&CInt(Mtype)&""" /><input name=""voteid"" id=""voteid"" type=""hidden"" value="""&Id&""" /><input type=""button"" name=""submit"" value=""投票"" onclick=""window.open(submit_vote("&Id&"),'_blank','scrollbars=yes,width=645,height=380')"" />&nbsp;<input type=""button"" name=""view"" value=""查看"" onclick=""window.open('vote.asp?VoteId="&ID&"','_blank','scrollbars=yes,width=645,height=380')"" /></td></tr>"
				Result=Result&"</table></form>"
			End If
		End If
		
		Load_Vote=Result
	End Function

%>