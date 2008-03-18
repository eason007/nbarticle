<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2004 - 2008 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：tag_adsense.asp
'= 摘    要：vote模版标签文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-03-18
'====================================================================

Sub MakeVote(ByRef PageContent)
	If EA_Temp.ChkTag("Vote.Single", PageContent) Then
		VoteSingle PageContent
	End If
End Sub

Sub VoteSingle (ByRef PageContent)
	Dim Parameter
	Dim List
	Dim Temp, ForTotal, i

	Do 
		Parameter = EA_Temp.GetParameter("Vote.Single", PageContent)
		If Not IsArray(Parameter) Then Exit Do

		List = EA_DBO.Get_Vote_Info(Parameter(0))
		If Not IsArray(List) Then 
			Exit Do
		Else
			If List(5, 0) = 0 Then
				Id		= List(0, 0)
				Title	= List(1, 0)
				VoteText= List(2, 0)
				Mtype	= List(4, 0)
				
				Content	= Split(VoteText, "|")
				
				Temp	= Temp & "<form action=""" & SystemFolder & "action.asp"" method=""post"" id=""vote_" & Id & """><table><tr><td>" & Title & "</td></tr>"
		
				If Mtype = False Then
					Mtype = "redio"
				Else If Mtype = True Then
					Mtype = "checkbox"
				End If
				
				ForTotal = Ubound(Content)

				For i = 0 To ForTotal
					Temp = Temp & "<tr><td>"
					Temp = Temp & "<input type=""" & Mtype & """ name=""vote"" value=""" & i & """ />&nbsp;&nbsp;&nbsp;" & Content(i)
					Temp = Temp & "</td></tr>"
				Next
		
				Temp = Temp&"<tr><td><input name=""votetype"" type=""hidden"" value=""" & CInt(Mtype) & """ />"
				Temp = Temp&"<input name=""voteid"" type=""hidden"" value=""" & Id & """ />'
				Temp = Temp&"<input type=""button"" name=""submit"" value=""投票"" onclick=""window.open(submit_vote(" & Id & "),'_blank','scrollbars=yes,width=645,height=380')"" />"
				Temp = Temp&"&nbsp;<input type=""button"" name=""view"" value=""查看"" onclick=""window.open('vote.asp?VoteId=" & Id & "','_blank','scrollbars=yes,width=645,height=380')"" /></td></tr>"
				Temp = Temp & "</table></form>"
			End If

			Call EA_Temp.SetSingle(Temp, PageContent)
		End If
	Loop While 1 = 1
End Sub
%>