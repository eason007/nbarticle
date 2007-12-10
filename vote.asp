<!--#Include File="conn.asp" -->
<!--#Include File="include/inc.asp"-->
<!--#Include File="include/vml_cake.asp"-->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：Vote.asp
'= 摘    要：投票文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2006-07-27
'====================================================================

Dim VoteId,VoteChoose,VoteType,VoteText,VoteNum,VoteUBound
Dim i,VoteInfo
Dim IsVoted
Dim ForTotal

VoteId=EA_Pub.SafeRequest(3,"voteid",0,0,3)
VoteChoose=EA_Pub.SafeRequest(3,"vote",1,"",3)
VoteChoose=Split(VoteChoose,",")
VoteType=EA_Pub.SafeRequest(3,"votetype",0,0,3)
If Request.Cookies(sCacheName&"Vote"&VoteId)="" Then 
	IsVoted=False
Else
	IsVoted=True
End If

If Not IsVoted Then
	If Ubound(VoteChoose)<0 Then
		ErrMsg="你还没有投票，不允许此操作。"
		Call EA_Pub.ShowErrMsg(0,0)
	End If
	
	If EA_Pub.SysInfo(10)="0" Then 
		ErrMsg="对不起，系统已设定游客不允许参加投票。"
		Call EA_Pub.ShowErrMsg(0,0)
	End If
End If

VoteInfo=EA_DBO.Get_Vote_Info(VoteId)
If IsArray(VoteInfo) Then
	If UBound(VoteChoose)>0 And VoteInfo(4,0)=0 Then 
		ErrMsg="你投票的项目不符合调查要求。"
		Call EA_Pub.ShowErrMsg(0,0)
	End If

	If VoteInfo(5,0)<>0 Then 
		ErrMsg="投票已关闭"
		Call EA_Pub.ShowErrMsg(0,0)
	End If

	VoteText=VoteInfo(2,0)
	VoteText=Split(VoteText,"|")
	VoteNum=VoteInfo(3,0)
	VoteNum=Split(VoteNum,"|")

	If Not IsVoted Then
		ForTotal = UBound(VoteChoose)
		For i=0 To ForTotal
			VoteNum(CLng(VoteChoose(i)))=VoteNum(CLng(VoteChoose(i)))+1
		Next
		VoteNum=Join(VoteNum,"|")
	
		Call EA_DBO.Set_Vote_SaveVoted(VoteId,VoteNum)

		Response.Cookies(sCacheName&"Vote"&VoteId)=1
		Response.Cookies(sCacheName&"Vote"&VoteId).Expires=Date()+24
		
		VoteNum=Split(VoteNum,"|")
	End If
	
	Call ShowResult(VoteText,VoteNum,VoteInfo(0,0),IsVoted)
Else
	ErrMsg="请不要手动修改地址参数,同时请填写好所有必须填写的项目！"
	Call EA_Pub.ShowErrMsg(0,0)
End If

Call EA_Pub.Close_Obj
Set EA_Pub=Nothing

Sub ShowResult(VoteText,VoteNum,VoteTitle,IsVote)
	Dim i,LoopNum,VoteCount,rate,j
	Dim CakeArray
	
	VoteCount=0
	If Not IsArray(VoteText) And Not IsArray(VoteNum) Then Response.End
	
	i=UBound(VoteText)
	If i>=5 Then i=4
	
	ReDim CakeArray(i+1,1)
	ForTotal = UBound(VoteText)

	For i=0 To ForTotal
		If i>=5 Then Exit For
		CakeArray(i+1,0)=VoteText(i)
		CakeArray(i+1,1)=VoteNum(i)
	Next
%>
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
<!--[if !mso]>
<style>
v\:*         { behavior: url(#default#VML) }
o\:*         { behavior: url(#default#VML) }
.shape       { behavior: url(#default#VML) }
</style>
<![endif]-->
<head>
<style type="text/css">
<!--
td {
	font-size: 12px;
}
-->
</style>
<title>查看投票</title>
</head>
<body topmargin=5 leftmargin=0 scroll=AUTO>
<table width="600" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#E8F7FF">
  <tr> 
    <td height="25" align="center" bgcolor="#C4ECFF">投票结果</td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF"> <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="150" align="center"><%Call Cake_Table(CakeArray,280,65,90,80,"B")%></td>
        </tr>
      </table>
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="22"><table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="ffffff">
              <tr> 
                <td height="25" bgcolor="#efefef" colspan="3">&nbsp;投票题目：<%=VoteTitle%></td>
              </tr>
              <%
			ForTotal = Ubound(VoteNum)

		   	For i=0 To ForTotal
				VoteCount=VoteCount+VoteNum(i)	
			Next

			j=1
			ForTotal = Ubound(VoteText)

			For i=0 To ForTotal
				if VoteCount>0 Then rate=FormatNumber((VoteNum(i)/VoteCount) * 100,2)
			  %>
              <tr> 
                <td width="200" height="25">&nbsp;<%=i+1&". "&VoteText(i)%> </td>
                <td width="80"><%=" 票数："&VoteNum(i)%></td>
                <td><img src="images/public/bar<%=j%>.gif" width="<%=rate%>" height="9" align="absmiddle"><img src="images/public/bar<%=j%>_1.gif" height="9" align="absmiddle"> 
                  <%=rate&"%"%> </td>
              </tr>
              <%j=j+1
				If j Mod 5 =0 Then j=1
              Next%>
            </table></td>
        </tr>
      </table></td>
  </tr>
  <tr bgcolor="#E8F7FF"> 
    <td height="22">&nbsp; <%If IsVote Then Response.Write "你已经投过票！" Else Response.Write "<font color=800000>这是你第一次投票！</font>"%></td>
  </tr>
</table>
</body>
</html>
<%End Sub%>
