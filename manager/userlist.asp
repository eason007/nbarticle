
<!--#Include File="comm/inc.asp" -->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2004 - 2008 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：/Manager/UserList.asp
'= 摘    要：后台-组内会员列表文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2006-07-27
'====================================================================

Call EA_Manager.Chk_IsMaster

If Not EA_Manager.Chk_Power(Admin_Power,"21") Then 
	ErrMsg=str_Comm_NotAccess
	Call EA_Manager.Error(1)
End If

Dim PostId
Dim Temp
PostId=EA_Pub.SafeRequest(3,"postid",0,0,0)

If Request.QueryString ("atcion")="move" then 
	Call move
End If
%>
<style type="text/css">
body {
	FONT-SIZE: 12px;
	BACKGROUND: #efefef;
	padding: 0;
	margin: 3px 3px 3px 3px;
}
th {
	BACKGROUND: #c5d7e2;
	height: 25px;
}
</style>
<script type="text/javascript">
function CheckAll(form)
{
	for (var i=0;i<form.elements.length;i++)
	{
		var e = form.elements[i];
		if (e.name != 'chkall'){
			e.checked = form.chkall.checked;
		}
	}
}
</script>
<table width="100%" cellpadding="1" cellspacing="1">
  <form name="form1" method="post" action="?atcion=move&postid=<%=PostId%>">
    <tr height="22" align="center">
      <th width="8%"><%=str_Comm_Select%></th>
      <th><%=str_Member_Account%></th>
      <th width="20%">E-Mail</th>
      <th width="25%"><%=str_Member_RegDate%></th>
      <th width="6%"><%=str_Member_State%></th>
    </tr>
    <%
    Dim i,TopicList
    Dim Count,OutStr
	Dim FieldName(0),FieldValue(0)
	Dim ForTotal
	
    Page=EA_Pub.SafeRequest(3,"page",0,1,0)
    
    FieldName(0)="postid"
	FieldValue(0)=PostId
    
	Count=EA_M_DBO.Get_Group_ForMemberTotal(PostId)(0,0)
	If Count=0 Then 
		OutStr="<tr height='22' bgcolor='ffffff' width='98%'>"
		OutStr=OutStr&"<td colspan='6'>&nbsp;<font color='red'>"&str_Comm_ListEmpty&"</font></td>"
		OutStr=OutStr&"</tr></table>"
		Response.write OutStr
		Response.flush
		Response.End 
	Else
		TopicList=EA_M_DBO.Get_Group_ForMemberList(PostId,Page,20)
		ForTotal = Ubound(TopicList,2)

		For i=0 To ForTotal
		%>
    <tr align='center' onmouseover=this.style.backgroundColor='E4E8EF' onmouseout=this.style.backgroundColor='' bgcolor="ffffff">
      <td><input type='checkbox' name='userid' value='<%=TopicList(4,i)%>'></td>
      <td><%=TopicList(0,i)%></td>
      <td><a href="mailto:<%=TopicList(1,i)%>"><%=TopicList(1,i)%></a></td>
      <td><%=TopicList(2,i)%></td>
      <td><%=TopicList(3,i)%></td>
    </tr>
    <%
		Next
	End If%>
    <tr bgcolor="ffffff">
      <td colspan="2" align="left" height="25">&nbsp;<input name="chkall" type="checkbox" value="on" onclick="CheckAll(this.form)"><%=str_Comm_SelectAll%>&nbsp;<%=str_Member_MoveTo%>:<select name="dest" size="1">
          <%
	Temp=EA_M_DBO.Get_Group_List
	If IsArray(Temp) Then
		ForTotal = UBound(Temp,2)

		For i=0 To ForTotal%>
          <option value="<%=Temp(0,i)%>"><%=Temp(1,i)%></option>
          <%
		Next
	End If%>
        </select>&nbsp;<input type="submit" name="Submit" value="<%=str_Comm_Submit_Button%>"></td>
      <td colspan="3" align="right" height="25"><%Response.Write EA_Manager.PageList(20,Count,Page,FieldName,FieldValue)%></td>
    </tr>
  </form>
</table>
<%
Call EA_Pub.Close_Obj
Set EA_Pub=Nothing

Function Move
	Call EA_Pub.Chk_Post
	Dim ForTotal
	
	If Len(Request.Form ("userid"))=0 then 
		ErrMsg="请选择用户！"
		Call EA_Manager.Error(1)
	End If

	Dim UserArray,i,Dest
	UserArray=EA_Pub.SafeRequest(2,"userid",1,"",0)
	UserArray=Replace(UserArray," ","")
	UserArray=Split(UserArray,",")
	Dest=EA_Pub.SafeRequest(2,"dest",0,0,0)
	
	If Dest<>0 Then 
		ForTotal = Ubound(UserArray)

		For i=0 To ForTotal
			EA_M_DBO.Get_Group_ChanngeMemberGroup Dest,UserArray(i)
		Next
		
		EA_M_DBO.Set_Group_MemberTotal "-"&Ubound(UserArray)+1,PostId
		EA_M_DBO.Set_Group_MemberTotal Ubound(UserArray)+1,Dest
	End If
End Function
%>
