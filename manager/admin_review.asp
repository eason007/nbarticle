<!--#Include File="../conn.asp" -->
<!--#Include File="comm/inc.asp" -->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：/Manager/Admin_Review.asp
'= 摘    要：后台-评论管理文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2006-07-27
'====================================================================

Call EA_Manager.Chk_IsMaster

If Not EA_Manager.Chk_Power(Admin_Power,"14") Then 
	ErrMsg=str_Comm_NotAccess
	Call EA_Manager.Error(1)
End If

Dim Atcion
Dim ForTotal

Atcion=Request.Form ("action")

Select Case LCase(Atcion)
Case "del"
	Call Del
Case "pass"
	Call Pass
Case "npass"
	Call NoPass
Case Else
	Call Main
End Select
Call EA_Pub.Close_Obj
Set EA_Pub=Nothing
	
Sub Main
	Dim Count,PageCount,i
	Dim WStr
	Dim TopicList
	Dim ListName(6),ListValue()
	
	Page = EA_Pub.SafeRequest(2,"nowPage",0,1,0)
	WStr = EA_Pub.SafeRequest(2,"w",1,"",0)
	
	Select Case WStr
	Case "1d"
		If iDataBaseType=0 Then 
			WStr=" Where DateDiff('d',a.AddDate,Now())=0"
		Else
			WStr=" Where DateDiff(d,a.AddDate,GetDate())=0"
		End If
	Case "1w"
		If iDataBaseType=0 Then 
			WStr=" Where DateDiff('ww',a.AddDate,Now())=0"
		Else
			WStr=" Where DateDiff(ww,a.AddDate,GetDate())=0"
		End If
	Case "1m"
		If iDataBaseType=0 Then 
			WStr=" Where DateDiff('m',a.AddDate,Now())=0"
		Else
			WStr=" Where DateDiff(m,a.AddDate,GetDate())=0"
		End If
	Case "pass"
		WStr=" Where a.IsPass="&EA_M_DBO.TrueValue
	Case "npass"
		WStr=" Where a.IsPass=0"
	Case Else
		WStr=""
	End Select

	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_Review_Help",str_Review_Help)

	Call EA_M_XML.AppElements("Language_Comm_AllColumn",str_Comm_AllColumn)
	Call EA_M_XML.AppElements("Language_Comm_Today",str_Comm_Today)
	Call EA_M_XML.AppElements("Language_Comm_Week",str_Comm_Week)
	Call EA_M_XML.AppElements("Language_Comm_Month",str_Comm_Month)
	Call EA_M_XML.AppElements("Language_Comm_State_NoPass",str_Comm_State_NoPass)
	Call EA_M_XML.AppElements("Language_Comm_State_Pass",str_Comm_State_Pass)

	Call EA_M_XML.AppElements("Language_Review_User",str_Review_User)
	Call EA_M_XML.AppElements("Language_Review_UnderArticle",str_Review_UnderArticle)
	Call EA_M_XML.AppElements("Language_Review_Content",str_Review_Content)
	Call EA_M_XML.AppElements("Language_Review_AddTime",str_Review_AddTime)
	Call EA_M_XML.AppElements("Language_Review_State",str_Review_State)

	Call EA_M_XML.AppElements("Comm_Del_Operation",str_Comm_Del_Operation)
	Call EA_M_XML.AppElements("Comm_State_Pass",str_Comm_State_Pass)
	Call EA_M_XML.AppElements("Comm_State_NoPass",str_Comm_State_NoPass)

	Call EA_M_XML.AppElements("Language_Comm_Bar_Operation",str_Comm_Bar_Operation)

	SQL="Select Count([Id]) From [NB_Review] a"&WStr
	Count=EA_M_DBO.DB_Query(SQL)(0, 0)
	If Count>0 Then 
		If Rs.State=1 Then Rs.Close
		'0=Id,1=Ip,2=Content,3=UserName,4=AddDate,5=IsPass,6=Article_Title,7=Article_Id,8=Article_AddDate
		If iDataBaseType=0 Then 
			SQL="Select a.[Id],a.Ip,a.Content,IIF(a.UserId=0,a.UserName,'[会员]'+a.UserName),a.AddDate,a.IsPass,IIF(b.Title Is Null,'<font color=800000>文章已被删除</font>',b.Title),a.ArticleId,IIF(b.AddDate Is Null,Now(),b.AddDate) From [NB_Review] a Left Join [NB_Content] b On a.ArticleId=b.[Id]"&WStr&" Order By a.[Id] Desc"
		Else
			SQL="Select a.[Id],a.Ip,a.Content,Case a.UserId When 0 Then a.UserName Else '[会员]'+a.UserName End,a.AddDate,a.IsPass,IsNull(b.Title,'<font color=800000>文章已被删除</font>'),a.ArticleId,IsNull(b.AddDate,GetDate()) From [NB_Review] a Left Join [NB_Content] b On a.ArticleId=b.[Id]"&WStr&" Order By a.[Id] Desc"
		End If
		Rs.Open SQL,Conn,1,1
		If Not rs.eof And Not rs.bof Then 
			Rs.AbsolutePosition=Rs.AbsolutePosition+((Abs(Page)-1)*10)
			TopicList=Rs.GetRows(10)
		End If
		Rs.Close:Set rs=Nothing

		ListName(0) = "checkbox"
		ListName(1) = "ID"
		ListName(2) = "Content"
		ListName(3) = "User"
		ListName(4) = "Article"
		ListName(5) = "State"
		ListName(6) = "AddTime"
		ForTotal = Ubound(TopicList,2)
	
	    For i=0 To ForTotal
			ReDim Preserve ListValue(8,i)

			ListValue(0,i) = "checkbox"
			ListValue(1,i) = TopicList(0,i)
			ListValue(2,i) = EA_Pub.Full_HTMLFilter(TopicList(2,i))
			ListValue(3,i) = EA_Pub.Full_HTMLFilter(TopicList(3,i)) & "(" & TopicList(1,i) & ")"
			ListValue(4,i) = "<a href='" & EA_Pub.Cov_ArticlePath(TopicList(7,i),TopicList(8,i),EA_Pub.SysInfo(18)) & "' target='_blank'>" & str_Review_ViewArticle & "</a>"
			If TopicList(5,i) Then
				ListValue(5,i) = "<strong>√</strong>"
			Else
				ListValue(5,i) = "<font color=""red""><strong>×</strong></font>"
			End If
			ListValue(6,i) = TopicList(4,i)
		Next

		Page = EA_M_XML.make(ListName,ListValue,Count)
	Else
		Page = EA_M_XML.make("","",0)
	End If

	Call EA_M_XML.Out(Page)
End Sub

Sub Del
	Dim IDs
	Dim i,Tmp

	IDs = Split(Request.Form("ID"), ",")
	ForTotal = UBound(IDs)

	For i = 0 To ForTotal
		Tmp = EA_Pub.SafeRequest(5,IDs(i),0,0,0)

		EA_M_DBO.Set_Review_Delete Tmp
	Next

	Application.Lock 
	Application(sCacheName&"IsFlush")=1
	Application.UnLock 
	
	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing
	
	Response.Write "1"
	Response.End
End Sub

Sub Pass
	Dim IDs
	Dim i,Tmp

	IDs = Split(Request.Form("ID"), ",")
	ForTotal = UBound(IDs)

	For i = 0 To ForTotal
		Tmp = EA_Pub.SafeRequest(5,IDs(i),0,0,0)

		EA_M_DBO.Set_Review_Pass 1,Tmp
	Next

	Application.Lock 
	Application(sCacheName&"IsFlush")=1
	Application.UnLock 
	
	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing
	
	Response.Write "1"
	Response.End
End Sub

Sub NoPass
	Dim IDs
	Dim i,Tmp

	IDs = Split(Request.Form("ID"), ",")
	ForTotal = UBound(IDs)

	For i = 0 To ForTotal
		Tmp = EA_Pub.SafeRequest(5,IDs(i),0,0,0)

		EA_M_DBO.Set_Review_Pass 0,Tmp
	Next

	Application.Lock 
	Application(sCacheName&"IsFlush")=1
	Application.UnLock 
	
	Call EA_Pub.Close_Obj
	Set EA_Pub=Nothing
	
	Response.Write "1"
	Response.End
End Sub
%>