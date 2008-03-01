<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：cls_DBOperation.asp
'= 摘    要：数据库操作类文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-02-28
'====================================================================

Class cls_DBOperation
	Private Conn
	Private Rs
	Private SQL
	Private Debug

	Public TrueValue
	Public ExecuteTotal, QueryTotal
	Public T_SQL_List
	
	Private Sub Class_Initialize()
		ExecuteTotal	= 0
		QueryTotal		= 0
		T_SQL_List		= ""
		Debug			= True

		Select Case iDataBaseType
		Case 0
			TrueValue = "-1"
		Case 1,2
			TrueValue = "1"
		End Select
	End Sub

	Private Sub ConnectionDatabase
		On Error Resume Next

		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open ConnStr
		If Err Then	
			Response.Clear
			Err.Clear

			CloseDataBase

			Response.Write "对不起，数据连接错误！如果第一次使用，请先运行setup.asp进行系统配置。"
			Response.End
		End If
	End Sub

	Public Sub Close()
		If Rs.State = 1 Then Rs.Close
		Set Rs=Nothing

		If Conn.State=1 Then Conn.Close
		Set Conn = Nothing
	End Sub
	
	Public Function Get_Nav_List(iStepNum,sCode)
		Dim i
		
		SQL="Select [Id],Title From [NB_Column] Where 1=2"
		For i=1 To iStepNum
			SQL=SQL&" Or Code='"&Left(sCode,i*4)&"'"
		Next
		SQL=SQL&" Order By Code"
		
		Get_Nav_List=DB_Query(SQL)
	End Function
	
	Public Function Get_InsideLink_ByColumn(ColumnId)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_InsideLink "&ColumnId
		Case 1
			SQL="SELECT Word, Link"
			SQL=SQL&" FROM NB_Link"
			SQL=SQL&" WHERE (ColumnId=0 Or ColumnId="&ColumnId&")"
		Case 2
			SQL="Exec sp_EliteArticle_InsideLink_ByColumn_Select "&ColumnId
		End Select
		
		Get_InsideLink_ByColumn=DB_Query(SQL)
	End Function
	
	Public Function Get_Group_Setting(iGroupId)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_GroupSetting "&iGroupId
		Case 1
			SQL="SELECT GroupName, IsLogin, Setting"
			SQL=SQL&" FROM NB_UserGroup"
			SQL=SQL&" WHERE [Id]="&iGroupId
		Case 2
			SQL="Exec sp_EliteArticle_UserGroup_Info_Select "&iGroupId
		End Select
		
		Get_Group_Setting=DB_Query(SQL)
	End Function
	
	Public Function Get_Ip_LockInfo(sIp)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_Chk_LockIp '"&sIp&"'"
		Case 1
			SQL="SELECT TOP 1 [Id]"
			SQL=SQL&" FROM NB_Ip"
			SQL=SQL&" WHERE Head_Ip<='"&sIp&"' And Foot_Ip>='"&sIp&"' And DateDiff(d,GetDate(),OverTime)>=0"
		Case 2
			SQL="Exec sp_EliteArticle_Ip_ChkLock_Select '"&sIp&"'"
		End Select
		
		Get_Ip_LockInfo=DB_Query(SQL)
	End Function
	
	Public Function Get_AdSense(iAdSense)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_Manager_AdSenseInfo "&iAdSense
		Case 1
			SQL="SELECT Title, Content"
			SQL=SQL&" FROM [NB_AdSense]"
			SQL=SQL&" WHERE [Id]="&iAdSense
		Case 2
			SQL="Exec sp_EliteArticle_AdSense_Info_Select "&iAdSense
		End Select
		
		Get_AdSense=DB_Query(SQL)
	End Function
	
	Public Function Get_Template_Info(iTemplateId, iTemplateType)
		SQL="SELECT TOP 1 Code, ThemesID, ID"
		SQL=SQL&" FROM NB_Module"
		If iTemplateId > 0 Then
			SQL=SQL&" WHERE [Id]="&iTemplateId
		Else 
			SQL=SQL&" WHERE ThemesID=(SELECT TOP 1 ID FROM NB_Themes WHERE IsDefault = 1) AND [Type] = "&iTemplateType
		End If
		
		Get_Template_Info=DB_Query(SQL)
	End Function

	Public Function Get_Theme_Name(iModuleID)
		SQL="SELECT Title"
		SQL=SQL&" FROM NB_Themes"
		SQL=SQL&" WHERE ID=(SELECT ThemesID FROM NB_Module WHERE ID=" & iModuleID & ")"
		
		Get_Theme_Name=DB_Query(SQL)
	End Function

	Public Function Get_System_Info()
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_Load_Config"
		Case 1
			SQL="SELECT TOP 1 ColumnNum, TopicNum, MangerTopicNum, RegUser, ReviewNum, Info, Source, BadWord"
			SQL=SQL&" FROM NB_System"
		Case 2
			SQL="Exec sp_EliteArticle_System_LoadConfig_Select"
		End Select

		Get_System_Info=DB_Query(SQL)
	End Function

	Public Sub Set_Vote_SaveVoted(iVoteId,sVoteNum)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_UpDate_SaveVote '"&sVoteNum&"',"&iVoteId
		Case 1
			SQL="UPDATE NB_VOTE SET VoteNum = '"&sVoteNum&"', VoteTotal = VoteTotal+1"
			SQL=SQL&" WHERE [Id]="&iVoteId
		Case 2
			SQL="Exec sp_EliteArticle_Vote_SaveVoted_UpDate"
			SQL=SQL&" @Vote_Num='"&sVoteNum&"'"
			SQL=SQL&",@Vote_Id="&iVoteId
		End Select

		DB_Execute SQL
	End Sub

	Public Function Get_Vote_Info(iVoteId)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_VoteInfo "&iVoteId
		Case 1
			SQL="SELECT Id, TITLE, VOTETEXT, VoteNum, TYPE, LOCK"
			SQL=SQL&" FROM NB_VOTE"
			SQL=SQL&" WHERE Id="&iVoteId
		Case 2
			SQL="Exec sp_EliteArticle_Vote_Info_Select"
			SQL=SQL&" @Vote_Id="&iVoteId
		End Select

		Get_Vote_Info=DB_Query(SQL)
	End Function

	Public Function Get_Rss_List(iTopNum,iTypes,sAuthorName,iColumnId)
		SQL="Select Top "&iTopNum&" a.Id,a.Title,AddDate,Content From [NB_Content] a"
		SQL=SQL & " INNER JOIN [NB_Column] b ON b.Id=a.ColumnId"
		SQL=SQL & " Where IsPass="&TrueValue&" And IsDel=0 AND ListPower=0 AND IsHide=0"

		Select Case iTypes
		Case "2"
			SQL=SQL&" And IsTop="&TrueValue
		Case "3"
			SQL=SQL&" And IsImg="&TrueValue
		Case "4"
			SQL=SQL&" And Author Like '"&sAuthorName&"'"
		End Select
		If iColumnId<>"0" And IsNumeric(iColumnId) And CLng(iColumnId)>0 Then SQL=SQL&" And ColumnId="&iColumnId
		SQL=SQL&" Order By TrueTime Desc"
		
		Get_Rss_List=DB_Query(SQL)
	End Function
	
	Private Function Get_Column_ChkIsReview(iArticleId)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_ColumnIsReview "&iArticleId
		Case 1
			SQL="SELECT IsReview"
			SQL=SQL&" FROM NB_Content AS a LEFT JOIN NB_Column AS b ON a.ColumnId=b.Id"
			SQL=SQL&" WHERE a.Id="&iArticleId
		Case 2
			SQL="Exec sp_EliteArticle_Column_IsReviewByArticleId_Select"
			SQL=SQL&" @Article_Id="&iArticleId
		End Select
		
		Get_Column_ChkIsReview=DB_Query(SQL)
	End Function

	Public Function Set_Review_Insert(iArticleId,sRUserId,sRUserName,sRContent,sIp,iRState)
		Dim Temp
		Dim Flag

		Flag=0

		Temp=Get_Column_ChkIsReview(iArticleId)
		If IsArray(Temp) Then 
			If Temp(0,0)=0 Then 
				ErrMsg="管理员已设置该栏目下的文章不允许发表评论，请稍后重试。"
				Flag=1
			Else
				rs.Open "NB_Review",Conn,2,2
				rs.AddNew
				rs("ArticleId")=iArticleId
				rs("UserId")=sRUserId
				rs("UserName")=sRUserName
				rs("Content")=sRContent
				rs("IP")=sIp
				rs("IsPass")=iRState
				rs.update
				Rs.Close:Set Rs=Nothing
				
				If iDataBaseType<>2 Then
					SQL="Update [NB_Content] Set CommentNum=CommentNum+1"
					SQL=SQL&" Where [Id]="&iArticleId
					DB_Execute SQL
					If iRState=1 Then 
						Set_System_ReviewTotal 1

						SQL="Update [NB_Content] Set LastComment='"&Left(sRContent,25)&"'"
						SQL=SQL&" Where [Id]="&iArticleId
						DB_Execute SQL
					End If
				End If
				
				Session("lastpost")=Now()

				ErrMsg="您已经成功发布您的评论,请返回刷新页面!"
				Flag=0
			End If
		Else
			ErrMsg="指定错误的参数。"
			Flag=1
		End If

		Set_Review_Insert=Flag
	End Function

	Public Function Get_Review_List(iArticleId)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_ReviewList "&iArticleId
		Case 1
			SQL="SELECT Case UserId When 0 Then '游客' Else '[会员]'+UserName End As [UserName], AddDate, Content"
			SQL=SQL&" FROM NB_Review"
			SQL=SQL&" WHERE ArticleId="&iArticleId&" And IsPass=1"
			SQL=SQL&" ORDER BY [Id] DESC"
		Case 2
			SQL="Exec sp_EliteArticle_Review_ListByArticleId_Select"
			SQL=SQL&" @Article_Id="&iArticleId
		End Select
		
		Get_Review_List=DB_Query(SQL)
	End Function
	
	Public Function Get_Review_NewList(iTop,iContentLen)
		Select Case iDataBaseType
		Case 0
			SQL="Select Top "&iTop&" ArticleId,Left(Content,"&iContentLen&"),IIF(UserId=0,'游客','[会员]'+UserName),AddDate"
			SQL=SQL&" From [NB_Review] Where IsPass=-1 Order By Id Desc"
		Case 1,2
			SQL="Select Top "&iTop&" ArticleId,Left(Content,"&iContentLen&"),Case UserId When 0 Then '游客' Else '[会员]'+UserName End As [UserName],AddDate"
			SQL=SQL&" From [NB_Review]"
			SQL=SQL&" Where IsPass=1"
			SQL=SQL&" Order By Id Desc"
		End Select
		
		Get_Review_NewList=DB_Query(SQL)
	End Function

'*******************************************************************
'article
	Public Function Set_Article_Insert(iArticleId,vArticleInfo)
		On Error Resume Next
		Dim Flag
		Flag=0

		If iArticleId<>0 Then
			Sql="Select * From [NB_Content] Where [Id]="&iArticleId
			rs.Open Sql,Conn,2,2
		Else
			rs.Open "SELECT * FROM [NB_Content] WHERE 1=0",Conn,2,2
			rs.AddNew
		End If
			rs("title")=vArticleInfo(0)
			rs("author")=vArticleInfo(14)
			rs("authorid")=vArticleInfo(13)
			rs("Content")=vArticleInfo(1)
			rs("KeyWord")=vArticleInfo(2)
			rs("ColumnId")=vArticleInfo(3)
			rs("ColumnName")=vArticleInfo(4)
			rs("ColumnCode")=vArticleInfo(5)
			rs("byte")=vArticleInfo(11)
			rs("isimg")=vArticleInfo(7)
			rs("img")=vArticleInfo(6)
			rs("IsDis")=vArticleInfo(17)
			rs("CutArticle")=vArticleInfo(15)
			If PostId=0 Then
				rs("AddDate")=vArticleInfo(16)
				rs("TrueTime")=vArticleInfo(18)
			End If
			rs("Source")=vArticleInfo(8)
			rs("SourceUrl")=vArticleInfo(9)
			rs("Summary")=vArticleInfo(10)
			rs("IsPass")=vArticleInfo(12)
			rs.update
			Rs.Close
			
			If iArticleId=0 Then 
				If vArticleInfo(12)=0 Then 
					Set_System_ManagerTopicTotal 1
					
					Set_Column_ManagerTopicTotal vArticleInfo(3),1
				Else
					Set_System_TopicTotal 1
					
					Set_Column_TopicTotal vArticleInfo(3),1
				End If
			End If

		If Err Then 
			Flag=-1
		ElseIf vArticleInfo(12)=0 Then 
			Flag=1
		Else
			Flag=0
		End If

		If Flag<>-1 Then Set_MemberAppearTotal vArticleInfo(13)

		Set_Article_Insert=Flag
	End Function
	
	Public Function Get_Article_List(iTop,iColumnId,iArticleType,iIsIncludeChildColumn)
		SQL="SELECT TOP "&iTop&" [ID],COLUMNID,COLUMNNAME,TITLE,TCOLOR,AddDate,IsImg,IsTop,Img,Author,Summary"
		SQL=SQL&" FROM [NB_Content]"
		SQL=SQL&" WHERE ISDEL=0 AND ISPASS="&TrueValue

		If iColumnId<>"0" Then 
			If iIsIncludeChildColumn="1" Then
				SQL=SQL&" And ColumnCode Like (Select Code From [NB_Column] Where [Id]="&iColumnId&")+'%'"
			Else
				SQL=SQL&" And ColumnId="&iColumnId
			End If
		End If

		Select Case iArticleType
		Case "1"
			SQL=SQL&" And IsTop="&TrueValue
		Case "2"
			SQL=SQL&" And IsDis="&TrueValue
		Case "3"
			SQL=SQL&" And IsImg="&TrueValue
		End Select

		If iArticleType="4" Then 
			SQL=SQL&" Order By ViewNum Desc,TrueTime Desc"
		Else
			SQL=SQL&" Order By TrueTime Desc"
		End If
		
		Get_Article_List=DB_Query(SQL)
	End Function

	Public Function Get_Article_Info(iArticleId,iIsUpData)
	'0=ColumnId,1=ColumnCode,2=ColumnName,3=Title,4=Summary,5=Content,6=ViewNum,7=AuthorId,8=Author,9=CommentNum,10=IsOut
	'11=OutUrl,12=[KeyWord],13=AddDate,14=CutArticle,15=Source,16=SourceUrl,17=TColor,18=Img,19=IsTop,20=IsPass
	'21=IsDel,22=ListPower,23=IsHide,24=Article_TempId,25=TrueTime
		Select Case iDataBaseType
		Case 0
			SQL="vi_Select_ArticleInfo "&iArticleId
		Case 1
			SQL="SELECT ColumnId, ColumnCode, ColumnName, a.Title, Summary, Content, a.ViewNum, AuthorId, Author, CommentNum, a.IsOut, a.OutUrl, [KeyWord], AddDate, CutArticle, Source, SourceUrl, TColor, Img, a.IsTop, IsPass, IsDel, b.ListPower, b.IsHide, b.Article_TempId,TrueTime"
			SQL=SQL&" FROM NB_Content AS a INNER JOIN NB_Column AS b ON a.ColumnId=b.Id"
			SQL=SQL&" WHERE a.Id="&iArticleId
		Case 2
			SQL="Exec sp_EliteArticle_Article_Info_Select"
			SQL=SQL&" @Article_Id="&iArticleId
			SQL=SQL&",@IsUpData="&iIsUpData
		End Select
		
		Get_Article_Info=DB_Query(SQL)
	End Function

	Public Function Get_Article_CorrList(sWSQL,iArticleId,iColumnId,iTopNum)
		SQL="SELECT TOP " & iTopNum & " [ID],COLUMNID,COLUMNNAME,TITLE,TCOLOR,AddDate,IsImg,IsTop,Img,Author,Summary"
		SQL=SQL&" FROM [NB_CONTENT]"
		SQL=SQL&" WHERE ISPass="&TrueValue&" And ID<>"&iArticleId&" And ("&sWSQL&"1=0) And IsDel=0 AND COLUMNID="&iColumnId
		SQL=SQL&" ORDER BY AddDate DESC"
		
		Get_Article_CorrList=DB_Query(SQL)
	End Function

	Public Function Get_Article_FirstArticle(iColumnId,iTrueTime,iArticleId)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_FirstArticle "&iColumnId&","&iTrueTime&","&iArticleId
		Case 1,2
			SQL="SELECT [Id], Title, TColor, AddDate"
			SQL=SQL&" FROM NB_Content"
			SQL=SQL&" WHERE ColumnId="&iColumnId&" And IsPass=1 And IsDel=0 And TrueTime>"&iTrueTime&" And [ID]<>"&iArticleId
			SQL=SQL&" ORDER BY TrueTime"
		End Select
		
		Get_Article_FirstArticle=DB_Query(SQL)
	End Function

	Public Function Get_Article_NextArticle(iColumnId,iTrueTime,iArticleId)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_NextArticle "&iColumnId&","&iTrueTime&","&iArticleId
		Case 1,2
			SQL="SELECT [Id], Title, TColor, AddDate"
			SQL=SQL&" FROM NB_Content"
			SQL=SQL&" WHERE ColumnId="&iColumnId&" And IsPass=1 And IsDel=0 And TrueTime<="&iTrueTime&" And [ID]<>"&iArticleId
			SQL=SQL&" ORDER BY TrueTime DESC"
		End Select
		
		Get_Article_NextArticle=DB_Query(SQL)
	End Function

	Public Sub Set_Article_ViewNum_UpDate(iArticleId)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_UpDate_ArticleViewNum "&iArticleId
		Case 1
			SQL="UPDATE NB_Content SET ViewNum = ViewNum+1"
			SQL=SQL&" WHERE [Id]="&iArticleId
		End Select

		DB_Execute SQL
	End Sub

	Public Sub Set_Article_CommentNum_UpDate(iArticleID)
		Dim Temp

		Select Case iDataBaseType
		Case 0
			SQL  = "SELECT COUNT([ID]) FROM [NB_Review] WHERE ArticleId=" & iArticleID & " AND IsPass=" & TrueValue
			Temp = DB_Query(SQL)

			SQL = "UPDATE [NB_Content] SET CommentNum=" & Temp(0,0) & " WHERE [ID]=" & iArticleID
			DB_Execute SQL
		Case 1,2
			SQL = "UPDATE [NB_Content] SET CommentNum=(SELECT COUNT([ID])"
			SQL = SQL & " FROM [NB_Review]"
			SQL = SQL & " WHERE ArticleId=" & iArticleID & " AND IsPass=" & TrueValue & ")"
			SQL = SQL & " WHERE [ID]=" & iArticleID
			DB_Execute SQL
		End Select
	End Sub

	Public Function Get_Article_ByColumnId(iColumnId,iPageNum,iPageSize)
		Dim Temp
		
		Select Case iDataBaseType
		'Case 0
		'	SQL="Exec vi_Select_ArticleListById "&iColumnId
		'	Temp=DB_CutPageQuery(SQL,iPageNum,iPageSize)
		Case 0,1
			SQL="SELECT [Id], TColor, Title, AddDate, CommentNum, Summary, LastComment, ViewNum, IsImg, Img, IsTop, Author, AuthorId, [KeyWord]"
			SQL=SQL&" FROM NB_Content"
			SQL=SQL&" WHERE ColumnId="&iColumnId&" And IsPass=" & TrueValue & " And IsDel=0"
			SQL=SQL&" ORDER BY TrueTime DESC"
			Temp=DB_CutPageQuery(SQL,iPageNum,iPageSize)
		Case 2
			SQL="Exec sp_EliteArticle_Article_ListById_Select"
			SQL=SQL&" @ColumnId="&iColumnId
			SQL=SQL&",@List_PageNum="&iPageNum
			SQL=SQL&",@List_PageSize="&iPageSize
			Temp=DB_Query(SQL)
		End Select
		
		Get_Article_ByColumnId=Temp
	End Function
'-------------------------------------------------------------------
	Public Function Get_Friend_List(iTop,iColumnId,iStyle)
		SQL="SELECT Top "&iTop&" LINKNAME,LINKURL,LINKIMGPATH,LINKINFO"
		SQL=SQL&" FROM [NB_FriendLink]"
		SQL=SQL&" Where Style="&iStyle&" And State="&TrueValue
		If iColumnId >= 0 Then SQL = SQL & " AND ColumnId=" & iColumnId
		SQL=SQL&" Order By ColumnId ASC, OrderNum Desc,Id ASC"
		
		Get_Friend_List=DB_Query(SQL)
	End Function

	Public Function Get_FriendList_All(iStyle)
		Select Case iDataBaseType
		Case 0, 1
			SQL="SELECT a.ColumnId, IsNull(b.Title,'首页'), LinkUrl, LinkInfo, LinkImgPath, LinkName, a.Style"
			SQL=SQL&" FROM NB_FriendLink AS a LEFT JOIN NB_Column AS b ON a.ColumnId=b.Id"
			SQL=SQL&" WHERE a.State = " & TrueValue & " And Style = " & iStyle
			SQL=SQL&" ORDER BY a.ColumnId, a.Style DESC , a.OrderNum DESC"
		Case 2
			SQL="Exec sp_EliteArticle_FriendLink_All_Select"
		End Select

		Get_FriendList_All=DB_Query(SQL)
	End Function

	Public Sub Set_FriendList_Insert(LinkName,LinkImg,LinkUrl,LinkInfo,ColumnId,Style,OrderNum,State)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Insert_AppFriend '"&LinkName&"','"&LinkImg&"','"&LinkUrl&"','"&LinkInfo&"',"&ColumnId&","&Style&","&OrderNum&","&State
		Case 1
			SQL="INSERT INTO NB_FriendLink ( LinkName, LinkImgPath, LinkUrl, LinkInfo, ColumnId, Style, OrderNum, State )"
			SQL=SQL&" VALUES ('"&LinkName&"', '"&LinkImg&"', '"&LinkUrl&"', '"&LinkInfo&"', "&ColumnId&", "&Style&", "&OrderNum&", "&State&")"
		Case 2
			SQL="Exec sp_EliteArticle_FriendLink_Insert"
			SQL=SQL&" @FriendLink_Name='"&LinkName&"'"
			SQL=SQL&",@FriendLink_ImgPath='"&LinkImg&"'"
			SQL=SQL&",@FriendLink_Url='"&LinkUrl&"'"
			SQL=SQL&",@FriendLink_Info='"&LinkInfo&"'"
			SQL=SQL&",@FriendLink_ColumnId="&ColumnId
			SQL=SQL&",@FriendLink_Style="&Style
			SQL=SQL&",@FriendLink_OrderNum="&OrderNum
			SQL=SQL&",@FriendLink_State="&State
		End Select

		DB_Execute SQL
	End Sub

'*******************************************************************
	Public Function Get_Column_Info(iColumnId)
	'0=title,1=code,2=info,3=coumnnum,4=managernum,5=viewnum,6=isout,7=outurl,8=styleid,9=list_tempid,10=article_tempid
	'11=0,12=listpower,13=ishide,14=isreview,15=ispost,16=istop,17=pagesize
		Select Case iDataBaseType
		Case 0, 1
			SQL="SELECT Title, Code, Info, CountNum, MangerNum, ViewNum, IsOut, OutUrl, StyleId, List_TempId, Article_TempId, 0, ListPower, IsHide, IsReview, IsPost, IsTop, PageSize"
			SQL=SQL&" FROM NB_Column"
			SQL=SQL&" WHERE [Id]="&iColumnId
		Case 2
			SQL="Exec sp_EliteArticle_Column_Info_Select"
			SQL=SQL&" @ColumnId="&iColumnId
		End Select
		
		Get_Column_Info=DB_Query(SQL)
	End Function

	Public Function Get_Column_ChildList(sMainCode)
		Dim Temp
		
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_ColumnChild '"&sMainCode&"'"
		Case 1
			Temp=Len(sMainCode)
			
			SQL="SELECT [ID], Title, [CountNum], [ViewNum]"
			SQL=SQL&" FROM NB_Column"
			SQL=SQL&" WHERE Left(Code,"&Temp&")='"&sMainCode&"' And Len(Code)="&Temp&"+4"
		Case 2
			SQL="Exec sp_EliteArticle_Column_ChildList_Select"
			SQL=SQL&" @Main_Code='"&sMainCode&"'"
		End Select
		
		Get_Column_ChildList=DB_Query(SQL)
	End Function

	Public Function Get_Column_List()
		Select Case iDataBaseType
		Case 0, 1
			SQL="SELECT [Id], Title, Code, Info, CountNum, MangerNum, '', IsTop"
			SQL=SQL&" FROM NB_Column"
			SQL=SQL&" ORDER BY Code"
		Case 2
			SQL="Exec sp_EliteArticle_Column_List_Select"
		End Select
		
		Get_Column_List=DB_Query(SQL)
	End Function
'-------------------------------------------------------------------

'*******************************************************************
	Public Function Get_FlorilegiumStat(s_AName,i_AId)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_FlorilegiumStat "&i_AId&",'"&s_AName&"'"
		Case 1
			SQL="SELECT Count([Id])"
			SQL=SQL&" FROM NB_Content"
			SQL=SQL&" WHERE AuthorId="&i_AId&" And Author='"&s_AName&"' And IsPass=1 And IsDel=0"
		Case 2
			SQL="Exec sp_EliteArticle_Florilegium_Total_Select"
			SQL=SQL&" @Florilegium_AuthorName='"&s_AName&"'"
			SQL=SQL&",@Florilegium_AuthorId="&i_AId
		End Select
		
		Get_FlorilegiumStat=DB_Query(SQL)
	End Function
'-------------------------------------------------------------------
	
'*******************************************************************
'placard
	Public Function Get_PlacardTopList(iTop)
		Select Case iDataBaseType
		Case 0
			SQL="SELECT Top "&iTop&" [Id], Title, AddTime, OverTime, Content FROM NB_Placard Where OverTime>=Now() ORDER BY Id DESC"
		Case 1, 2
			SQL="SELECT Top "&iTop&" [Id], Title, AddTime, OverTime, Content FROM NB_Placard Where OverTime>=GetDate() ORDER BY Id DESC"
		End Select

		Get_PlacardTopList=DB_Query(SQL)
	End Function
	
	Public Function Get_PlacardInfo(iPlacardId)
		Select Case iDataBaseType
		Case 0, 1
			SQL="SELECT [Id], Title, AddTime, OverTime, Content"
			SQL=SQL&" FROM NB_Placard"
			SQL=SQL&" WHERE Id="&iPlacardId
		Case 2
			SQL="Exec sp_EliteArticle_Placard_Info_Select"
			SQL=SQL&" @Placard_Id="&iPlacardId
		End Select
		
		Get_PlacardInfo=DB_Query(SQL)
	End Function
'-------------------------------------------------------------------
	Public Function Get_MemberTopPostList()
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_MemberTopPostList"
		Case 1
			SQL="SELECT TOP 10 [Id], Reg_Name, PostTotal"
			SQL=SQL&" FROM NB_User"
			SQL=SQL&" ORDER BY PostTotal DESC"
		Case 2
			SQL="Exec sp_EliteArticle_Member_PostList_Select"
		End Select
		
		Get_MemberTopPostList=DB_Query(SQL)
	End Function

	Public Sub Set_MemberLoginKey(sIp,sKey,iMemberId)
		Select Case iDataBaseType
		Case 0
			SQL="vi_UpDate_MemberLogin '"&sIp&"','"&sKey&"',"&iMemberId
		Case 1
			SQL="UPDATE NB_User SET Login = [Login]+1, LastIp = '"&sIp&"', LasTime = GetDate(), Cookies = '"&sKey&"'"
			SQL=SQL&" WHERE [Id]="&iMemberId
		Case 2
			SQL="Exec sp_EliteArticle_Member_LoginKey_UpDate"
			SQL=SQL&" @LoginIp='"&sIp&"'"
			SQL=SQL&",@LoginKey='"&sKey&"'"
			SQL=SQL&",@LoginId='"&iMemberId
		End Select
		
		DB_Execute SQL
	End Sub

	Public Function Get_MemberLogin(iAccountName)
		Select Case iDataBaseType
		Case 0
			SQL="vi_Select_Chk_MemLogin '"&iAccountName&"'"
		Case 1
			SQL="SELECT Id, Reg_Pass, State, RegTime, User_Group"
			SQL=SQL&" FROM NB_User"
			SQL=SQL&" WHERE Reg_Name='"&iAccountName&"'"
		Case 2
			SQL="Exec sp_EliteArticle_Member_Login_Select"
			SQL=SQL&" @Login_Account='"&iAccountName&"'"
		End Select
		
		Get_MemberLogin=DB_Query(SQL)
	End Function

	Public Function Get_MemberLoginInfo(iAccountId)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_Member_Info "&iAccountId
		Case 1
			SQL="SELECT Id, Reg_Name, Sex, Email, RegTime, Login, [UserName], BirtDay, User_Group, State, HomePage, QQ, ICQ, MSN, Comefrom, Reg_Pass, Cookies"
			SQL=SQL&" FROM NB_User"
			SQL=SQL&" WHERE Id="&iAccountId
		Case 2
			SQL="Exec sp_EliteArticle_Member_Info_Select"
			SQL=SQL&" @Member_Id="&iAccountId
		End Select
		
		Get_MemberLoginInfo=DB_Query(SQL)
	End Function

'*******************************************************************
'myappear list
	Public Function Get_Member_AppearTotal(iAccountId)
		SQL="SELECT Count(Id)"
		SQL=SQL&" FROM NB_Content"
		SQL=SQL&" WHERE AuthorId="&iAccountId&" AND IsDel = 0"
		
		Get_Member_AppearTotal=DB_Query(SQL)
	End Function
	
	Public Function Get_MemberAppearList(iAccountId,iPageNum,iPageSize)
		Dim Temp

		'0=articleid,1=TColor,2=Title,3=ColumnName,4=ViewNum,5=CommentNum,6=AddDate,7=IsPass,8=ColumnId
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_Member_AppearList "&iAccountId
			Temp=DB_CutPageQuery(SQL,iPageNum,iPageSize)
		Case 1
			SQL="SELECT Id, TColor, Title, ColumnName, ViewNum, CommentNum, AddDate, IsPass, ColumnId"
			SQL=SQL&" FROM NB_Content"
			SQL=SQL&" WHERE AuthorId="&iAccountId&" And IsDel=0"
			SQL=SQL&" ORDER BY TrueTime DESC"
			Temp=DB_CutPageQuery(SQL,iPageNum,iPageSize)
		Case 2
			SQL="Exec sp_EliteArticle_Member_AppearList_Select"
			SQL=SQL&" @Member_Id="&iAccountId
			SQL=SQL&",@List_PageNum="&iPageNum
			SQL=SQL&",@List_PageSize="&iPageSize
			Temp=DB_Query(SQL)
		End Select
		
		Get_MemberAppearList=Temp
	End Function
	
	Public Sub Del_MemberAppear(iAppearId,iAccountId,iColumnId,iIsPass)
		'05/10/12
		'eason
		'only channge article status
		'begin
		SQL = "UPDATE [NB_Content] SET IsDel=1 WHERE [ID]=" & iAppearId & " AND AuthorId=" & iAccountId
		DB_Execute(SQL)

		If iIsPass=1 Then
			Set_System_TopicTotal -1
			Set_Column_TopicTotal iColumnId,-1

			Set_System_ManagerTopicTotal 1
			Set_Column_ManagerTopicTotal iColumnId,1
		End If
		'end
	End Sub

	Public Function Set_Member_Info(iMemberId,vMemberInfo,sMemberName)
		Dim Flag,Temp
		Dim RsCopy

		Set RsCopy=Server.CreateObject("adodb.recordSet")
		Flag=0

		If iMemberId = 0 And sMemberName <> "" Then
			SQL = "SELECT [ID] FROM [NB_User] Where Reg_Name = '" & sMemberName & "'"
			Temp= DB_Query(SQL)
			If IsArray(Temp) Then
				iMemberId = Temp(0,0)
			Else
				Set_Member_Info = 2
				Exit Function
			End If
		End If

		Select Case iDataBaseType
		Case 0
			If RsCopy.State=1 Then RsCopy.Close
			Sql="SELECT Reg_Pass,Email,[Sex],HomePage,QQ,ICQ,MSN,[UserName],Birtday,Comefrom FROM NB_User WHERE ID="&iMemberId
			RsCopy.Open SQL,Conn,1,3
			If Not RsCopy.eof And Not RsCopy.Bof Then
				If RsCopy("Reg_Pass")=vMemberInfo(0) Then
					If RsCopy("Email")<>vMemberInfo(1) Then
						If EA_Pub.SysInfo(8)="1" And Not EA_DBO.Get_MemberChkEMail(vMemberInfo(1),iMemberId) Then 
							Flag=1
						Else
							RsCopy("Email")=vMemberInfo(1)
						End If
					End If
					
					If Flag=0 Then
						RsCopy("Sex")=vMemberInfo(2)
						RsCopy("HomePage")=vMemberInfo(3)
						RsCopy("QQ")=vMemberInfo(4)
						RsCopy("ICQ")=vMemberInfo(5)
						RsCopy("MSN")=vMemberInfo(6)
						RsCopy("UserName")=vMemberInfo(7)
						RsCopy("Birtday")=vMemberInfo(8)
						RsCopy("Comefrom")=vMemberInfo(9)		
						RsCopy.Update
						Flag=0
					End If
				Else
					Flag=-1
				End If
			Else
				Flag=2
			End If
		Case 1
			SQL="Select Reg_Pass,Email From NB_User Where [Id]="&iMemberId
			Temp=DB_Query(SQL)
			If IsArray(Temp) Then 
				If Temp(0,0)=vMemberInfo(0) Then
					SQL="UpDate NB_User"
					SQL=SQL&" Set Sex="&vMemberInfo(2)
					SQL=SQL&",HomePage='"&vMemberInfo(3)&"'"
					SQL=SQL&",QQ="&vMemberInfo(4)
					SQL=SQL&",ICQ="&vMemberInfo(5)
					SQL=SQL&",MSN='"&vMemberInfo(6)&"'"
					SQL=SQL&",UserName='"&vMemberInfo(7)&"'"
					SQL=SQL&",Birtday='"&vMemberInfo(8)&"'"
					SQL=SQL&",Comefrom='"&vMemberInfo(9)&"'"
				
					If Temp(1,0)<>vMemberInfo(1) Then
						If EA_Pub.SysInfo(8)="1" And Not EA_DBO.Get_MemberChkEMail(vMemberInfo(1),iMemberId) Then 
							Flag=1
						Else
							SQL=SQL&",Email='"&vMemberInfo(1)&"'"
						End If
					End If
					
					If Flag=0 Then
						DB_Execute(SQL)
						Flag=0
					End If
				Else
					Flag=-1
				End If
			Else
				Flag=2
			End If
		Case 2
			If EA_Pub.SysInfo(8)="1" And Not EA_DBO.Get_MemberChkEMail(vMemberInfo(1),iMemberId) Then 
				Flag=1
			Else
				SQL="Exec sp_EliteArticle_Member_Info_UpDate"
				SQL=SQL&" @MemberId="&iMemberId
				SQL=SQL&",@RegPass='"&vMemberInfo(0)&"'"
				SQL=SQL&",@Email='"&vMemberInfo(1)
				SQL=SQL&",@Sex="&vMemberInfo(2)
				SQL=SQL&",@HomePage='"&vMemberInfo(3)&"'"
				SQL=SQL&",@QQ="&vMemberInfo(4)
				SQL=SQL&",@ICQ="&vMemberInfo(5)
				SQL=SQL&",@MSN='"&vMemberInfo(6)&"'"
				SQL=SQL&",@UserName='"&vMemberInfo(7)&"'"
				SQL=SQL&",@Birtday='"&vMemberInfo(8)&"'"
				SQL=SQL&",@Comefrom='"&vMemberInfo(9)&"'"
				
				Flag=DB_Query(SQL)(0,0)
			End If
		End Select

		Set_Member_Info=Flag
	End Function

	Public Function Set_Member_SafetyInfo(iMemberId,vMemberInfo,sMemberName)
		Dim Temp,Flag
		Flag=0

		If iMemberId = 0 And sMemberName <> "" Then
			SQL = "SELECT [ID] FROM [NB_User] Where Reg_Name = '" & sMemberName & "'"
			Temp= DB_Query(SQL)
			If IsArray(Temp) Then
				iMemberId = Temp(0,0)
			Else
				Set_Member_Info = -1
				Exit Function
			End If
		End If
		
		Select Case iDataBaseType
		Case 0
			Temp=Get_MemberInfo(iMemberId)
			If IsArray(Temp) Then 
				If Temp(15,0)=vMemberInfo(0) Then 
					SQL="Exec vi_UpDate_Member_SafeInfo '"&vMemberInfo(1)&"','"&vMemberInfo(2)&"','"&vMemberInfo(3)&"',"&iMemberId
					
					DB_Execute SQL
					Flag=0
				Else
					Flag=1
				End If
			Else
				Flag=-1
			End If
		Case 1
			SQL="Select [ID] From [NB_User] Where [Id]="&iMemberId&" And Reg_Pass='"&vMemberInfo(0)&"'"
			Temp=DB_Query(SQL)
			If IsArray(Temp) Then 
				SQL="UPDATE NB_User SET Reg_Pass = '"&vMemberInfo(1)&"', Question = '"&vMemberInfo(2)&"', Answer = '"&vMemberInfo(3)&"'"
				SQL=SQL&" WHERE [Id]="&iMemberId
				
				DB_Execute SQL
				Flag=0
			Else
				Flag=1
			End If
		Case 2
			SQL="Exec sp_EliteArticle_Member_SafetyInfo_UpDate"
			SQL=SQL&" @Member_Id="&iMemberId
			SQL=SQL&",@Old_Password='"&vMemberInfo(0)&"'"
			SQL=SQL&",@New_Password='"&vMemberInfo(1)&"'"
			SQL=SQL&",@Question='"&vMemberInfo(2)&"'"
			SQL=SQL&",@Answer='"&vMemberInfo(3)&"'"
			
			Flag=DB_Query(SQL)(0,0)
		End Select

		Set_Member_SafetyInfo=Flag
	End Function
'-------------------------------------------------------------------

'*******************************************************************
'member channgeinfo
	Public Function Get_MemberInfo(iAccountId)
	'0=Id,1=Reg_Name,2=Sex,3=Email,4=RegTime,5=Login,6=UserName,7=BirtDay,8=User_Group,9=State,10=HomePage,11=QQ,12=ICQ,13=MSN,14=Comefrom,15=Password
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_Member_Info "&iAccountId
		Case 1
			SQL="SELECT [Id], Reg_Name, Sex, Email, RegTime, Login, [UserName], BirtDay, User_Group, State, HomePage, QQ, ICQ, MSN, Comefrom, Reg_Pass, Cookies"
			SQL=SQL&" From [NB_User]"
			SQL=SQL&" Where [Id]="&iAccountId
		Case 2
			SQL="Exec sp_EliteArticle_Member_Info_Select"
			SQL=SQL&" @Member_Id="&iAccountId
		End Select
		
		Get_MemberInfo=DB_Query(SQL)
	End Function
'-------------------------------------------------------------------
	
'*******************************************************************
'getpass
	Public Function Get_MemberQuestionByAccountId(iAccountId)
		'0=id,1=question,2=answer
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_Member_SafetyInfoById "&iAccountId
		Case 1
			SQL="SELECT Id, Question, Answer"
			SQL=SQL&" FROM NB_User"
			SQL=SQL&" WHERE Id="&iAccountId
		Case 2
			SQL="Exec sp_EliteArticle_Member_SafetyInfoById_Select"
			SQL=SQL&" @MemberId="&iAccountId
		End Select
		
		Get_MemberQuestionByAccountId=DB_Query(SQL)
	End Function
	
	Public Function Get_MemberQuestionByAccountName(sAccountName)
		'0=id,1=question,2=answer
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_Member_SafetyInfoByName '"&sAccountName&"'"
		Case 1
			SQL="SELECT Id, Question, Answer"
			SQL=SQL&" FROM NB_User"
			SQL=SQL&" WHERE Reg_Name='"&sAccountName&"'"
		Case 2
			SQL="Exec sp_EliteArticle_Member_SafetyInfoByName_Select"
			SQL=SQL&" @MemberName='"&sAccountName&"'"
		End Select
		
		Get_MemberQuestionByAccountName=DB_Query(SQL)
	End Function
	
	Public Function Set_MemberPasswordByAccountName(sAccountName,NewPassword)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_UpDate_Member_Password '"&NewPassword&"','"&sAccountName&"'"
		Case 1
			SQL="UPDATE NB_User SET Reg_Pass = '"&NewPassword&"'"
			SQL=SQL&" WHERE Reg_Name='"&sAccountName&"'"
		Case 2
			SQL="Exec sp_EliteArticle_Member_Password_UpDate"
			SQL=SQL&" @New_Password='"&NewPassword&"'"
			SQL=SQL&",@Member_Name='"&sAccountName&"'"
		End Select
		
		Set_MemberPasswordByAccountName=DB_Execute(SQL)
	End Function
'-------------------------------------------------------------------
	
'*******************************************************************
'member_appear
	Public Function Get_MemberDayPostTotal(iAccountId)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_Member_DayPostTotal "&iAccountId
		Case 1
			SQL="SELECT Count([Id])"
			SQL=SQL&" FROM NB_Content"
			SQL=SQL&" WHERE AuthorId="&iAccountId&" And DateDiff(d,GetDate(),AddDate)=0"
		Case 2
			SQL="Exec sp_EliteArticle_Member_DayPost_Select"
			SQL=SQL&" @Member_Id="&iAccountId
		End Select
		
		Get_MemberDayPostTotal=DB_Query(SQL)
	End Function
	
	Public Function Get_MemberAppearColumnList()
		Select Case iDataBaseType
		Case 0, 1
			SQL="SELECT [Id], Title, Code, Case When IsTop = 0 Then '' Else '[导航]' End"
			SQL=SQL&" FROM NB_Column"
			SQL=SQL&" WHERE IsPost=1"
			SQL=SQL&" ORDER BY Code"
		Case 2
			SQL="Exec sp_EliteArticle_Column_MemberAppearList_Select"
		End Select
		
		Get_MemberAppearColumnList=DB_Query(SQL)
	End Function
	
	Public Sub Set_MemberAppearTotal(iAccountId)
		Dim Temp
		
		Select Case iDataBaseType
		Case 0
			SQL="Select Count([Id]) From [NB_Content] Where AuthorId="&iAccountId&" And IsPass=-1 And IsDel=0"
			Temp=DB_Query(SQL)
			
			SQL="UpDate [NB_User] Set PostTotal="&Temp(0,0)&" Where [Id]="&iAccountId
			DB_Execute SQL
		Case 1
			SQL="UpDate NB_User Set PostTotal=("
			SQL=SQL&" Select Count([Id])"
			SQL=SQL&" From [NB_Content]"
			SQL=SQL&" Where AuthorId="&iAccountId&" And IsPass=1 And IsDel=0"
			SQL=SQL&")"
			SQL=SQL&" Where [Id]="&iAccountId
			DB_Execute SQL
		End Select
	End Sub
'-------------------------------------------------------------------
	
'*******************************************************************
'member_reg
	Public Function Get_MemberChkEMail(sMailAddress,iMemberId)
		Dim Temp,SQL
		
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_Member_ChkMail '"&sMailAddress&"'"
		Case 1
			SQL="SELECT Id"
			SQL=SQL&" FROM NB_User"
			SQL=SQL&" WHERE Email='"&sMailAddress&"'"
		Case 2
			SQL="Exec sp_EliteArticle_Member_ChkMail_Select"
			SQL=SQL&" @Member_Mail='"&sMailAddress&"'"
		End Select
		
		Temp=DB_Query(SQL)
		
		If IsArray(Temp) Then
			If iMemberId<>0 Then
				If CLng(Temp(0,0))=CLng(iMemberId) Then 
					Get_MemberChkEMail=True
				Else
					Get_MemberChkEMail=False
				End If
			Else
				Get_MemberChkEMail=False
			End If
		Else
			Get_MemberChkEMail=True
		End If
	End Function
	
	Public Function Set_RegistrationMember(vMemberInfo)
		Dim ChkUser
		Dim Sys_IsPass
		Dim Flag

		ChkUser=False
		Sys_IsPass=EA_Pub.SysInfo(9)
		Flag=0
		
		If EA_Pub.SysInfo(8)="1" And Not Get_MemberChkEMail(vMemberInfo(2),0) Then Flag=-1
		
		If Flag=0 Then
			If Rs.State=1 Then Rs.Close

			Sql="SELECT * FROM [NB_User] WHERE Reg_Name='"&vMemberInfo(0)&"'"
			Rs.Open Sql,Conn,1,3
			If Rs.RecordCount>0 Then
				Flag=2
			Else
				Rs.AddNew
					Rs("Reg_Name")=vMemberInfo(0)
					Rs("Reg_Pass")=vMemberInfo(1)
					Rs("Email")=vMemberInfo(2)
					Rs("Question")=vMemberInfo(3)
					Rs("Answer")=vMemberInfo(4)
					Rs("sex")=vMemberInfo(5)
					Rs("HomePage")=vMemberInfo(6)
					Rs("QQ")=vMemberInfo(7)
					Rs("ICQ")=vMemberInfo(8)
					Rs("MSN")=vMemberInfo(9)
					Rs("UserName")=vMemberInfo(10)
					Rs("BirtDay")=vMemberInfo(11)
					Rs("ComeFrom")=vMemberInfo(12)
					Rs("RegIP")=vMemberInfo(13)
					Rs("State")=Sys_IsPass
					Rs("User_Group")=1
					Rs("Cookies")=0
				Rs.Update

				If iDataBaseType<>2 Then Set_SystemUserTotal 1

				SQL="UPDATE [NB_UserGroup] SET UserTotal = UserTotal + 1 WHERE [Id] = 1"
				Conn.Execute(SQL)

				If Sys_IsPass="0" Then 
					Flag=0
				Else
					Flag=1
				End If
			End If
			Rs.Close
		End If
		
		Set_RegistrationMember=Flag
	End Function

'*******************************************************************
	Public Sub Set_Column_ManagerTopicTotal(iColumnId,iValue)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_UpDate_Column_ManagerTopicTotal "&iValue&","&iColumnId
		Case 1
			SQL="UPDATE NB_System SET MangerNum = MangerNum+"&iValue&" Where Id="&iColumnId
		Case 2
			SQL="Exec sp_EliteArticle_Column_Stat_UpDate"
			SQL=SQL&" @Action=2"
			SQL=SQL&",@ColumnId="&iColumnId
			SQL=SQL&",@Values="&iValue
		End Select
	
		DB_Execute SQL
	End Sub
	
	Public Sub Set_Column_TopicTotal(iColumnId,iValue)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_UpDate_Column_TopicTotal "&iValue&","&iColumnId
		Case 1
			SQL="UPDATE NB_Column SET CountNum = CountNum+"&iValue&" Where Id="&iColumnId
		Case 2
			SQL="Exec sp_EliteArticle_Column_Stat_UpDate"
			SQL=SQL&" @Action=1"
			SQL=SQL&",@ColumnId="&iColumnId
			SQL=SQL&",@Values="&iValue
		End Select
	
		DB_Execute SQL
	End Sub
	
	Public Sub Set_System_ManagerTopicTotal(iValue)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_UpDate_System_MangerTopicTotal "&iValue
		Case 1
			SQL="UPDATE NB_System SET MangerTopicNum = MangerTopicNum+"&iValue
		End Select
	
		DB_Execute SQL
	End Sub

	Public Sub Set_System_TopicTotal(iValue)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_UpDate_System_TopicTotal "&iValue
		Case 1
			SQL="UPDATE NB_System SET TopicNum = TopicNum+"&iValue
		End Select
	
		DB_Execute SQL
	End Sub
	
	Public Sub Set_System_ReviewTotal(iValue)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_UpDate_System_ReviewTotal "&iValue
		Case 1
			SQL="UPDATE NB_System SET ReviewNum = ReviewNum+"&iValue
		End Select
	
		DB_Execute SQL 
	End Sub
	
	Public Sub Set_SystemUserTotal(iValue)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_UpDate_System_UserTotal "&iValue
		Case 1
			SQL="UPDATE NB_System SET RegUser = RegUser+"&iValue
		End Select
	
		DB_Execute SQL 
	End Sub
'-------------------------------------------------------------------
	
'*******************************************************************
'member_fav
	Public Function Set_AddFav(iArticleId,iAccountId)
		Dim Temp,Flag

		Select Case iDataBaseType
		Case 0,1
			Temp=Get_MemberFavTotalByAccountId(iAccountId)(0,0)
			If CLng(Temp)<CLng(EA_Pub.Mem_GroupSetting(12)) Then 
				If Get_IsFavedByArticleId(iArticleId,iAccountId) Then 
					Flag=1
				Else
					SQL="INSERT INTO NB_MyFavorites (ArticleId, UserId, Title)"
					SQL=SQL&" VALUES ("&iArticleId&","&iAccountId&",'')"

					DB_Execute(SQL)
					
					Flag=0
				End If
			Else
				Flag=-1
			End If
		Case 2
			SQL="Exec sp_EliteArticle_Fav_Insert"
			SQL=SQL&" @Member_Id="&iAccountId
			SQL=SQL&",@Article_Id="&iArticleId
			SQL=SQL&",@Fav_Max="&EA_Pub.Mem_GroupSetting(12)
			
			Flag=DB_Execute(SQL)(0,0)
		End Select

		Set_AddFav=Flag
	End Function
	
	Private Function Get_IsFavedByArticleId(iArticleId,iAccountId)
		Dim Temp
		
		Select Case iDataBaseType
		Case 0
			SQL="vi_Select_Member_IsFaved "&iArticleId&","&iAccountId
		Case 1
			SQL="SELECT [Id]"
			SQL=SQL&" FROM NB_MyFavorites"
			SQL=SQL&" WHERE ArticleId="&iArticleId&" And UserId="&iAccountId
		End Select
		
		Temp=DB_Query(SQL)
		
		If IsArray(Temp) Then 
			Get_IsFavedByArticleId=True
		Else
			Get_IsFavedByArticleId=False
		End If
	End Function
	
	Public Function Get_MemberFavTotalByAccountId(iAccountId)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_Member_FavStat "&iAccountId
		Case 1
			SQL="SELECT Count([Id])"
			SQL=SQL&" FROM NB_MyFavorites"
			SQL=SQL&" WHERE UserId="&iAccountId
		Case 2
			SQL="Exec sp_EliteArticle_Fav_Total_Select"
			SQL=SQL&" @Member_Id="&iAccountId
		End Select
		
		Get_MemberFavTotalByAccountId=DB_Query(SQL)
	End Function
	
	Public Function Get_MemberFavListByAccountId(iAccountId,iPageNum,iPageSize)
		Dim Temp
		
		'0=articleid,1=article_posttime,2=article_title,3=favid,4=fav_posttime,5=author,6=author_id
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_Member_FavList "&iAccountId
			Temp=DB_CutPageQuery(SQL,iPageNum,iPageSize)
		Case 1
			SQL="SELECT ArticleId, IsNull(b.Title,GetDate()), Case When b.IsPass=0 Or b.IsDel=1 Or b.Title Is Null Then '该文章已被删除或未通过审核' Else b.Title End, a.[Id], a.AddDate, b.Author, b.AuthorId"
			SQL=SQL&" FROM NB_MyFavorites AS a LEFT JOIN NB_Content AS b ON a.ArticleId=b.Id"
			SQL=SQL&" WHERE UserId="&iAccountId
			SQL=SQL&" ORDER BY a.Id DESC"
			Temp=DB_CutPageQuery(SQL,iPageNum,iPageSize)
		Case 2
			SQL="Exec sp_EliteArticle_Fav_List_Select"
			SQL=SQL&" @Member_Id="&iAccountId
			SQL=SQL&",@List_PageNum="&iPageNum
			SQL=SQL&",@List_PageSize="&iPageSize
			Temp=DB_Query(SQL)
		End Select
		
		Get_MemberFavListByAccountId=Temp
	End Function
	
	Public Sub Del_MemberFav(iFavId,iAccountId)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Delete_Member_Fav "&iFavId&","&iAccountId
		Case 1
			SQL="DELETE "
			SQL=SQL&" FROM NB_MyFavorites"
			SQL=SQL&" WHERE [Id]="&iFavId&" And UserId="&iAccountId
		Case 2
			SQL="Exec sp_EliteArticle_Fav_Delete"
			SQL=SQL&" @Fav_Id="&iFavId
			SQL=SQL&",@Member_Id="&iAccountId
		End Select
		
		DB_Execute(SQL)
	End Sub
'-------------------------------------------------------------------


'*******************************************************************
'Interface
	Public Function Get_InterfaceList(iInterfaceType)
		SQL = "SELECT RemoteURL, StructFile, SKey"
		SQL = SQL & " FROM [NB_Interface]"
		SQL = SQL & " WHERE Type=" & iInterfaceType

		Get_InterfaceList = DB_Query(SQL)
	End Function
	
'-------------------------------------------------------------------


'*******************************************************************
	Private Sub chkDB ()
		If Not IsObject(Conn) Then 
			ConnectionDatabase

			Set Rs=Server.CreateObject("adodb.recordSet")
		End If
	End Sub

	Public Function DB_Execute(sSQL)
		chkDB()

		On Error Resume Next
		Err.Clear 
		
		Conn.Execute(sSQL)
		
		ExecuteTotal=ExecuteTotal+1
		T_SQL_List = T_SQL_List & sSQL & "<br />"
		
		If Err Then 
			If Debug Then
				ErrMsg="在执行以下语句：<br>"
				ErrMsg=ErrMsg&"&nbsp;&nbsp;<font color=800000>"&sSQL&"</font><br>"
				ErrMsg=ErrMsg&"时，发生以下错误：<br>"
				ErrMsg=ErrMsg&"&nbsp;&nbsp;<font color=800000>"&Err.Description&"</font>"
			Else
				ErrMsg="查询数据的时候发现错误。系统已关闭"
			End If
			Call EA_Pub.ShowErrMsg(0,0)
		Else
			DB_Execute=0
		End If
	End Function
	
	Public Function DB_Query(sSQL)
		chkDB()

		On Error Resume Next
		Err.Clear 

		Set Rs=Conn.Execute(sSQL)
		If Not Rs.EOF And Not Rs.BOF Then 
			DB_Query=Rs.GetRows()
		Else
			DB_Query=0
		End If
		Rs.Close 
		
		QueryTotal=QueryTotal+1
		T_SQL_List = T_SQL_List & sSQL & "<br />"
		
		If Err Then 
			If Debug Then
				ErrMsg="在执行以下语句：<br>"
				ErrMsg=ErrMsg&"&nbsp;&nbsp;<font color=800000>"&sSQL&"</font><br>"
				ErrMsg=ErrMsg&"时，发生以下错误：<br>"
				ErrMsg=ErrMsg&"&nbsp;&nbsp;<font color=800000>"&Err.Description&"</font>"
			Else
				ErrMsg="查询数据的时候发现错误。系统已关闭"
			End If
			Call EA_Pub.ShowErrMsg(0,0)
		End If
	End Function
	
	Public Function DB_CutPageQuery(sSQL,iPageNum,iPageSize)
		chkDB()

		On Error Resume Next
		Err.Clear 
		If Rs.State=1 Then Rs.Close

		Rs.Open sSQL,Conn,1,1
		If Not rs.Eof And Not rs.bof Then 
			Rs.AbsolutePosition=Rs.AbsolutePosition+((Abs(iPageNum)-1)*iPageSize)
			DB_CutPageQuery=Rs.GetRows(iPageSize)
		Else
			DB_CutPageQuery=0
		End If
		Rs.Close 
		
		QueryTotal=QueryTotal+1
		T_SQL_List = T_SQL_List & sSQL & "<br />"
		
		If Err Then 
			If Debug Then
				ErrMsg="在执行以下语句：<br>"
				ErrMsg=ErrMsg&"&nbsp;&nbsp;<font color=800000>"&sSQL&"</font><br>"
				ErrMsg=ErrMsg&"时，发生以下错误：<br>"
				ErrMsg=ErrMsg&"&nbsp;&nbsp;<font color=800000>"&Err.Description&"</font>"
			Else
				ErrMsg="查询数据的时候发现错误。系统已关闭"
			End If
			Call EA_Pub.ShowErrMsg(0,0)
		End If
	End Function
'-------------------------------------------------------------------
End Class
%>