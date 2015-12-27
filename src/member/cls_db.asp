<%
Class cls_Member_DBOperation
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

			Close()

			Response.Write SysMsg(27)
			Response.End
		End If
	End Sub

	Public Sub Close()
		If Rs.State = 1 Then Rs.Close
		Set Rs=Nothing

		If Conn.State=1 Then Conn.Close
		Set Conn = Nothing
	End Sub

	Public Function Set_Article_Insert(iArticleId,vArticleInfo)
		'On Error Resume Next
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
					EA_DBO.Set_System_ManagerTopicTotal 1
					
					EA_DBO.Set_Column_ManagerTopicTotal vArticleInfo(3),1
				Else
					EA_DBO.Set_System_TopicTotal 1
					
					EA_DBO.Set_Column_TopicTotal vArticleInfo(3),1
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


	Public Function Get_Member_AppearTotal(iAccountId)
		SQL="SELECT Count(Id)"
		SQL=SQL&" FROM NB_Content"
		SQL=SQL&" WHERE AuthorId="&iAccountId&" AND IsDel = 0"
		
		Get_Member_AppearTotal=DB_Query(SQL)
	End Function

	'*******************************************************************
'myappear list
	
	
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


	Public Function Get_MemberAppearColumnList()
		Select Case iDataBaseType
		Case 0, 1
			SQL="SELECT [Id], Title, Code"
			SQL=SQL&" FROM NB_Column"
			SQL=SQL&" WHERE IsPost=" & TrueValue
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
			Call EA_Pub.ShowErrMsg(ErrMsg, 1)
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
			Call EA_Pub.ShowErrMsg(ErrMsg, 1)
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
			Call EA_Pub.ShowErrMsg(ErrMsg, 1)
		End If
	End Function

End Class
%>