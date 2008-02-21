<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：/Manager/Comm/cls_Manager_DB.asp
'= 摘    要：管理-数据库操作类文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-02-21
'====================================================================

Class Cls_Manager_DBOperation
	Private Rs
	Private SQL

	Public TrueValue
	Public ExecuteTotal,QueryTotal
	
	Private Sub Class_Initialize()
		Set Rs=Server.CreateObject("adodb.recordSet")
		ExecuteTotal=0
		QueryTotal=0

		Select Case iDataBaseType
		Case 0
			TrueValue="-1"
		Case 1,2
			TrueValue="1"
		End Select
	End Sub

	Public Sub Close_DB()
		Set Rs=Nothing
	End Sub

	Public Sub Set_Master_LoginLog(sLogin_Key,sCome_Ip,iLogin_Id)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_UpDate_Manager_LoginLog '"&sLogin_Key&"','"&sCome_Ip&"',"&iLogin_Id
		Case 1
			SQL="UPDATE NB_Master SET Cookiess = '"&sLogin_Key&"', LasTime = GetDate(), LastIp = '"&sCome_Ip&"'"
			SQL=SQL&" WHERE Master_Id="&iLogin_Id
		Case 2
			SQL="Exec sp_EliteArticle_Master_LoginLog_UpDate"
			SQL=SQL&" @Login_Key="&sLogin_Key
			SQL=SQL&",@Login_Ip='"&sCome_Ip&"'"
			SQL=SQL&",@Login_Id="&iLogin_Id
		End Select
		
		DB_Execute SQL
	End Sub

	Public Sub Set_Master_LoginKey(iMaster_Key,iMaster_Id)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_UpDate_Manager_LoginKey '"&iMaster_Key&"',"&iMaster_Id
		Case 1
			SQL="UPDATE NB_Master SET Cookiess = "&iMaster_Key
			SQL=SQL&" WHERE Master_Id="&iMaster_Id
		Case 2
			SQL="Exec sp_EliteArticle_Master_LoginKey_UpDate"
			SQL=SQL&" @Master_Key="&iMaster_Key
			SQL=SQL&",@Master_Id="&iMaster_Id
		End Select
		
		DB_Execute SQL
	End Sub
	
	Public Function Get_Master_ChkLogin(iMaster_Id,iMaster_Key)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_Manager_ChkLogin "&iMaster_Id&",'"&iMaster_Key&"'"
		Case 1
			SQL="SELECT Setting, Column_Setting"
			SQL=SQL&" FROM NB_Master"
			SQL=SQL&" WHERE Master_Id="&iMaster_Id&" And Cookiess="&iMaster_Key&" And State=1"
		Case 2
			SQL="Exec sp_EliteArticle_Master_ChkLogin_Select"
			SQL=SQL&" @Master_Id="&iMaster_Id
			SQL=SQL&",@Master_Key="&iMaster_Key
		End Select
		
		Get_Master_ChkLogin=DB_Query(SQL)
	End Function

	Public Function Get_Master_Login(sLogin_Name)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_Manager_Login '"&sLogin_Name&"'"
		Case 1
			SQL="SELECT Master_Password, Master_Id, State"
			SQL=SQL&" FROM NB_Master"
			SQL=SQL&" WHERE Master_Name='"&sLogin_Name&"'"
		Case 2
			SQL="Exec sp_EliteArticle_Master_Login_Select"
			SQL=SQL&" @Master_Name='"&sLogin_Name&"'"
		End Select
		
		Get_Master_Login=DB_Query(SQL)
	End Function
	
	Public Sub Set_Vote_State(iValue,iVote_Id)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_UpDate_Manager_Vote_State "&iValue&","&iVote_Id
		Case 1
			SQL="UPDATE NB_Vote SET Lock = "&iValue
			SQL=SQL&" WHERE Id="&iVote_Id
		Case 2
			SQL="Exec sp_EliteArticle_Vote_State_Manager_UpDate"
			SQL=SQL&" @State="&iValue
			SQL=SQL&",@Vote_Id="&iVote_Id
		End Select
		
		DB_Execute SQL
	End Sub
	
	Public Sub Set_Vote_Delete(iVote_Id)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Delete_Manager_Vote "&iVote_Id
		Case 1
			SQL="DELETE"
			SQL=SQL&" FROM NB_Vote"
			SQL=SQL&" WHERE Id="&iVote_Id
		Case 2
			SQL="Exec sp_EliteArticle_Vote_Delete"
			SQL=SQL&" @Vote_Id="&iVote_Id
		End Select
		
		DB_Execute SQL
	End Sub
	
	Public Function Get_Vote_List(iPageNum,iPageSize)
		Dim Temp
		
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_Manager_VoteList"
			Temp=DB_CutPageQuery(SQL,iPageNum,iPageSize)
		Case 1
			SQL="SELECT Id, Title, VoteTotal, Case Type When 0 Then '单选' Else '多选' End, Case Lock When 0 Then '正常' Else '关闭' End, Lock"
			SQL=SQL&" FROM NB_Vote"
			SQL=SQL&" ORDER BY Id DESC"
			Temp=DB_CutPageQuery(SQL,iPageNum,iPageSize)
		Case 2
			SQL="Exec sp_EliteArticle_Vote_List_Manager_Select"
			SQL=SQL&" @List_PageNum="&iPageNum
			SQL=SQL&",@List_PageSize="&iPageSize
			Temp=DB_Query(SQL)
		End Select
		
		Get_Vote_List=Temp
	End Function
	
	Public Function Get_Vote_Stat()
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_Manager_VoteStat"
		Case 1
			SQL="SELECT Count([ID])"
			SQL=SQL&" FROM NB_Vote"
		Case 2
			SQL="Exec sp_EliteArticle_Vote_Stat_Manager_Select"
		End Select
		
		Get_Vote_Stat=DB_Query(SQL)
	End Function
	
	Public Function Get_Group_ChanngeMemberGroup(iDest_Id,iSour_Id)
		Select Case iDataBaseType
		Case 0
			Sql="Exec vi_UpDate_Manager_ChanngeGroupByMemberId "&iDest_Id&","&iSour_Id
		Case 1
			SQL="UPDATE NB_User SET User_Group = "&iDest_Id
			SQL=SQL&" WHERE [ID]="&iSour_Id
		Case 2
			SQL="Exec sp_EliteArticle_UserGroup_ChanngeMemberGroup_Manager_UpDate"
			SQL=SQL&" @Dest_Id="&iDest_Id
			SQL=SQL&",@Sour_Id="&iSour_Id
		End Select
		
		DB_Execute SQL
	End Function
	
	Public Function Get_Group_ForMemberList(iGroup_Id,iPageNum,iPageSize)
		Dim Temp
		
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_Manager_GroupMemberList "&iGroup_Id
			Temp=DB_CutPageQuery(SQL,iPageNum,iPageSize)
		Case 1
			SQL="SELECT Reg_Name, Email, RegTime, Case State When 0 Then '等待审核' Else '正常' End, Id"
			SQL=SQL&" FROM NB_User"
			SQL=SQL&" WHERE User_Group="&iGroup_Id
			SQL=SQL&" ORDER BY State DESC"
			Temp=DB_CutPageQuery(SQL,iPageNum,iPageSize)
		Case 2
			SQL="Exec sp_EliteArticle_UserGroup_MemberList_Manager_Select"
			SQL=SQL&" @Group_Id="&iGroup_Id
			SQL=SQL&",@List_PageNum="&iPageNum
			SQL=SQL&",@List_PageSize="&iPageSize
			Temp=DB_Query(SQL)
		End Select
		
		Get_Group_ForMemberList=Temp
	End Function
	
	Public Function Get_Group_ForMemberTotal(iGroup_Id)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_Manager_GroupMemberTotal "&iGroup_Id
		Case 1
			SQL="SELECT Count([ID])"
			SQL=SQL&" FROM NB_User"
			SQL=SQL&" WHERE User_Group="&iGroup_Id
		Case 2
			SQL="Exec sp_EliteArticle_UserGroup_MemberTotal_Manager_Select"
			SQL=SQL&" @Group_Id="&iGroup_Id
		End Select
		
		Get_Group_ForMemberTotal=DB_Query(SQL)
	End Function
	
	Public Sub Set_Group_Delete(iGroup_Id)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Delete_Manager_Group "&iGroup_Id
			DB_Execute SQL
			
			SQL="Exec vi_UpDate_Manager_ChanngeGroupByGroupId 1,"&iGroup_Id
			DB_Execute SQL
		Case 1
			SQL="DELETE"
			SQL=SQL&" FROM NB_UserGroup"
			SQL=SQL&" WHERE Id="&iGroup_Id
			DB_Execute SQL
			
			SQL="UPDATE NB_User SET User_Group = 1"
			SQL=SQL&" WHERE User_Group="&iGroup_Id
			DB_Execute SQL
		Case 2
			SQL="Exec sp_EliteArticle_UserGroup_Delete"
			SQL=SQL&" @Group_Id="&iGroup_Id
		End Select
	End Sub
	
	Public Function Get_Group_List()
		Select Case iDataBaseType
		Case 0
			Sql="Exec vi_Select_Manager_GroupList"
		Case 1
			SQL="SELECT Id, GroupName, UserTotal"
			SQL=SQL&" FROM NB_UserGroup"
		Case 2
			SQL="Exec sp_EliteArticle_UserGroup_List_Manager_Select"
		End Select
		
		Get_Group_List=DB_Query(SQL)
	End Function
	
	Public Sub Set_Theme_Delete(iTheme_Id)
		SQL="DELETE"
		SQL=SQL&" FROM NB_Themes"
		SQL=SQL&" WHERE Id="&iTheme_Id
		
		DB_Execute SQL
	End Sub
	
	Public Sub Set_DefaultTheme(iTheme_Id)
		SQL="UPDATE NB_Themes SET IsDefault = 0"
		DB_Execute SQL
		
		SQL="UPDATE NB_Themes SET IsDefault = 1"
		SQL=SQL&" WHERE Id="&iTheme_Id
		DB_Execute SQL
	End Sub
	
	Public Function Get_Theme_Info(iTheme_Id)
		SQL="SELECT Id, Title, IsDefault"
		SQL=SQL&" FROM NB_Themes"
		SQL=SQL&" WHERE Id="&iTheme_Id
		
		Get_Theme_Info=DB_Query(SQL)
	End Function
	
	Public Function Get_Theme_List()
		SQL="SELECT Id, Title, IsDefault"
		SQL=SQL&" FROM NB_Themes"
		SQL=SQL&" ORDER BY IsDefault DESC, ID ASC"

		Get_Theme_List=DB_Query(SQL)
	End Function

	Public Function Get_Module_Total(iThemeId)
		SQL="SELECT COUNT(ID)"
		SQL=SQL&" FROM NB_Module"
		SQL=SQL&" WHERE ThemesID="&iThemeId

		Get_Module_Total=DB_Query(SQL)
	End Function

	Public Function Get_Module_List(iThemeId)
		SQL="SELECT Id, Title, Desc, ThemesID, [Code], [Type]"
		SQL=SQL&" FROM NB_Module"
		SQL=SQL&" WHERE ThemesID="&iThemeId

		Get_Module_List=DB_Query(SQL)
	End Function

	Public Function Get_Module_Info(iModule_Id)
		SQL="SELECT Id, Title, Desc, [Code], [Type]"
		SQL=SQL&" FROM NB_Module"
		SQL=SQL&" WHERE Id="&iModule_Id
		
		Get_Module_Info=DB_Query(SQL)
	End Function

	Public Sub Set_Module_Delete(iModule_Id)
		SQL="DELETE"
		SQL=SQL&" FROM NB_Module"
		SQL=SQL&" WHERE Id="&iModule_Id
		
		DB_Execute SQL
	End Sub
	
	Public Sub Set_Review_Pass(iValue,iReview_Id)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_UpDate_Manager_ReviewPassStat "&iValue&","&iReview_Id
		Case 1
			SQL="UPDATE NB_Review SET IsPass = "&iValue
			SQL=SQL&" WHERE Id="&iReview_Id
		Case 2
			SQL="Exec sp_EliteArticle_Review_PassState_UpDate"
			SQL=SQL&" @Value="&iValue
			SQL=SQL&",@Review_Id="&iReview_Id
		End Select
		
		DB_Execute SQL
		
		EA_DBO.Set_System_ReviewTotal 1
	End Sub
	
	Public Sub Set_Review_Delete(iReview_Id)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Delete_Manager_Review "&iReview_Id
		Case 1
			SQL="DELETE"
			SQL=SQL&" FROM NB_Review"
			SQL=SQL&" WHERE Id="&iReview_Id
		Case 2
			SQL="Exec sp_EliteArticle_Review_Delete"
			SQL=SQL&" @Review_Id="&iReview_Id
		End Select
		
		DB_Execute SQL
	End Sub
	
	Public Function Get_Review_Info(iReview_Id)
		Select Case iDataBaseType
		Case 0
			Sql="Exec vi_Select_Manager_ReviewInfo "&iReview_Id
		Case 1
			SQL="SELECT Id, Content"
			SQL=SQL&" FROM NB_Review"
			SQL=SQL&" WHERE Id="&iReview_Id
		Case 2
			SQL="Exec sp_EliteArticle_Review_Info_Manager_Select"
			SQL=SQL&" @Review_Id="&iReview_Id
		End Select
		
		Get_Review_Info=DB_Query(SQL)
	End Function
	
	Public Function Get_Placard_List(iPageNum,iPageSize)
		Dim Temp
		
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_Manager_PlacardList"
			Temp=DB_CutPageQuery(SQL,iPageNum,iPageSize)
		Case 1
			SQL="SELECT Id, Title, AddTime, OverTime"
			SQL=SQL&" FROM NB_Placard"
			SQL=SQL&" ORDER BY OverTime DESC"
			Temp=DB_CutPageQuery(SQL,iPageNum,iPageSize)
		Case 2
			SQL="Exec sp_EliteArticle_Placard_List_Select"
			SQL=SQL&" @List_PageNum="&iPageNum
			SQL=SQL&",@List_PageSize="&iPageSize
			Temp=DB_Query(SQL)
		End Select
		
		Get_Placard_List=Temp
	End Function
	
	Public Function Get_Placard_Total()
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_Manager_PlacardStat"
		Case 1
			SQL="SELECT Count(Id)"
			SQL=SQL&" FROM NB_Placard"
		Case 2
			SQL="Exec sp_EliteArticle_Placard_Total_Select"
		End Select
		
		Get_Placard_Total=DB_Query(SQL)
	End Function
	
	Public Sub Set_Member_Delete(iMember_Id)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Delete_Manager_Member "&iMember_Id
		Case 1
			SQL="DELETE"
			SQL=SQL&" FROM NB_User"
			SQL=SQL&" WHERE Id="&iMember_Id
		Case 2
			SQL="Exec sp_EliteArticle_Member_Delete"
			SQL=SQL&" @Member_Id="&iMember_Id
		End Select
		
		DB_Execute SQL
		
		EA_DBO.Set_SystemUserTotal -1
	End Sub
	
	Public Function Get_Member_Info(iMember_Id)
		Select Case iDataBaseType
		Case 0
			Sql="Exec vi_Select_Member_Info "&iMember_Id
		Case 1
			SQL="SELECT Id, Reg_Name, Sex, Email, RegTime, Login, [UserName], BirtDay, User_Group, State, HomePage, QQ, ICQ, MSN, Comefrom, Reg_Pass, Cookies"
			SQL=SQL&" FROM NB_User"
			SQL=SQL&" WHERE Id="&iMember_Id
		Case 2
			SQL="Exec sp_EliteArticle_Member_Info_Select"
			SQL=SQL&" @Member_Id="&iMember_Id
		End Select
		
		Get_Member_Info=DB_Query(SQL)
	End Function
	
	Public Function Get_Master_Info(iMasterId)
		Select Case iDataBaseType
		Case 0
			Sql="Exec vi_Select_Manager_MasterInfo "&iMasterId
		Case 1
			SQL="SELECT Master_Name, State, Setting, Column_Setting"
			SQL=SQL&" FROM NB_Master"
			SQL=SQL&" WHERE Master_Id="&iMasterId
		Case 2
			SQL="Exec sp_EliteArticle_Master_Info_Manager_Select"
			SQL=SQL&" @Master_Id="&iMasterId
		End Select
		
		Get_Master_Info=DB_Query(SQL)
	End Function
	
	Public Function Get_Master_List()
		Select Case iDataBaseType
		Case 0
			Sql="Exec vi_Select_Manager_MasterList"
		Case 1
			SQL="SELECT Master_Id, Master_Name, LasTime, LastIp, Case State When 1 Then '正常' Else '禁止' End"
			SQL=SQL&" FROM NB_Master"
			SQL=SQL&" ORDER BY Master_Id DESC"
		Case 2
			SQL="Exec sp_EliteArticle_Master_List_Manager_Select"
		End Select
		
		Get_Master_List=DB_Query(SQL)
	End Function
	
	Public Sub Set_Js_Delete(iJs_Id)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Delete_Manager_JsFile "&iJs_Id
		Case 1
			SQL="DELETE"
			SQL=SQL&" FROM NB_JsFile"
			SQL=SQL&" WHERE Id="&iJs_Id
		Case 2
			SQL="Exec sp_EliteArticle_JsFile_Delete"
			SQL=SQL&" @Js_Id="&iJs_Id
		End Select
		
		DB_Execute SQL
	End Sub
	
	Public Function Get_Js_Info(iJsId)
		Select Case iDataBaseType
		Case 0
			Sql="Exec vi_Select_Manager_JsFileInfo "&iJsId
		Case 1
			SQL="SELECT Title, Info, FileName, Setting"
			SQL=SQL&" FROM NB_JsFile"
			SQL=SQL&" WHERE Id="&iJsId
		Case 2
			SQL="Exec sp_EliteArticle_JsFile_Info_Manager_Select"
			SQL=SQL&" @Js_Id="&iJsId
		End Select
		
		Get_Js_Info=DB_Query(SQL)
	End Function
	
	Public Function Get_Js_List()
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_Manager_JsFileList"
		Case 1
			SQL="SELECT Id, Title, Info, FileName"
			SQL=SQL&" FROM NB_JsFile"
		Case 2
			SQL="Exec sp_EliteArticle_JsFile_List_Manager_Select"
		End Select
		
		Get_Js_List=DB_Query(SQL)
	End Function
	
	Public Sub Set_IP_Delete(iIP_Id)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Delete_Manager_IP "&iIP_Id
		Case 1
			SQL="DELETE"
			SQL=SQL&" FROM NB_IP"
			SQL=SQL&" WHERE Id="&iIP_Id
		Case 2
			SQL="Exec sp_EliteArticle_IP_Delete"
			SQL=SQL&" @IP_Id="&iIP_Id
		End Select
		
		DB_Execute SQL
	End Sub
	
	Public Function Get_IP_Info(iIP_Id)
		Select Case iDataBaseType
		Case 0
			Sql="Exec vi_Select_Manager_IPInfo "&iIP_Id
		Case 1
			SQL="SELECT Id, Head_Ip, Foot_Ip, OverTime"
			SQL=SQL&" FROM NB_Ip"
			SQL=SQL&" WHERE Id="&iIP_Id
		Case 2
			SQL="Exec sp_EliteArticle_IP_Info_Manager_Select"
			SQL=SQL&" @IP_Id="&iIP_Id
		End Select
		
		Get_IP_Info=DB_Query(SQL)
	End Function
	
	Public Function Get_IP_List(iPageNum,iPageSize)
		Dim Temp
		
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_Manager_IPList"
			Temp=DB_CutPageQuery(SQL,iPageNum,iPageSize)
		Case 1
			SQL="SELECT Id, Head_Ip, Foot_Ip, OverTime"
			SQL=SQL&" FROM NB_Ip"
			SQL=SQL&" ORDER BY Id DESC"
			Temp=DB_CutPageQuery(SQL,iPageNum,iPageSize)
		Case 2
			SQL="Exec sp_EliteArticle_IP_List_Manager_Select"
			SQL=SQL&" @List_PageNum="&iPageNum
			SQL=SQL&",@List_PageSize="&iPageSize
			Temp=DB_Query(SQL)
		End Select
		
		Get_IP_List=Temp
	End Function
	
	Public Function Get_Ip_Total()
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_Manager_IPStat"
		Case 1
			SQL="SELECT Count(Id)"
			SQL=SQL&" FROM NB_IP"
		Case 2
			SQL="Exec sp_EliteArticle_InsideLink_Total_Manager_Select"
		End Select
		
		Get_Ip_Total=DB_Query(SQL)
	End Function
	
	Public Sub Set_InsideLink_Delete(iInsideLink_Id)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Delete_Manager_InsideLink "&iInsideLink_Id
		Case 1
			SQL="DELETE"
			SQL=SQL&" FROM NB_Link"
			SQL=SQL&" WHERE Id="&iInsideLink_Id
		Case 2
			SQL="Exec sp_EliteArticle_InsideLink_Delete"
			SQL=SQL&" @InsideLink_Id="&iInsideLink_Id
		End Select
		
		DB_Execute SQL
	End Sub
	
	Public Function Get_InsideLink_Info(iInsideLink_Id)
		Select Case iDataBaseType
		Case 0
			Sql="Exec vi_Select_Manager_InsideLinkInfo "&iInsideLink_Id
		Case 1
			SQL="SELECT Word, Link, ColumnId"
			SQL=SQL&" FROM NB_Link"
			SQL=SQL&" WHERE Id="&iInsideLink_Id
		Case 2
			SQL="Exec sp_EliteArticle_InsideLink_Info_Manager_Select"
			SQL=SQL&" @InsideLink_Id="&iInsideLink_Id
		End Select
		
		Get_InsideLink_Info=DB_Query(SQL)
	End Function
	
	Public Function Get_InsideLink_List(iPageNum,iPageSize)
		Dim Temp
		
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_Manager_InsideLinkList"
			Temp=DB_CutPageQuery(SQL,iPageNum,iPageSize)
		Case 1
			SQL="SELECT l.Id, Word, Link, IsNull(c.Title,'全站')"
			SQL=SQL&" FROM NB_Link AS l LEFT JOIN NB_Column AS c ON l.ColumnId=c.Id"
			SQL=SQL&" ORDER BY l.Id DESC"
			Temp=DB_CutPageQuery(SQL,iPageNum,iPageSize)
		Case 2
			SQL="Exec sp_EliteArticle_InsideLink_List_Manager_Select"
			SQL=SQL&" @List_PageNum="&iPageNum
			SQL=SQL&",@List_PageSize="&iPageSize
			Temp=DB_Query(SQL)
		End Select
		
		Get_InsideLink_List=Temp
	End Function
	
	Public Function Get_InsideLink_Total()
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_Manager_InsideLinkStat"
		Case 1
			SQL="SELECT Count(Id)"
			SQL=SQL&" FROM NB_Link"
		Case 2
			SQL="Exec sp_EliteArticle_InsideLink_Total_Manager_Select"
		End Select
		
		Get_InsideLink_Total=DB_Query(SQL)
	End Function
	
	Public Sub Set_Friend_Delete(iFriend_Id)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Delete_Manager_Friend "&iFriend_Id
		Case 1
			SQL="DELETE"
			SQL=SQL&" FROM NB_FriendLink"
			SQL=SQL&" WHERE Id="&iFriend_Id
		Case 2
			SQL="Exec sp_EliteArticle_FriendLink_Delete"
			SQL=SQL&" @FriendLink_Id="&iFriend_Id
		End Select
		
		DB_Execute SQL
	End Sub
	
	Public Function Get_Friend_Info(iFriend_Id)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_Manager_FriendInfo "&iFriend_Id
		Case 1
			SQL="SELECT LinkName, LinkURL, LinkImgPath, LinkInfo, ColumnId, OrderNum, State, Style"
			SQL=SQL&" FROM NB_FriendLink"
			SQL=SQL&" WHERE id="&iFriend_Id
		Case 2
			SQL="Exec sp_EliteArticle_FriendLink_Info_Select"
			SQL=SQL&" @FriendLink_Id="&iFriend_Id
		End Select
		
		Get_Friend_Info=DB_Query(SQL)
	End Function
	
	Public Function Get_Friend_List(iPageNum,iPageSize)
		Dim Temp
		
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_Manager_FriendList"
			Temp=DB_CutPageQuery(SQL,iPageNum,iPageSize)
		Case 1,2
			SQL="SELECT a.Id, LinkName, LinkImgPath, LinkUrl, IsNull(b.Title,'首页'), a.OrderNum, Case a.State When 1 Then '审核通过' Else '审核不通过' End"
			SQL=SQL&" FROM NB_FriendLink AS a LEFT JOIN NB_Column AS b ON a.columnid=b.id"
			SQL=SQL&" ORDER BY ColumnId, OrderNum DESC"
			Temp=DB_CutPageQuery(SQL,iPageNum,iPageSize)
		End Select
		
		Get_Friend_List=Temp
	End Function
	
	Public Function Get_Friend_Total()
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_Manager_FriendStat"
		Case 1
			SQL="SELECT Count(Id)"
			SQL=SQL&" FROM NB_FriendLink"
		Case 2
			SQL="Exec sp_EliteArticle_FriendLink_Total_Select"
		End Select

		Get_Friend_Total=DB_Query(SQL)
	End Function
	
	Public Sub Set_Article_PassState(iStateValue,iArticleId)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_UpDate_Manager_ArticlePassState "&iStateValue&","&iArticleId
		Case 1
			SQL="UPDATE NB_Content SET IsPass = "&iStateValue
			SQL=SQL&" WHERE Id="&iArticleId
		Case 2
			SQL="Exec sp_EliteArticle_Article_PassState_Manager_UpDate"
			SQL=SQL&" @Value="&iStateValue
			SQL=SQL&",@Article_Id="&iArticleId
		End Select
		
		DB_Execute SQL
	End Sub
	
	Public Sub Set_Article_Resume(iArticle_Id)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_UpDate_Manager_ResumeArticle "&iArticle_Id
		Case 1
			SQL="UPDATE NB_Content SET IsDel = 0"
			SQL=SQL&" WHERE Id="&iArticle_Id
		Case 2
			SQL="Exec sp_EliteArticle_Article_Resume_Manager_UpDate"
			SQL=SQL&" @Article_Id="&iArticle_Id
		End Select
		
		DB_Execute SQL
	End Sub
	
	Public Sub Set_Column_Delete(iColumn_Id)
		Select Case iDataBaseType
		Case 0
			Sql="Exec vi_Delete_Manager_Column "&iColumn_Id
		Case 1
			SQL="Delete"
			SQL=SQL&" FROM NB_Column"
			SQL=SQL&" WHERE Id="&iColumn_Id
		Case 2
			SQL="Exec sp_EliteArticle_Column_Delete"
			SQL=SQL&" @Column_Id="&iColumn_Id
		End Select
		
		DB_Execute SQL
	End Sub
	
	Public Function Get_Column_ArticleTotal(iColumn_Id)
		Select Case iDataBaseType
		Case 0
			Sql="Exec vi_Select_Manager_Column_ArticleTotal "&iColumn_Id
		Case 1
			SQL="SELECT Count(Id)"
			SQL=SQL&" FROM NB_Content"
			SQL=SQL&" WHERE ColumnId="&iColumn_Id
		Case 2
			SQL="Exec sp_EliteArticle_Column_ArticleTotal_Manager_Select"
			SQL=SQL&" @ColumnId="&iColumn_Id
		End Select

		Get_Column_ArticleTotal=DB_Query(SQL)
	End Function
	
	Public Sub Set_ArticleTemp_Del(iTemp_Id)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Delete_Manager_ArticleTemplate "&iTemp_Id
		Case 1
			SQL="DELETE"
			SQL=SQL&" FROM NB_ArticleTemplate"
			SQL=SQL&" WHERE Id="&iTemp_Id
		Case 2
			SQL="Exec sp_EliteArticle_ArticleTemp_Delete"
			SQL=SQL&" @Temp_Id="&iTemp_Id
		End Select
		
		DB_Execute SQL
	End Sub
	
	Public Function Get_ArticleTemp_Info(iTemp_Id)
		Select Case iDataBaseType
		Case 0
			Sql="Exec vi_Select_Manager_ArticleTempInfo "&iTemp_Id
		Case 1
			SQL="SELECT Title, Content"
			SQL=SQL&" FROM NB_ArticleTemplate"
			SQL=SQL&" WHERE Id="&iTemp_Id
		Case 2
			SQL="Exec sp_EliteArticle_ArticleTemp_Info_Manager_Select"
			SQL=SQL&" @Temp_Id="&iTemp_Id
		End Select
		
		Get_ArticleTemp_Info=DB_Query(SQL)
	End Function
	
	Public Function Get_ArticleTemp_List(iPageNum,iPageSize)
		Dim Temp
		
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_Manager_ArticleTempList"
			Temp=DB_CutPageQuery(SQL,iPageNum,iPageSize)
		Case 1
			SQL="SELECT Id, Title"
			SQL=SQL&" FROM NB_ArticleTemplate"
			Temp=DB_CutPageQuery(SQL,iPageNum,iPageSize)
		Case 2
			SQL="Exec sp_EliteArticle_ArticleTemp_List_Manager_Select"
			SQL=SQL&" @List_PageNum="&iPageNum
			SQL=SQL&",@List_PageSize="&iPageSize
			Temp=DB_Query(SQL)
		End Select
		
		Get_ArticleTemp_List=Temp
	End Function
	
	Public Function Get_ArticleTemp_Total()
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_Manager_ArticleTempStat"
		Case 1
			SQL="SELECT Count([Id])"
			SQL=SQL&" FROM NB_ArticleTemplate"
		Case 2
			SQL="Exec sp_EliteArticle_ArticleTemp_Total_Manager_Select"
		End Select
		
		Get_ArticleTemp_Total=DB_Query(SQL)
	End Function
	
	Public Sub Set_AdSense_Del(iAdSense_Id)
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Delete_Manager_AdSense "&iAdSense_Id
		Case 1
			SQL="DELETE"
			SQL=SQL&" FROM NB_AdSense"
			SQL=SQL&" WHERE [ID]="&iAdSense_Id
		Case 2
			SQL="Exec sp_EliteArticle_AdSense_Delete"
			SQL=SQL&" @AdSense_Id="&iAdSense_Id
		End Select
		
		DB_Execute SQL
	End Sub
	
	Public Function Get_AdSense_Info(iAdSense_Id)
		Select Case iDataBaseType
		Case 0
			Sql="Exec vi_Select_Manager_AdSenseInfo "&iAdSense_Id
		Case 1
			SQL="SELECT Title, Content"
			SQL=SQL&" FROM [NB_AdSense]"
			SQL=SQL&" WHERE [Id]="&iAdSense_Id
		Case 2
			SQL="Exec sp_EliteArticle_AdSense_Info_Select"
			SQL=SQL&" @AdSense_Id="&iAdSense_Id
		End Select
		
		Get_AdSense_Info=DB_Query(SQL)
	End Function
	
	Public Function Get_AdSense_List(iPageNum,iPageSize)
		Dim Temp
		
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_Manager_AdSenseList"
			Temp=DB_CutPageQuery(SQL,iPageNum,iPageSize)
		Case 1
			SQL="SELECT [Id], Title"
			SQL=SQL&" FROM NB_AdSense"
			Temp=DB_CutPageQuery(SQL,iPageNum,iPageSize)
		Case 2
			SQL="Exec sp_EliteArticle_AdSense_List_Manager_Select"
			SQL=SQL&" @List_PageNum="&iPageNum
			SQL=SQL&",@List_PageSize="&iPageSize
			Temp=DB_Query(SQL)
		End Select
		
		Get_AdSense_List=Temp
	End Function
	
	Public Function Get_AdSense_Total()
		Select Case iDataBaseType
		Case 0
			SQL="Exec vi_Select_Manager_AdSenseStat"
		Case 1
			SQL="SELECT Count([Id])"
			SQL=SQL&" FROM NB_AdSense"
		Case 2
			SQL="Exec sp_EliteArticle_AdSense_Total_Manager_Select"
		End Select
		
		Get_AdSense_Total=DB_Query(SQL)
	End Function

	Public Function Get_Interface_Total()
		SQL = "SELECT COUNT([ID]) FROM [NB_Interface]"

		Get_Interface_Total=DB_Query(SQL)
	End Function

	Public Function Get_Interface_List(iPageNum,iPageSize)
		Dim Temp
		
		SQL="SELECT [Id],Title,RemoteURL,StructFile,Type"
		SQL=SQL&" FROM [NB_Interface]"
		Temp=DB_CutPageQuery(SQL,iPageNum,iPageSize)
		
		Get_Interface_List=Temp
	End Function

	Public Function Get_Interface_Info(iInterface_Id)
		SQL="SELECT Title,RemoteURL,StructFile,Type,SKey"
		SQL=SQL&" FROM [NB_Interface]"
		SQL=SQL&" WHERE [Id]="&iInterface_Id
		
		Get_Interface_Info=DB_Query(SQL)
	End Function
	
'*******************************************************************
	Public Function DB_Execute(sSQL)
		On Error Resume Next
		Err.Clear 
		
		Conn.Execute(sSQL)
		
		ExecuteTotal=ExecuteTotal+1
		
		If Err Then 
			If EA_Pub.SysInfo(25)="1" Then
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
		
		If Err Then 
			If EA_Pub.SysInfo(25)="1" Then
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
		
		If Err Then 
			If EA_Pub.SysInfo(25)="1" Then
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