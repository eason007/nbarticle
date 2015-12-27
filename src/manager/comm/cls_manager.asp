<%
'====================================================================
'= Team Elite - EliteCMS
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：/Manager/Comm/cls_Manager.asp
'= 摘    要：后台-管理后台类文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-03-10
'====================================================================

Class cls_Manager
	Public MasterID
	Public MasterName

	Sub Chk_IsMaster()
		Dim Temp

		If Session(sCacheName&"master_id")="" Then
			Response.Write "-100"
			Response.End
		Else
			Temp=EA_M_DBO.Get_Master_ChkLogin(Session(sCacheName&"master_id"),Session(sCacheName&"master_key"))
			If Not IsArray(Temp) Then
				Response.Write "-100"
				Response.End
			Else
				Randomize
				Login_Key=Cstr(Int((999999-1+100000)*Rnd+1))
				
				Session(sCacheName&"master_key")=Login_Key
				
				EA_M_DBO.Set_Master_LoginKey Login_Key,Session(sCacheName&"master_id")
				
				Admin_Power=Temp(0,0)
				Column_Power=Temp(1,0)

				MasterID	= Session(sCacheName&"master_id")
				MasterName	= Session(sCacheName&"master_name")
			End If
		End If
	End Sub

	Public Sub Error(action)
		'If Rs.State=1 Then Rs.Close:Set Rs=Nothing
		Call EA_Pub.Close_Obj
		Set EA_Pub=Nothing

		Response.Write "505"
		Response.End 
	End Sub

	Public Sub BackUpAccessDataBase()
		Dim dbpath,bkfolder,bkdbname
		Dim fso
		Set Fso=server.createobject("scripting.filesystemobject")
		
		dbpath=server.MapPath (DataBaseFilePath)
		bkfolder=Request.Form ("bkfolder")
		bkdbname=Request.Form ("bkDBname")

		If bkfolder <> "" And bkdbname <> "" Then
			If fso.fileexists(dbpath) Then
				If EA_Pub.CheckDir(bkfolder) = True Then
					fso.copyfile dbpath,bkfolder& "\"& bkdbname
				Else
					EA_Pub.MakeNewsDir bkfolder
					fso.copyfile dbpath,bkfolder& "\"& bkdbname
				End If		
			End If
			Set fso=Nothing
			Response.Write "1"
		Else
			Response.Write "-1"
		End If
	End Sub
	
	Public Function Chk_Power(PowerValue,DestValue)
		If InStr(1,PowerValue&",",DestValue&",")<=0 Then 
			Chk_Power=False
		Else
			Chk_Power=True
		End If
	End Function

	Public Function ShowIp(IpStr)
		Dim TmpStr
		Dim i
		For i=1 To 12 Step 3
			TmpStr=TmpStr&CInt(Mid(IpStr,i,3))&"."
		Next
		ShowIp=Left(TmpStr,Len(TmpStr)-1)
	End Function

	Public Function SplitIp(IpStr)
		Dim TmpStr
		Dim i
		For i=1 To 12 Step 3
			TmpStr=TmpStr&CInt(Mid(IpStr,i,3))&"."
		Next
		TmpStr=Left(TmpStr,Len(TmpStr)-1)
		SplitIp=Split(TmpStr,".")
	End Function

	Function IsObjInstalled(strClassString)
		On Error Resume Next
		IsObjInstalled = False
		Err = 0
		Dim xTestObj
		Set xTestObj = Server.CreateObject(strClassString)
		If 0 = Err Then IsObjInstalled = True
		Set xTestObj = Nothing
		Err = 0
	End Function

	Public Function PageList (iPageValue,iRetCount,iCurrentPage,FieldName,FieldValue)
		Dim Url
		Dim PageCount				'页总数
		Dim PageRoot				'页列表头
		Dim PageFoot				'页列表尾
		Dim OutStr
		Dim i						'输出字符串
		Const StepNum=3				'页码步长
		
		Url=URLStr(FieldName,FieldValue)
		
		If iRetCount = 0 Then iRetCount = 1

		If (iRetCount Mod iPageValue)=0 Then
			PageCount= iRetCount \ iPageValue
		Else
			PageCount= (iRetCount \ iPageValue)+1
		End If
		
		If iCurrentPage-StepNum<=1 Then 
			PageRoot=1
		Else
			PageRoot=iCurrentPage-StepNum
		End If	
		If iCurrentPage+StepNum>=PageCount Then 
			PageFoot=PageCount
		Else
			PageFoot=iCurrentPage+StepNum
		End If
		
		OutStr=iCurrentPage&"/"&PageCount&"页 "
		
		If PageRoot=1 Then
			If iCurrentPage=1 Then 
				OutStr=OutStr&"<font color=888888 face=webdings>9</font></a>"
				OutStr=OutStr&"<font color=888888 face=webdings>7</font></a> "
			Else
				OutStr=OutStr&"<a href='?page=1"
				OutStr=OutStr&Url
				OutStr=OutStr&"' title=""首页""><font face=webdings>9</font></a>"
				OutStr=OutStr&"<a href='?page="&iCurrentPage-1
				OutStr=OutStr&Url
				OutStr=OutStr&"' title=""上页""><font face=webdings>7</font></a> "
			End If
		Else
			OutStr=OutStr&"<a href='?page=1"
			OutStr=OutStr&Url
			OutStr=OutStr&"' title=""首页""><font face=webdings>9</font></a>"
			OutStr=OutStr&"<a href='?page="&iCurrentPage-1
			OutStr=OutStr&Url
			OutStr=OutStr&"' title=""上页""><font face=webdings>7</font></a>..."
		End If
		
		For i=PageRoot To PageFoot
			If i=Cint(iCurrentPage) Then
				OutStr=OutStr&"<font color='red'>["+Cstr(i)+"]</font>&nbsp;"
			Else
				OutStr=OutStr&"<a href='?page="&Cstr(i)
				OutStr=OutStr&Url
				OutStr=OutStr&"'>["+Cstr(i)+"]</a>&nbsp;"
			End If
			If i=PageCount Then Exit For
		Next

		If PageFoot=PageCount Then
			If Cint(iCurrentPage)=Cint(PageCount) Then 
				OutStr=OutStr&"<font color=888888 face=webdings>8</font></a>"
				OutStr=OutStr&"<font color=888888 face=webdings>:</font></a>"
			Else
				OutStr=OutStr&"<a href='?page="&iCurrentPage+1
				OutStr=OutStr&Url
				OutStr=OutStr&"' title=""下页""><font face=webdings>8</font></a>"
				OutStr=OutStr&"<a href='?page="&PageCount
				OutStr=OutStr&Url
				OutStr=OutStr&"' title=""尾页""><font face=webdings>:</font></a>"
			End If
		Else
			OutStr=OutStr&"... <a href='?page="&iCurrentPage+1
			OutStr=OutStr&Url
			OutStr=OutStr&"' title=""下页""><font face=webdings>8</font></a>"
			OutStr=OutStr&"<a href='?page="&PageCount
			OutStr=OutStr&Url
			OutStr=OutStr&"' title=""尾页""><font face=webdings>:</font></a>"
		End If
		
		OutStr=OutStr&"&nbsp;&nbsp;<INPUT TYPE=text class=iptA size=3 value="&iCurrentPage&" onmouseover='this.focus();this.select()' NAME=PGNumber> <INPUT TYPE=button id=button1 name=button1 class=btnA value=GO onclick="&""""&"if(document.all.PGNumber.value>0 && document.all.PGNumber.value<="&PageCount&"){window.location='?Page='+document.all.PGNumber.value+'"&Url&"'}"&""""&" onmouseover='this.focus()' onfocus='this.blur()'>&nbsp;"
		PageList=OutStr
	End Function
	
	Private Function URLStr(FieldName,FieldValue)
		Dim i
		For i=0 to Ubound(FieldName)
			URLStr=URLStr&"&"&CStr(FieldName(i))&"="&CStr(FieldValue(i))
		Next
	End Function
End Class
%>