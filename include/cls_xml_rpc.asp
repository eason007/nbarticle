<!--#Include File="md5.asp"-->
<%
'====================================================================
'= Team Elite - EliteCMS
'= Copyright (c) 2004 - 2008 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：cls_xml_rpc.asp
'= 摘    要：XML-RPC接口类文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-02-18
'====================================================================

Class cls_XML_RPC
	Public objXML,objXMLHttp,objStream,objXmlDoc
	Public OutInterfaceType,StructData
	Private sStructFile
	Private sStructFileContent
	Private sAppKey

	Private Sub Class_Initialize()
		Set objXML		= Server.CreateObject("Microsoft.XMLDOM")
		Set objXMLHttp	= Server.CreateObject("Msxml2.ServerXMLHTTP")
		Set objStream	= Server.CreateObject("ADOD" & "B.S" & "tream")
	End Sub

	Public Sub Close_Obj()
		Set objXML		= Nothing
		Set objXMLHttp	= Nothing
		Set objStream	= Nothing
	End Sub

	Private Property Let XMLStructFile(sFilePath)
		sStructFile = Server.Mappath(SystemFolder & "plugins/xml_rpc/" & sFilePath)
	End Property

	Private Property Let AppKey(sKey)
		sAppKey = sKey
	End Property

	Public Sub Start_OutInterface()
		Dim Mission
		Dim i
		Dim ForTotal

		Response.ContentType = "text/xml"
		Response.Clear

		Mission = EA_DBO.Get_InterfaceList(OutInterfaceType)
		If IsArray(Mission) Then
			ForTotal = UBound(Mission,2)

			For i=0 To ForTotal
				XMLStructFile = Mission(1,i)
				AppKey		  = Mission(2,i)
				Load_XMLStructFile
				ReplaceData

				Set objXmlDoc = Server.CreateObject("msxml2.FreeThreadedDOMDocument")
				objXmlDoc.ASYNC = False
				objXmlDoc.LoadXml(sStructFileContent)

				PostData Mission(0,i)

				Set objXmlDoc = Nothing
			Next
		End If
	End Sub

	Private Sub PostData(sRemoteURL)
		On Error Resume Next

		objXMLHttp.open "POST", sRemoteURL, false
		objXMLHttp.SetRequestHeader "Content-type", "text/xml"
		objXMLHttp.Send objXmlDoc
	End Sub

	Private Sub ReplaceData()
		Dim i
		Dim ForTotal

		sStructFileContent = Replace(sStructFileContent, "$AppKey", sAppKey)
		sStructFileContent = Replace(sStructFileContent, "$DPOKey", MD5(StructData(0) & sAppKey))

		ForTotal = UBound(StructData)

		For i = 0 To ForTotal
			sStructFileContent = Replace(sStructFileContent, "$" & i & "$", StructData(i))
		Next
	End Sub

	Public Function registerUser(structData)
		Dim MemberInfo(14)
		Dim Feedback

		objXML.LoadXML(structData)

		MemberInfo(0) = objXML.documentElement.selectSingleNode("member[name=""username""]/value/string").text
		MemberInfo(1) = objXML.documentElement.selectSingleNode("member[name=""password""]/value/string").text
		MemberInfo(2) = objXML.documentElement.selectSingleNode("member[name=""email""]/value/string").text
		MemberInfo(3) = objXML.documentElement.selectSingleNode("member[name=""question""]/value/string").text
		MemberInfo(4) = objXML.documentElement.selectSingleNode("member[name=""answer""]/value/string").text

		MemberInfo(5) = objXML.documentElement.selectSingleNode("member[name=""sex""]/value/int").text
		MemberInfo(6) = objXML.documentElement.selectSingleNode("member[name=""homepage""]/value/string").text
		MemberInfo(7) = objXML.documentElement.selectSingleNode("member[name=""qq""]/value/int").text
		MemberInfo(8) = objXML.documentElement.selectSingleNode("member[name=""icq""]/value/int").text
		MemberInfo(9) = objXML.documentElement.selectSingleNode("member[name=""msn""]/value/string").text

		MemberInfo(10) = objXML.documentElement.selectSingleNode("member[name=""name""]/value/string").text
		MemberInfo(11) = objXML.documentElement.selectSingleNode("member[name=""birthday""]/value/string").text
		MemberInfo(12) = objXML.documentElement.selectSingleNode("member[name=""comefrom""]/value/string").text
		MemberInfo(13) = "255.255.255.255"

		MemberInfo(1) = MD5(MemberInfo(1))
		MemberInfo(4) = MD5(MemberInfo(4))

		Feedback = EA_DBO.Set_RegistrationMember(MemberInfo)

		Select Case Feedback
		Case 0
			Application.Lock 
			Application(sCacheName&"IsFlush") = 1
			Application.UnLock 
		Case -1
			ErrMsg = "email registered"
			Feedback = 1
		Case 2
			ErrMsg = "name registered"
		Case 1
			Application.Lock 
			Application(sCacheName&"IsFlush") = 1
			Application.UnLock 

			Feedback = 0
		End Select

		registerUser = Feedback
	End Function

	Public Function editUserInfo(structData)
		Dim Feedback
		Dim MemberInfo(9),MemberName

		objXML.LoadXML(structData)
		
		MemberName	  = objXML.documentElement.selectSingleNode("member[name=""username""]/value/string").text
		MemberInfo(0) = objXML.documentElement.selectSingleNode("member[name=""password""]/value/string").text
		MemberInfo(1) = objXML.documentElement.selectSingleNode("member[name=""email""]/value/string").text
		MemberInfo(2) = objXML.documentElement.selectSingleNode("member[name=""sex""]/value/int").text
		MemberInfo(3) = objXML.documentElement.selectSingleNode("member[name=""homepage""]/value/string").text
		MemberInfo(4) = objXML.documentElement.selectSingleNode("member[name=""qq""]/value/int").text
		MemberInfo(5) = objXML.documentElement.selectSingleNode("member[name=""icq""]/value/int").text
		MemberInfo(6) = objXML.documentElement.selectSingleNode("member[name=""msn""]/value/string").text
		MemberInfo(7) = objXML.documentElement.selectSingleNode("member[name=""name""]/value/string").text
		MemberInfo(8) = objXML.documentElement.selectSingleNode("member[name=""birthday""]/value/string").text
		MemberInfo(9) = objXML.documentElement.selectSingleNode("member[name=""comefrom""]/value/string").text

		MemberInfo(0) = MD5(MemberInfo(0))

		Feedback=EA_DBO.Set_Member_Info(0,MemberInfo,MemberName)

		Select Case Feedback
		Case 0
			
		Case 1
			ErrMsg = "email registered"
		Case -1
			ErrMsg = "password error"
		Case 2
			ErrMsg = "not this member"
		End Select

		editUserInfo = Feedback
	End Function

	Public Function editUserSafeInfo(structData)
		Dim MemberInfo(3)
		Dim MemberName
		Dim Feedback

		objXML.LoadXML(structData)
		
		MemberName	  = objXML.documentElement.selectSingleNode("member[name=""username""]/value/string").text
		MemberInfo(0) = objXML.documentElement.selectSingleNode("member[name=""password""]/value/string").text
		MemberInfo(1) = objXML.documentElement.selectSingleNode("member[name=""newpassword""]/value/string").text
		MemberInfo(2) = objXML.documentElement.selectSingleNode("member[name=""question""]/value/string").text
		MemberInfo(3) = objXML.documentElement.selectSingleNode("member[name=""answer""]/value/string").text

		MemberInfo(0) = MD5(MemberInfo(0))
		MemberInfo(1) = MD5(MemberInfo(1))
		MemberInfo(3) = MD5(MemberInfo(3))

		Feedback=EA_DBO.Set_Member_SafetyInfo(0,MemberInfo,MemberName)

		Select Case Feedback
		Case -1
			ErrMsg	 = "not this member"
			Feedback = 2
		Case 1
			ErrMsg	 = "password error"
		Case 0

		End Select

		editUserSafeInfo = Feedback
	End Function

	Public Sub ResponseError(iErrNum, sErrMsg)
		XMLStructFile = "error.xml"
		Load_XMLStructFile

		sStructFileContent = Replace(sStructFileContent, "$1", iErrNum)
		sStructFileContent = Replace(sStructFileContent, "$2", sErrMsg)

		Response.Clear
		Response.Write sStructFileContent
		Response.End
	End Sub

	Public Sub ResponseSuccess(iCode)
		XMLStructFile = "success.xml"
		Load_XMLStructFile

		sStructFileContent = Replace(sStructFileContent, "$1", iCode)

		Response.Clear
		Response.Write sStructFileContent
		Response.End
	End Sub

	Private Function Load_XMLStructFile()
		On Error Resume Next

		objStream.Mode = 3
		objStream.Type = 2
		objStream.Open
		objStream.LoadFromFile(sStructFile)
		sStructFileContent = objStream.ReadText(objStream.Size)
		objStream.Close

		Load_XMLStructFile = EA_Pub.Bytes2bStr(sStructFileContent, "gb2312")
	End Function
End Class
%>