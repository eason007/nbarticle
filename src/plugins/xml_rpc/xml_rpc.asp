<!--#Include File="../../include/inc.asp"-->
<!--#Include File="../../include/cls_xml_rpc.asp"-->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：XML-RPC.asp
'= 摘    要：XML-RPC接口文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-02-18
'====================================================================
Dim XML_RPC
Dim objXML
Dim XMLResponse

XMLResponse = Request.BinaryRead(Request.TotalBytes)
Set XML_RPC = New cls_XML_RPC
Set objXML	= Server.CreateObject("Microsoft.XMLDOM")

objXML.load(XMLResponse)

Response.ContentType = "text/xml"

If objXML.readyState = 4 Then
	If objXML.parseError.errorCode <> 0 Then
		'call error
	Else
		Dim objRootNode
		Dim Action
		Dim PostData
		Dim AppKey
		Dim Return,ReturnStr

		Set objRootNode = objXML.documentElement
		FoundErr		= False

		AppKey = objRootNode.selectSingleNode("params/param[0]/value/string").text
		Action = objRootNode.selectSingleNode("methodName").text

		If AppKey <> sCacheName Then FoundErr = True
		
		If Not FoundErr Then
			Select Case Action
			Case "nbarticle.registerUser"
				PostData = objRootNode.selectSingleNode("params/param[1]/value/struct").xml

				Return = XML_RPC.registerUser(PostData)

				If Return <> 0 Then
					XML_RPC.ResponseError Return,ErrMsg
				Else
					XML_RPC.ResponseSuccess 0
				End If
			Case "nbarticle.editUserInfo"
				PostData = objRootNode.selectSingleNode("params/param[1]/value/struct").xml

				Return = XML_RPC.editUserInfo(PostData)

				If Return <> 0 Then
					XML_RPC.ResponseError Return,ErrMsg
				Else
					XML_RPC.ResponseSuccess 0
				End If
			Case "nbarticle.editUserSafeInfo"
				PostData = objRootNode.selectSingleNode("params/param[1]/value/struct").xml

				Return = XML_RPC.editUserSafeInfo(PostData)

				If Return <> 0 Then
					XML_RPC.ResponseError Return,ErrMsg
				Else
					XML_RPC.ResponseSuccess 0
				End If
			End Select
		Else
			'call error
		End If	
	End If
End If
%>
