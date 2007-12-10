<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：/Manager/Comm/cls_XML.asp
'= 摘    要：生成XML类文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2006-05-24
'====================================================================

Class Cls_XML
	Public XMLInfoName(),XMLInfoValue()
	Public XMLElementsName(),XMLElementsValue()
	Public XMLListName(),XMLListValue
	Private infoTmp,elementsTmp
	Private isInfo,isElements

	Private Sub Class_Initialize()
		infoTmp		= 0
		elementsTmp = 0
	End Sub

	Public Sub AppInfo(sName,sValue)
		ReDim Preserve XMLInfoName(infoTmp)
		ReDim Preserve XMLInfoValue(infoTmp)

		XMLInfoName(infoTmp) = sName
		XMLInfoValue(infoTmp) = sValue

		infoTmp = infoTmp + 1
		isInfo = True
	End Sub

	Public Sub AppElements(sName,sValue)
		ReDim Preserve XMLElementsName(elementsTmp)
		ReDim Preserve XMLElementsValue(elementsTmp)

		XMLElementsName(elementsTmp) = sName
		XMLElementsValue(elementsTmp) = sValue

		elementsTmp = elementsTmp + 1
		isElements = True
	End Sub

	Public Function make(ByRef listName,ByRef listValue,ByRef listTotal)
		Dim OutStr

		OutStr = "<?xml version=""1.0"" encoding=""UTF-8""?>"
		OutStr = OutStr & "<NBArticle>"

		If isElements Then	OutStr = OutStr & XMLElements(XMLElementsName,XMLElementsValue)
		If isInfo Then		OutStr = OutStr & XMLInfo(XMLInfoName,XMLInfoValue)

		OutStr = OutStr & "<list total=""" & listTotal & """>"
		If IsArray(listValue) Then		OutStr = OutStr & XMLList(listName,listValue,listTotal)
		OutStr = OutStr & "</list>"

		OutStr = OutStr & "</NBArticle>"

		make = OutStr
	End Function

	Public Function XMLInfo(ByRef vInfoName,ByRef vInfoValue)
		Dim i
		Dim OutStr

		OutStr = "<info>"
		For i = 0 To UBound(vInfoValue)
			If Not IsEmpty(vInfoValue(i)) Then
				If IsNumeric(vInfoValue(i)) Then
					OutStr = OutStr & "<" & vInfoName(i) & ">" & vInfoValue(i) & "</" & vInfoName(i) & ">"
				Else
					OutStr = OutStr & "<" & vInfoName(i) & "><![CDATA[" & vInfoValue(i) & "]]></" & vInfoName(i) & ">"
				End If
			End If
		Next
		OutStr = OutStr & "</info>"

		XMLInfo = OutStr
	End Function

	Public Function XMLElements(ByRef vElementsName,ByRef vElementsValue)
		Dim i
		Dim OutStr

		OutStr = "<elements>"
		For i = 0 To UBound(vElementsValue)
			If Not IsEmpty(vElementsValue(i)) Then
				If IsNumeric(vElementsValue(i)) Then
					OutStr = OutStr & "<element elementID=""" & vElementsName(i) & """>" & vElementsValue(i) & "</element>"
				Else
					OutStr = OutStr & "<element elementID=""" & vElementsName(i) & """><![CDATA[" & vElementsValue(i) & "]]></element>"
				End If
			End If
		Next
		OutStr = OutStr & "</elements>"

		XMLElements = OutStr
	End Function

	Public Function XMLList(ByRef listName,ByRef listValue,ByRef listTotal)
		Dim i, j
		Dim OutStr, tmpStr
		Dim iID
		Dim iCheckbox

		iCheckbox = False

		For i = 0 To UBound(listValue, 2)
			tmpStr = ""

			For j = 0 To UBound(listName)
				If Not IsEmpty(listValue(j,i)) Then
					Select Case listName(j)
						Case "ID"
							iID = listValue(j,i)
						Case "checkbox"
							iCheckbox = True
						Case Else
							If IsNumeric(listValue(j,i)) Then
								tmpStr = tmpStr & "<" & listName(j) & ">" & listValue(j,i) & "</" & listName(j) & ">"
							Else
								If listValue(j,i) = "" Then
									tmpStr = tmpStr & "<" & listName(j) & "><![CDATA[&nbsp;]]></" & listName(j) & ">"
								Else
									tmpStr = tmpStr & "<" & listName(j) & "><![CDATA[" & listValue(j,i) & "]]></" & listName(j) & ">"
								End If
							End If
					End Select
				End If
			Next

			
			OutStr = OutStr & "<item ID=""" & iID & """>"
			If iCheckbox Then OutStr = OutStr & "<checkbox>" & iID & "</checkbox>"
			OutStr = OutStr & tmpStr
			OutStr = OutStr & "</item>"
		Next

		XMLList = OutStr
	End Function

	Public Sub Out(sText)
		Response.Clear
		Response.ContentType = "text/XML"
		Response.Charset = "UTF-8"
		Response.Expires = -1
		Response.ExpiresAbsolute = Now() - 1
		Response.CacheControl = "no-cache"

		Response.Write sText

		Response.Flush
		Response.End
	End Sub
End Class
%>