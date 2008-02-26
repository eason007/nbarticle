<%
'====================================================================
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= FileName:	_cls_template.asp
'= Description:	Template Function Class
'=-------------------------------------------------------------------
'= License: GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= Author:		eason007
'= LastDate:	2006-11-12
'====================================================================

Class cls_NEW_TEMPLATE
	Private S
	Private TemplatePath

	Private Sub Class_Initialize()
		TemplatePath="templates/"
	End Sub
	
	Public Function LoadTemplate(ByRef sFileName)
		Err.Clear 
		On Error Resume Next
		
		Set S = Server.CreateObject("ADOD" & "B.S" & "TREAM")
		With S
			.Mode = 3
			.Type = 2
			.Open
			.LoadFromFile(Server.MapPath(TemplatePath&sFileName))
			LoadTemplate = Bytes2bStr(.ReadText)
			.Close
		End With
		
		Set S = Nothing
		
		If Err Then 
			Response.Clear 
			Response.Write "Load Tempate File:" & sFileName & " Error<br>"
			Response.Write Err.Description 
			Response.End 
		End If
	End Function
	
	Private Function Bytes2bStr(ByVal vin)
	'二进制转换为字符串
		if lenb(vin) =0 then
			Bytes2bStr = ""
			exit function
		end if
	
		Dim BytesStream,StringReturn
		Set BytesStream = Server.CreateObject("ADOD" & "B.S" & "tream")
		BytesStream.Type = 2 
		BytesStream.Open
		BytesStream.WriteText vin
		BytesStream.Position = 0
		BytesStream.Charset = "gb2312"
		BytesStream.Position = 2
		StringReturn = BytesStream.ReadText
		BytesStream.close
		Set BytesStream = Nothing
		Bytes2bStr = StringReturn
	End Function

	Public Function ChkBlock (ByRef sBlockName,ByRef sContent)
		Dim sBlockBeginStr,sBlockEndStr

		sBlockBeginStr	= "<!-- " & sBlockName & " Begin -->"
		sBlockEndStr	= "<!-- " & sBlockName & " End -->"

		If InStr(1,sContent,sBlockBeginStr) And InStr(1,sContent,sBlockEndStr) Then
			ChkBlock = True
		Else
			ChkBlock = False
		End If
	End Function

	Public Function GetBlock(ByRef sBlockName,ByRef sContent)
		Dim iBlockBegin,iBlockEnd
		Dim sBlockBeginStr,sBlockEndStr

		sBlockBeginStr	= "<!-- " & sBlockName & " Begin -->"
		sBlockEndStr	= "<!-- " & sBlockName & " End -->"

		iBlockBegin	= InStr(1,sContent,sBlockBeginStr)
		If iBlockBegin > 0 Then
			iBlockEnd	= InStr(iBlockBegin,sContent,sBlockEndStr)

			GetBlock	= Mid(sContent,iBlockBegin + Len(sBlockBeginStr),iBlockEnd - (iBlockBegin + Len(sBlockBeginStr)))
			
			sContent	= Left(sContent,iBlockBegin-1) & VBCrlf & "<!-- " & sBlockName & "s -->" & VBCrlf &  Right(sContent,Len(sContent)-(iBlockEnd+Len(sBlockEndStr)-1))
		End If
	End Function

	Public Sub SetBlock(ByRef sBlockName,ByRef sBlockContent,ByRef sContent)
		sContent=Replace(sContent & "","<!-- " & sBlockName & "s -->",sBlockContent & VBCrlf & "<!-- " & sBlockName & "s -->")
	End Sub

	Public Sub CloseBlock(ByRef sBlockName,ByRef sContent)
		sContent=Replace(sContent & "","<!-- " & sBlockName & "s -->","")
	End Sub

	Public Sub SetVariable(ByRef sVariableName,ByRef sVariableContent,ByRef sContent)
		sContent=Replace(sContent & "","{$" & sVariableName & "$}",sVariableContent & "")
	End Sub

	Public Sub OutStr(ByRef sContent)
		Response.Clear
		Response.Write sContent
		Set sContent = Nothing
		Response.End
	End Sub

	Public Sub BaseReplace(ByRef sMain)
		Dim Top,Foot
		Dim re
		Set re=new RegExp
		re.IgnoreCase =true
		re.Global=True

		re.Pattern="\{\$(\w+)\$\}"
		sMain=re.Replace(sMain,"")

		re.Pattern="<%(\w+)%\>"
		sMain=re.Replace(sMain,"")
		Set re=Nothing
	End Sub
End Class
%>