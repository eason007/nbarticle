<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：cls_template.asp
'= 摘    要：模版类文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-02-24
'====================================================================

	Function Load_ColumnList(Parameter)
		Dim TempStr,ds8,i,j,MainId,URL,URL2,IsCrlf
		Dim ForTotal

		ds8=EA_DBO.Get_Column_List()
		
		If IsArray(ds8) Then
			TempStr="<table>"
			ForTotal = Ubound(ds8,2)

			For i=0 To ForTotal
				URL=EA_Pub.Cov_ColumnPath(ds8(0,i),EA_Pub.SysInfo(18))
				URL2=EA_Pub.Cov_ColumnPath(MainId,EA_Pub.SysInfo(18))

				If Parameter(0)="0" Then 
					If Len(ds8(2,i))=4 Then
						TempStr=TempStr&"<tr>"
						TempStr=TempStr&"<td>&nbsp;*&nbsp;<a href="""&URL&""" title="""&ds8(3,i)&"""><strong>"&ds8(1,i)&"</strong></a></td>"
						TempStr=TempStr&"</tr>"

						If i+1<=UBound(ds8,2) Then
							If Len(ds8(2,i+1))>4 Then
								TempStr=TempStr&"<tr>"
								TempStr=TempStr&"<td>&nbsp;</td>"
								TempStr=TempStr&"<td><table>"
								TempStr=TempStr&"<tr style=""HEIGHT: 20px;"">"
							End If
						End If
						j=1

						IsCrlf = 0
					Else
						TempStr=TempStr&"<td><a href="""&URL&""" title="""&ds8(3,i)&""">"
						TempStr=TempStr&ds8(1,i)&"</a></td>"
						j=j+1

						If j>=CInt(Parameter(1))+1 Then TempStr=TempStr&"</tr><tr><td colspan="""&Parameter(1)&""" style=""height: 1px;BACKGROUND: #000;""></td></tr>":j=1:IsCrlf = 1

						If i+1<=UBound(ds8,2) Then
							If Len(ds8(2,i+1))=4 Then
								If IsCrlf = 0 Then TempStr=TempStr&"</tr>"

								TempStr=TempStr&"</table></td></tr>"
								j=0
							ElseIf IsCrlf = 1 Then
								TempStr=TempStr&"<tr>"
								IsCrlf = 0
							End if
						ElseIf i+1>UBound(ds8,2) Then 
							If IsCrlf = 0 Then TempStr=TempStr&"</tr>"

							TempStr=TempStr&"</table></td></tr>"
						End If
					End If
				Else
					If Len(ds8(2,i))=4 Then
						MainId=ds8(0,i)
						TempStr=TempStr&"<tr>"
						TempStr=TempStr&"<td style=""WIDTH: 25%;"">&nbsp;*&nbsp;<a href="""&URL&""" title="""&ds8(3,i)&"""><strong>"&ds8(1,i)&"</strong></a></td>"
						j=1
						
						If i+1<=UBound(ds8,2) Then
							If Len(ds8(2,i+1))=4 Then 
								TempStr=TempStr&"</tr>"
							Else
								TempStr=TempStr&"<td>"
							End If
						Else
							TempStr=TempStr&"<td></td></tr>"
						End If
					Else
						If j<>0 Then 
							TempStr=TempStr&"<a href="""&URL&""" title="""&ds8(3,i)&""">"
							TempStr=TempStr&ds8(1,i)&"</a>&nbsp;|&nbsp;"
							j=j+1
						
							If j-1>=CInt(Parameter(1)) Then 
								TempStr=TempStr&"<a href="""&URL2&""">更多&gt;&gt;</a></td></tr>"
								j=0
							Else
								If (i+1) > UBound(ds8,2) Then 
									TempStr=TempStr&"</td></tr>"
								Else
									If Len(ds8(2,i+1))=4 Then TempStr=TempStr&"</td></tr>"
								End If
							End If
						End If
					End If
				End If			
			Next
			
			TempStr=TempStr&"</table>"
		Else
			TempStr = ""
		End If
		
		Load_ColumnList=TempStr
	End Function

%>