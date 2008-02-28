<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：/Manager/Comm/cls_MakeJs.asp
'= 摘    要：生成自定义js类文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-02-28
'====================================================================

Class Cls_MakeJs
	Public Function MakeTxtJs(DataArray,IsShowColumn,IsShowTime,IsShowNew,IsShowTypes,IsShowReview,TitleLen,ContentLen,OpenStyle,ColNum)
		Dim TempStr
		Dim i,j
	
		TempStr="document.write ('<table>');"&chr(10)
		TempStr=TempStr&"document.write ('<tr>');"&chr(10)
		j=1
		For i=0 To UBound(DataArray,2)
			TempStr=TempStr&"document.write ('<td>"
			If CBool(IsShowTypes) Then TempStr=TempStr&EA_Pub.Chk_ArticleType(DataArray(6,i),DataArray(7,i))
			If CBool(IsShowColumn) Then 
				TempStr=TempStr&" [<a href="""&EA_Pub.Cov_ColumnPath(DataArray(1,i),EA_Pub.SysInfo(18))&""" class=""link-Column"">"&DataArray(2,i)&"</a>]&nbsp;"
			End If
			TempStr=TempStr&"<a href="""&EA_Pub.Cov_ArticlePath(DataArray(0,i),DataArray(5,i),EA_Pub.SysInfo(18))&""" title="""&EA_Pub.Base_HTMLFilter(DataArray(3,i))&""""
			If OpenStyle=1 Then TempStr=TempStr&" target=""_blank"""
			TempStr=TempStr&">"
			TempStr=TempStr&EA_Pub.Add_ArticleColor(DataArray(4,i),EA_Pub.Cut_Title(EA_Pub.Base_HTMLFilter(DataArray(3,i)),TitleLen))
			TempStr=TempStr&"</a>"
			If CBool(IsShowNew) Then TempStr=TempStr&EA_Pub.Chk_ArticleTime(DataArray(5,i)) 
			If CBool(IsShowTime) Then TempStr=TempStr&"&nbsp;<span class=""link-Date"">["&Month(DataArray(5,i))&"."&Day(DataArray(5,i))&"]</span>"
			If CBool(IsShowReview) Then TempStr=TempStr&"&nbsp;[<a href="""&SystemFolder&"review.asp?articleid="&DataArray(0,i)&""">评</a>]"
			TempStr=TempStr&"</td>');"&chr(10)
			If j Mod ColNum=0 Then 
				TempStr=TempStr&"document.write ('</tr><tr>');"&chr(10)
				j=1
			Else
				j=j+1
			End If
		Next
		
		If (j-1) Mod ColNum<>0 Then TempStr=TempStr&"document.write ('</tr>');"&chr(10)
		TempStr=TempStr&"document.write ('</table>');"&chr(10)
	
		MakeTxtJs=TempStr
	End Function
	
	Public Function MakeGlsJs(DataArray,IsShowColumn,IsShowTime,IsShowNew,IsShowTypes,IsShowReview,TitleLen,ContentLen,OpenStyle,ColNum,ImgWidth,ImgHeight)
		Dim TempStr
		Dim i,j
	
		TempStr="document.write ('<table>');"&chr(10)
		TempStr=TempStr&"document.write ('<tr>');"&chr(10)
		TempStr=TempStr&"document.write ('<td style=""width: "&ImgWidth+2&"px;"">"
		If DataArray(6,0) Then 
			TempStr=TempStr&"&nbsp;<a href="""&EA_Pub.Cov_ArticlePath(DataArray(0,0),DataArray(5,0),EA_Pub.SysInfo(18))&""" title="""&EA_Pub.Base_HTMLFilter(DataArray(3,i))&""">"
			TempStr=TempStr&"<img src="""&DataArray(8,0)&""""
			If ImgWidth<>0 Then TempStr=TempStr&" width="""&ImgWidth&""""
			If ImgHeight<>0 Then TempStr=TempStr&" height="""&ImgHeight&""""
			TempStr=TempStr&" alt="""&EA_Pub.Base_HTMLFilter(DataArray(3,i))&""" /></a>"
		Else
			TempStr=TempStr&"无图片"
		End If
		TempStr=TempStr&"</td>');"&chr(10)
		
		TempStr=TempStr&"document.write ('<td>');"&chr(10)
			TempStr=TempStr&"document.write ('<table>');"&chr(10)
			TempStr=TempStr&"document.write ('<tr>');"&chr(10)
			j=1
			For i=1 To UBound(DataArray,2)
				TempStr=TempStr&"document.write ('<td>"
				If CBool(IsShowTypes) Then TempStr=TempStr&EA_Pub.Chk_ArticleType(DataArray(6,i),DataArray(7,i))
				If CBool(IsShowColumn) Then 
					TempStr=TempStr&" [<a href="""&EA_Pub.Cov_ColumnPath(DataArray(1,i),EA_Pub.SysInfo(18))&""" class=""link-Column"">"&DataArray(2,i)&"</a>]"
				End If
				TempStr=TempStr&"&nbsp;<a href="""&EA_Pub.Cov_ArticlePath(DataArray(0,i),DataArray(5,i),EA_Pub.SysInfo(18))&""" title="""&EA_Pub.Base_HTMLFilter(DataArray(3,i))&""""
				If OpenStyle=1 Then TempStr=TempStr&" target=""_blank"""
				TempStr=TempStr&">"
				TempStr=TempStr&EA_Pub.Add_ArticleColor(DataArray(4,i),EA_Pub.Cut_Title(EA_Pub.Base_HTMLFilter(DataArray(3,i)),TitleLen))
				TempStr=TempStr&"</a>"
				If CBool(IsShowNew) Then TempStr=TempStr&EA_Pub.Chk_ArticleTime(DataArray(5,i)) 
				If CBool(IsShowTime) Then TempStr=TempStr&"&nbsp;<span class=""link-Date"">["&Month(DataArray(5,i))&"."&Day(DataArray(5,i))&"]</span>"
				If CBool(IsShowReview) Then TempStr=TempStr&"&nbsp;[<a href="""&SystemFolder&"review.asp?articleid="&DataArray(0,i)&""">评</a>]"
				TempStr=TempStr&"</td>');"&chr(10)
				If j Mod ColNum=0 Then 
					TempStr=TempStr&"document.write ('</tr><tr>');"&chr(10)
					j=1
				Else
					j=j+1
				End If
			Next
		
			If (j-1) Mod ColNum<>0 Then TempStr=TempStr&"document.write ('</tr>');"&chr(10)
			TempStr=TempStr&"document.write ('</table>');"&chr(10)
		TempStr=TempStr&"document.write ('</td>');"&chr(10)		
		TempStr=TempStr&"document.write ('</tr>');"&chr(10)
		TempStr=TempStr&"document.write ('</table>');"&chr(10)
		
		MakeGlsJs=TempStr
	End Function
	
	Public Function MakeImgJs(DataArray,IsShowColumn,IsShowTime,IsShowNew,IsShowTypes,IsShowReview,TitleLen,OpenStyle,ColNum,ImgWidth,ImgHeight)
		Dim TempStr
		Dim i,j
	
		TempStr="document.write ('<table>');"&chr(10)
		TempStr=TempStr&"document.write ('<tr>');"&chr(10)
		j=1
		For i=0 To UBound(DataArray,2)
			TempStr=TempStr&"document.write ('<td style=""width: "&ImgWidth+2&"px;"">"
			If DataArray(6,i) Then 
				TempStr=TempStr&"&nbsp;<a href="""&EA_Pub.Cov_ArticlePath(DataArray(0,i),DataArray(5,i),EA_Pub.SysInfo(18))&""" title="""&EA_Pub.Base_HTMLFilter(DataArray(3,i))&""">"
				TempStr=TempStr&"<img src="""&DataArray(8,i)&""""
				If ImgWidth<>0 Then TempStr=TempStr&" width="""&ImgWidth&""""
				If ImgHeight<>0 Then TempStr=TempStr&" height="""&ImgHeight&""""
				TempStr=TempStr&" alt="""&EA_Pub.Base_HTMLFilter(DataArray(3,i))&""" /></a>"
			Else
				TempStr=TempStr&"暂无图片"
			End If
			
			If TitleLen > 0 Then
				TempStr=TempStr&"<br />"
				If CBool(IsShowTypes) Then TempStr=TempStr&EA_Pub.Chk_ArticleType(DataArray(6,i),DataArray(7,i))
				If CBool(IsShowColumn) Then 
					TempStr=TempStr&"[<a href="""&EA_Pub.Cov_ColumnPath(DataArray(1,i),EA_Pub.SysInfo(18))&""" class=""link-Column"">"&DataArray(2,i)&"</a>]&nbsp;"
				End If
				TempStr=TempStr&"<a href="""&EA_Pub.Cov_ArticlePath(DataArray(0,i),DataArray(5,i),EA_Pub.SysInfo(18))&""" title="""&EA_Pub.Base_HTMLFilter(DataArray(3,i))&""""
				If OpenStyle=1 Then TempStr=TempStr&" target=""_blank"""
				TempStr=TempStr&">"
				TempStr=TempStr&EA_Pub.Add_ArticleColor(DataArray(4,i),EA_Pub.Cut_Title(EA_Pub.Base_HTMLFilter(DataArray(3,i)),TitleLen))
				TempStr=TempStr&"</a>"
				If CBool(IsShowNew) Then TempStr=TempStr&EA_Pub.Chk_ArticleTime(DataArray(5,i)) 
				If CBool(IsShowTime) Then TempStr=TempStr&"&nbsp;<span class=""link-Date"">["&Month(DataArray(5,i))&"."&Day(DataArray(5,i))&"]</span>"
				If CBool(IsShowReview) Then TempStr=TempStr&"&nbsp;[<a href="""&SystemFolder&"review.asp?articleid="&DataArray(0,i)&""">评</a>]"
				TempStr=TempStr&"<br><br>"
			End If

			TempStr=TempStr&"</td>');"&chr(10)
			If j Mod ColNum=0 Then 
				TempStr=TempStr&"document.write ('</tr><tr>');"&chr(10)
				j=1
			Else
				j=j+1
			End If
		Next
		
		If (j-1) Mod ColNum<>0 Then TempStr=TempStr&"document.write ('</tr>');"&chr(10)
		TempStr=TempStr&"document.write ('</table>');"&chr(10)
	
		MakeImgJs=TempStr
	End Function
	
	Public Function MakeTxtMoreJs(DataArray,IsShowColumn,IsShowTime,IsShowNew,IsShowTypes,IsShowReview,TitleLen,OpenStyle,ColNum,ImgWidth,ImgHeight)
		Dim TempStr
		Dim i,j
	
		TempStr="document.write ('<table>');"&chr(10)
		TempStr=TempStr&"document.write ('<tr>');"&chr(10)
		j=1
		For i=0 To UBound(DataArray,2)
			TempStr=TempStr&"document.write ('<td>');"&chr(10)
			
				TempStr=TempStr&"document.write ('<table>');"&chr(10)
				TempStr=TempStr&"document.write ('<tr>');"&chr(10)
				TempStr=TempStr&"document.write ('<td colspan=""2"">"
				If CBool(IsShowTypes) Then TempStr=TempStr&EA_Pub.Chk_ArticleType(DataArray(6,i),DataArray(7,i))
				If CBool(IsShowColumn) Then 
					TempStr=TempStr&"[<a href="""&EA_Pub.Cov_ColumnPath(DataArray(1,i),EA_Pub.SysInfo(18))&""" class=""link-Column"">"&DataArray(2,i)&"</a>]&nbsp;"
				End If
				TempStr=TempStr&"<a href="""&EA_Pub.Cov_ArticlePath(DataArray(0,i),DataArray(5,i),EA_Pub.SysInfo(18))&""" title="""&EA_Pub.Base_HTMLFilter(DataArray(3,i))&""""
				If OpenStyle=1 Then TempStr=TempStr&" target=""_blank"""
				TempStr=TempStr&">"
				TempStr=TempStr&EA_Pub.Add_ArticleColor(DataArray(4,i),EA_Pub.Cut_Title(EA_Pub.Base_HTMLFilter(DataArray(3,i)),TitleLen))&"</a>"
				If CBool(IsShowNew) Then TempStr=TempStr&EA_Pub.Chk_ArticleTime(DataArray(5,i)) 
				If CBool(IsShowTime) Then TempStr=TempStr&"&nbsp;<span class=""link-Date"">["&Month(DataArray(5,i))&"."&Day(DataArray(5,i))&"]</span>"
				If CBool(IsShowReview) Then TempStr=TempStr&"&nbsp;[<a href="""&SystemFolder&"review.asp?articleid="&DataArray(0,i)&""">评</a>]"
				TempStr=TempStr&"</td></tr>');"&chr(10)
				
				TempStr=TempStr&"document.write ('<tr>');"&chr(10)
				TempStr=TempStr&"document.write ('<td style=""width: "&ImgWidth+2&"px;"">"
				If DataArray(6,i) Then 
					TempStr=TempStr&"<a href="""&EA_Pub.Cov_ArticlePath(DataArray(0,i),DataArray(5,i),EA_Pub.SysInfo(18))&""" title="""&EA_Pub.Base_HTMLFilter(DataArray(3,i))&""">"
					TempStr=TempStr&"<img src="""&DataArray(8,i)&""""
					If ImgWidth<>0 Then TempStr=TempStr&" width="""&ImgWidth&""""
					If ImgHeight<>0 Then TempStr=TempStr&" height="""&ImgHeight&""""
					TempStr=TempStr&" alt="""&EA_Pub.Base_HTMLFilter(DataArray(3,i))&""" /></a>"
				End If
				TempStr=TempStr&"</td>');"&chr(10)
				TempStr=TempStr&"document.write ('<td>"&EA_Pub.Full_HTMLFilter(DataArray(9,i))&"</td>');"&chr(10)
				TempStr=TempStr&"document.write ('</td></tr></table>');"&chr(10)
				
			TempStr=TempStr&"document.write ('</td>');"&chr(10)
			
			If j Mod ColNum=0 Then 
				TempStr=TempStr&"document.write ('</tr><tr>');"&chr(10)
				j=1
			Else
				j=j+1
			End If
		Next
		
		If (j-1) Mod ColNum<>0 Then TempStr=TempStr&"document.write ('</tr>');"&chr(10)
		TempStr=TempStr&"document.write ('</table>');"&chr(10)
	
		MakeTxtMoreJs=TempStr
	End Function
	
End Class
%>