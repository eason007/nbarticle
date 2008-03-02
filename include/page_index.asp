<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：page_index.asp
'= 摘    要：首页文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-03-03
'====================================================================

Class page_Index
	Public Function Make ()
		Dim PageContent

		PageContent		= EA_Temp.Load_Template(0, 0)

		EA_Temp.Title	= EA_Pub.SysInfo(0) & " - " & SysMsg(10)
		EA_Temp.Nav		= "<a href=""<!--Page.Path-->""><strong>" & EA_Pub.SysInfo(0) & "</strong></a> - " & SysMsg(10)

		PageContent		= EA_Temp.Replace_PublicTag(PageContent)

		Make = PageContent
	End Function
End Class
%>