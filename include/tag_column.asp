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
'= 最后日期：2008-02-28
'====================================================================

	Function Load_Column(Parameter)
		Dim ColumnId, ColumnInfo
		Dim Url
		
		ColumnId	= Parameter(0)
		ColumnInfo	= EA_DBO.Get_Column_Info(ColumnId)

		Url = "<a href='" & EA_Pub.Cov_ColumnPath(ColumnId,EA_Pub.SysInfo(18)) & "'>"
		Url = Url & ColumnInfo(0,0) & "</a>"

		Load_Column = Url
	End Function

%>