<!--#Include File="comm/inc.asp" -->
<%
'====================================================================
'= Team Elite - EliteCMS
'= Copyright (c) 2004 - 2008 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：/Manager/Admin_MailOut.asp
'= 摘    要：后台-导出邮件文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-03-18
'====================================================================

Call EA_Manager.Chk_IsMaster

If Not EA_Manager.Chk_Power(Admin_Power,"54") Then 
	ErrMsg=str_Comm_NotAccess
	Call EA_Manager.Error(1)
End If

Dim Atcion
Atcion=Request.Form("way")

Select Case LCase(Atcion)
Case "todatabase"
	Call OutToDataBase
Case "totxt"
	Call OutToTxt
Case Else
	Call Main
End Select
Call EA_Pub.Close_Obj
Set EA_Pub=Nothing

Sub Main
	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_MailOut_Help",str_MailOut_Help)

	Call EA_M_XML.AppElements("Language_MailOut_OutToDataBase",str_MailOut_OutToDataBase)
	Call EA_M_XML.AppElements("Language_MailOut_OutToTxt",str_MailOut_OutToTxt)
	Call EA_M_XML.AppElements("Language_MailOut_ExportType",str_MailOut_ExportType)
	Call EA_M_XML.AppElements("Language_MailOut_ExportList",str_MailOut_ExportList)
	Call EA_M_XML.AppElements("Language_MailOut_Now",str_MailOut_Now)

	Call EA_M_XML.AppElements("btnSubmit",str_MailOut_Now)

	Page = EA_M_XML.make("","",0)
	Call EA_M_XML.Out(Page)
End Sub

Sub OutToDataBase
	On Error Resume Next
	
	Dim TConn,TempCount,TStr
	Dim ExportTotal

	FoundErr = False

	Set TConn = Server.CreateObject("ADODB.Connection")
	TStr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("databackup/nb_maillist.mdb")
	TConn.Open TStr

	If Err Then 
		Response.Write "-1"
		Response.End
	End If

	SQL="Select Reg_Name,Email From [NB_User] Where Email Like '%@%'"
	Set Rs=Conn.Execute(SQL)

	Do While Not Rs.eof
		TConn.Execute("insert into [NB_MailList]([UserName],EMailAddress) values ('"&rs(0)&"','"&rs(1)&"')")
		ExportTotal = ExportTotal + 1
		rs.MoveNext
	Loop
	tConn.Close
	Set tConn = Nothing
	Rs.Close
	Set Rs=Nothing
	
	Response.Write "1"
	Response.End
End Sub

Sub OutToTxt
	Dim file,filepath,writefile
	Dim ExportTotal

	Set file = CreateObject("Scripting.FileSystemObject")
	filepath=Server.MapPath("mail.txt")
	Set Writefile = file.CreateTextFile(filepath,true)

	SQL="Select Reg_Name,Email From [NB_User] Where Email Like '%@%'"
	Set Rs=Conn.Execute(SQL)

	do while not rs.eof
		Writefile.WriteLine rs(0)&Chr(9)&rs(1)
		ExportTotal = ExportTotal + 1
		rs.movenext
	loop
	Writefile.close
	Set file=Nothing
	
	Rs.close
	Set Rs=Nothing
	
	Response.Write "0"
	Response.End
End Sub
%>