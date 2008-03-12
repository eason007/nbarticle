
<!--#Include File="comm/inc.asp" -->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：/Manager/Admin_UpFile.asp
'= 摘    要：后台-上传文件管理文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2006-11-12
'====================================================================

Call EA_Manager.Chk_IsMaster

If Not EA_Manager.Chk_Power(Admin_Power,"61") Then 
	ErrMsg=str_Comm_NotAccess
	Call EA_Manager.Error(1)
End If

Dim  Action,sFor(31,1)
Action=Request.Form("action")

Select Case LCase(Action)
Case Else
	Call Main()
End Select
Set Rs=Nothing
Call EA_Pub.Close_Obj
Set EA_Pub=Nothing
	
Sub Main()
	Dim path
	Dim mainpath
	Dim requestpath
	Dim ParentFolder
	Dim objFSO
	Dim uploadfolder,uploadChildFolder
	Dim uploadfiles
	Dim upname
	Dim UpFolder
	Dim i
	Dim pagesize, page, filenum, pagenum
	Dim Tmp
	Dim ListName(6),ListValue()

	sFor(0,0)="txt":sFor(0,1)="txt"
	sFor(1,0)="chm":sFor(1,1)="chm"
	sFor(2,0)="hlp":sFor(2,1)="chm"
	sFor(3,0)="doc":sFor(3,1)="doc"
	sFor(4,0)="pdf":sFor(4,1)="pdf"
	sFor(5,0)="gif":sFor(5,1)="gif"
	sFor(6,0)="jpg":sFor(6,1)="jpg"
	sFor(7,0)="png":sFor(7,1)="png"
	sFor(8,0)="bmp":sFor(8,1)="bmp"
	sFor(9,0)="asp":sFor(9,1)="asp"
	sFor(10,0)="jsp":sFor(10,1)="asp"
	sFor(11,0)="js" :sFor(11,1)="asp"
	sFor(12,0)="htm":sFor(12,1)="html"
	sFor(13,0)="html":sFor(13,1)="html"
	sFor(14,0)="shtml":sFor(14,1)="html"
	sFor(15,0)="zip":sFor(15,1)="zip"
	sFor(16,0)="rar":sFor(16,1)="rar"
	sFor(17,0)="exe":sFor(17,1)="exe"
	sFor(18,0)="avi":sFor(18,1)="avi"
	sFor(19,0)="mpg":sFor(19,1)="mpg"
	sFor(20,0)="ra" :sFor(20,1)="ra"
	sFor(21,0)="ram":sFor(21,1)="ra"
	sFor(22,0)="mid":sFor(22,1)="mid"
	sFor(23,0)="wav":sFor(23,1)="wav"
	sFor(24,0)="mp3":sFor(24,1)="mp3"
	sFor(25,0)="asf":sFor(25,1)="asf"
	sFor(26,0)="php":sFor(26,1)="aspx"
	sFor(27,0)="php3":sFor(27,1)="aspx"
	sFor(28,0)="aspx":sFor(28,1)="aspx"
	sFor(29,0)="xls":sFor(29,1)="xls"
	sFor(30,0)="mdb":sFor(30,1)="mdb"

	path="../UserFiles"
	mainpath=server.MapPath (path)
	requestpath=Request.Form("path")

	if requestpath<>"" then	
		requestpath=Replace(requestpath,"..","")
		path=path&requestpath
	end if

	pagesize=10
	page=request.querystring("page")
	if page="" or not isnumeric(page) then
		page=1
	else
		page=int(page)
	end if
	on error resume next

	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_UpFile_Help",str_UpFile_Help)

	Call EA_M_XML.AppElements("Language_UpFile_Type",str_UpFile_Type)
	Call EA_M_XML.AppElements("Language_UpFile_FileName",str_UpFile_FileName)
	Call EA_M_XML.AppElements("Language_UpFile_Size",str_UpFile_Size)
	Call EA_M_XML.AppElements("Language_UpFile_LastTime",str_UpFile_LastTime)
	Call EA_M_XML.AppElements("Language_UpFile_UploadTime",str_UpFile_UploadTime)
	Call EA_M_XML.AppElements("Language_UpFile_CurrentPath",str_UpFile_CurrentPath)
	Call EA_M_XML.AppElements("Path",path)
	Call EA_M_XML.AppElements("RequsetPath",request.form("path"))

	Call EA_M_XML.AppElements("Comm_Del_Operation",str_Comm_Del_Operation)

	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	if request("filename")<>"" then
	   if objFSO.fileExists(Server.MapPath(""&path&"\"&request("filename"))) then
			objFSO.DeleteFile(Server.MapPath(""&path&"\"&request("filename")))
	   end if
	end if

	Set uploadFolder=objFSO.GetFolder(Server.MapPath(""&path&"\"))

	if err.number<>0 then
		ErrMsg = Err.Description
		Call EA_Manager.Error(1)
	end if

	If requestpath <> "" Then
		ParentFolder=objFSO.GetParentFolderName(Server.MapPath(""&path&"\"))
		ParentFolder=Replace(ParentFolder,mainpath,"",1,-1,1)
		ParentFolder=Replace(ParentFolder,"\","/",1,-1,1)
	End If

	Call EA_M_XML.AppElements("ParentFolder",ParentFolder & "")

	ListName(0) = "checkbox"
	ListName(1) = "ID"
	ListName(2) = "Type"
	ListName(3) = "FileName"
	ListName(4) = "Size"
	ListName(5) = "LastTime"
	ListName(6) = "UploadTime"

	Tmp = -1

	Set uploadChildFolder=uploadFolder.SubFolders 

	For Each Upname In uploadChildFolder
		Tmp = Tmp + 1

		ReDim Preserve ListValue(6,Tmp)

		ListValue(0,Tmp) = "checkbox"
		ListValue(1,Tmp) = Tmp
		ListValue(2,Tmp) = "[Dir]"
		ListValue(3,Tmp) = "<a href='javascript:vod();' onclick='javascript:objTable.dateVal=""path=" & requestpath & "/" & Upname.name & """;objTable.start();'>" & Upname.name & "</a>"
	Next


	Set uploadFiles=uploadFolder.Files
	filenum=uploadfiles.count
	i=0

	For Each Upname In uploadFiles
		i=i+1
		if i>(page-1)*pagesize and i<=page*pagesize then
			Tmp = Tmp + 1

			ReDim Preserve ListValue(6,Tmp)

			ListValue(0,Tmp) = "checkbox"
			ListValue(1,Tmp) = Tmp
			ListValue(2,Tmp) = "<img src=""images/files/" & procGetFormat(upname.name) & ".gif"" border=0>"
			ListValue(3,Tmp) = "<a href='../userfiles" & requestpath & "/" & Upname.name & "' target='_blank'>" & Upname.name & "</a>"
			ListValue(4,Tmp) = showfilesize(upname.size)
			ListValue(5,Tmp) = upname.datelastaccessed
			ListValue(6,Tmp) = upname.datecreated
		elseif i>page*pagesize then
			exit for
		end if
	next


	set uploadFolder=nothing
	set uploadChildFolder=nothing
	set uploadFiles=nothing
	set objFSO=nothing


	If Tmp = -1 Then
		Page = EA_M_XML.make("","",0)
	Else
		Page = EA_M_XML.make(ListName,ListValue,Tmp+1)
	End If

	Call EA_M_XML.Out(Page)
end sub

function procGetFormat(sName)
	Dim  i,str
	
	procGetFormat=0
	
	if instrRev(sName,".")=0 then exit function
	
	str=LCase(Mid(sName,instrRev(sName,".")+1))
	
	for i=0 to uBound(sFor,1)
		if str=sFor(i,0) then 
			procGetFormat=sFor(i,1)
			exit for
		end if
	next
end function

function showfilesize(size)
	if size>1024 then
		size=(Size/1024)
		showfilesize=formatnumber(size,2) & "&nbsp;KB"
	end if
	if size>1024 then
		size=(size/1024)
		showfilesize=formatnumber(size,2) & "&nbsp;MB"
	end if
	if size>1024 then
 	   size=(size/1024)
 	   showfilesize=formatnumber(size,2) & "&nbsp;GB"
 	end if 
end function
%>