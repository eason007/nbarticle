<!--#Include File="../conn.asp" -->
<!--#Include File="comm/inc.asp" -->
<!--#include file="../include/_cls_teamplate.asp"-->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：/Manager/Admin_Data.asp
'= 摘    要：后台-数据库管理文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2006-11-12
'====================================================================

Call EA_Manager.Chk_IsMaster

Dim Action
Action=Request.Form("action")

Select Case LCase(Action)
Case "backupdata"
	If iDataBaseType=0 Then 
		If Not EA_Manager.Chk_Power(Admin_Power,"51") Then 
			ErrMsg=str_Comm_NotAccess
			Call EA_Manager.Error(1)
		End If
		Call BackUpDataBase
	Else
		Call SQLUserReadme
	End If
Case "backup"
	If iDataBaseType=0 Then 
		If Not EA_Manager.Chk_Power(Admin_Power,"51") Then 
			ErrMsg=str_Comm_NotAccess
			Call EA_Manager.Error(1)
		End If
		Call EA_Manager.BackUpAccessDataBase()
	Else
		Call SQLUserReadme
	End If
Case "execute"
	If Not EA_Manager.Chk_Power(Admin_Power,"52") Then 
		ErrMsg=str_Comm_NotAccess
		Call EA_Manager.Error(1)
	End If
	Call Execute()
Case "exe"
	If Not EA_Manager.Chk_Power(Admin_Power,"52") Then 
		ErrMsg=str_Comm_NotAccess
		Call EA_Manager.Error(1)
	End If
	Err.Clear 
	On Error Resume Next
	SQL=Request.Form("text")
	Conn.Execute(SQL)
	If Err.number=0 Then 
		Response.Write "1"
	Else
		Response.Write Err.Description 
	End If
Case "diskview"
	If Not EA_Manager.Chk_Power(Admin_Power,"53") Then 
		ErrMsg=str_Comm_NotAccess
		Call EA_Manager.Error(1)
	End If
	Call DiskView
End Select
Call EA_Pub.Close_Obj
Set EA_Pub=Nothing

Sub BackUpDataBase
	Dim Template
	Set Template=New cls_NEW_TEMPLATE

	PageContent=Template.LoadTemplate("admin_data_backup.htm")

	Template.SetVariable "Language_OperationNotice",str_OperationNotice,PageContent
	Template.SetVariable "Language_Data_Backup_Help",str_Data_Backup_Help,PageContent

	Template.SetVariable "Language_Data_Backup_Title",str_Data_Backup_Title,PageContent
	Template.SetVariable "Language_Data_Backup_BackupFolder",str_Data_Backup_BackupFolder,PageContent
	Template.SetVariable "Language_Data_Backup_BackupFileName",str_Data_Backup_BackupFileName,PageContent

	Template.SetVariable "Language_Comm_Save_Button",str_Comm_Save_Button,PageContent

	Template.SetVariable "Date",Date(),PageContent

	Template.BaseReplace PageContent
	Template.OutStr PageContent
End Sub

Sub Execute
	Call EA_M_XML.AppElements("Language_OperationNotice",str_OperationNotice)
	Call EA_M_XML.AppElements("Language_Data_Execute_Help",str_Data_Execute_Help)

	Call EA_M_XML.AppElements("Language_Data_Execute_Title",str_Data_Execute_Title)
	Call EA_M_XML.AppElements("btnSubmit",str_Comm_Submit_Button)

	Page = EA_M_XML.make("","",0)
	Call EA_M_XML.Out(Page)
End Sub

Sub DiskView()
	On Error resume next

	Call EA_M_XML.AppElements("Language_Data_SpaceView_Title",str_Data_SpaceView_Title)
	Call EA_M_XML.AppElements("Language_Data_SpaceView_BackupFolder",str_Data_SpaceView_BackupFolder)
	Call EA_M_XML.AppElements("Language_Data_SpaceView_DatabaseFolder",str_Data_SpaceView_DatabaseFolder)
	Call EA_M_XML.AppElements("Language_Data_SpaceView_UploadFolder",str_Data_SpaceView_UploadFolder)
	Call EA_M_XML.AppElements("Language_Data_SpaceView_ManagerFolder",str_Data_SpaceView_ManagerFolder)
	Call EA_M_XML.AppElements("Language_Data_SpaceView_HTMLListFolder",str_Data_SpaceView_HTMLListFolder)
	Call EA_M_XML.AppElements("Language_Data_SpaceView_HTMLViewFolder",str_Data_SpaceView_HTMLViewFolder)
	Call EA_M_XML.AppElements("Language_Data_SpaceView_All",str_Data_SpaceView_All)

	Call EA_M_XML.AppInfo("hPic_1",drawbar("databackup"))
	Call EA_M_XML.AppInfo("hPic_2",drawbar("../depot"))
	Call EA_M_XML.AppInfo("hPic_3",drawbar("../userfiles"))
	Call EA_M_XML.AppInfo("hPic_4",drawbar("./"))
	Call EA_M_XML.AppInfo("hPic_5",drawbar("../articlelist"))
	Call EA_M_XML.AppInfo("hPic_6",drawbar("../articleview"))

	Call EA_M_XML.AppElements("InfoBackup",showSpaceinfo("databackup"))
	Call EA_M_XML.AppElements("InfoDepot",showSpaceinfo("../depot"))
	Call EA_M_XML.AppElements("InfoUserFiles",showSpaceinfo("../userfiles"))
	Call EA_M_XML.AppElements("InfoThis",showSpaceinfo("./"))
	Call EA_M_XML.AppElements("InfoHTMLList",showSpaceinfo("../articlelist"))
	Call EA_M_XML.AppElements("InfoHTMLView",showSpaceinfo("../articleview"))

	Call EA_M_XML.AppElements("InfoAll",showspecialspaceinfo("All"))
	
	Page = EA_M_XML.make("","",0)
	Call EA_M_XML.Out(Page)
End Sub

'=====================系统空间参数=========================
	Function ShowSpaceInfo(drvpath)
 		dim fso,d,size,showsize
 		
 		set fso=server.createobject("scriptin" & "g.f" & "ilesystemobject")

 		drvpath=server.mappath(drvpath)
 		set d=fso.getfolder(drvpath)
 		size=d.size
 		showsize=size & "&nbsp;Byte"

 		Set fso=Nothing
 		
 		if size>1024 then
 		   size=(Size/1024)
 		   showsize=formatnumber(size,2) & "&nbsp;KB"
 		end if
 		if size>1024 then
 		   size=(size/1024)
 		   showsize=formatnumber(size,2) & "&nbsp;MB"
 		end if
 		if size>1024 then
 		   size=(size/1024)
 		   showsize=formatnumber(size,2) & "&nbsp;GB"
 		end if

 		ShowSpaceInfo = "<font face=verdana>" & showsize & "</font>"
 	End Function	
 	
 	Function Showspecialspaceinfo(method)
 		dim fso,d,fc,f1,size,showsize,drvpath 
 				
 		set fso=server.createobject("scriptin" & "g.f" & "ilesystemobject")

 		drvpath=server.mappath("./")
 		drvpath=left(drvpath,(instrrev(drvpath,"\")-1))

 		set d=fso.getfolder(drvpath)
 		
 		if method="All" then
 			size=d.size
 		elseif method="Program" then
 			drvpath=server.MapPath("../editor/UploadFile")
 			'Response.Write drvpath
 			set d=fso.getfolder(drvpath)
 			size=d.size
 		end if

		Set fso=Nothing
 		
 		showsize=size & "&nbsp;Byte"
 		if size>1024 then
 		   size=(Size/1024)
 		   showsize=formatnumber(size,2) & "&nbsp;KB"
 		end if
 		if size>1024 then
 		   size=(size/1024)
 		   showsize=formatnumber(size,2) & "&nbsp;MB"
 		end if
 		if size>1024 then
 		   size=(size/1024)
 		   showsize=formatnumber(size,2) & "&nbsp;GB"
 		end if 
 		Showspecialspaceinfo = "<font face=verdana>" & showsize & "</font>"
 	end Function
 	
 	Function Drawbar(drvpath)
 		dim fso,drvpathroot,d,size,totalsize,barsize
 		set fso=server.createobject("scripting.filesystemobject")
 		drvpathroot=server.mappath("../")
 		
 		set d=fso.getfolder(drvpathroot)
 		totalsize=d.size
 		
 		drvpath=server.mappath(drvpath)
 		set d=fso.getfolder(drvpath)
 		size=d.size
 		'Response.Write "["&size&"]"
 
 		barsize=cint((size/totalsize)*400)
 		Drawbar=barsize
 		
 		Set fso=Nothing
 	End Function 
 	
 	Function Drawspecialbar()
 		dim fso,drvpathroot,d,fc,f1,size,totalsize,barsize
 		set fso=server.createobject("scripting.filesystemobject")
 		drvpathroot=server.mappath("../")
 		drvpathroot=left(drvpathroot,(instrrev(drvpathroot,"\")-1))
 		set d=fso.getfolder(drvpathroot)
 		totalsize=d.size
 		
 		set fc=d.files
 		for each f1 in fc
 			size=size+f1.size
 		next
 		
 		barsize=cint((size/totalsize)*400)
 		Drawspecialbar=barsize
 		Set fso=Nothing
 	End Function 

	Sub SQLUserReadme()
	%>
<style type="text/css">
body {
	FONT-SIZE: 12px;
	BACKGROUND: #efefef;
	padding: 0;
	margin: 3px 3px 3px 3px;
}
th {
	BACKGROUND: #c5d7e2;
	height: 25px;
}
</style>
<table border="0" cellspacing="0" cellpadding="0" align=center width="95%"> 
  <TR>
      <TD><TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
          <TBODY>
            <TR>
              <TD width=3 height=3><IMG height=3 src="images/O_angle_up_left.gif" width=3></TD>
              <TD width="100%" bgColor=#6795b4></TD>
              <TD width=3><IMG height=3 src="images/O_angle_up_right.gif" width=3></TD>
            </TR>
          </TBODY>
        </TABLE></TD>
    </TR>
	<TR>
      <TD bgColor=#6795b4 height=25>&nbsp;&nbsp;<font color=ffffff>SQL数据库数据处理说明</font></TD>
    </TR>
  <tr>
    <td class=form_border> <blockquote> <B>一、备份数据库</B> <BR> 
        <BR> 
        1、打开SQL企业管理器，在控制台根目录中依次点开Microsoft SQL Server<BR> 
        2、SQL Server组-->双击打开你的服务器-->双击打开数据库目录<BR> 
        3、选择你的数据库名称（如论坛数据库Forum）-->然后点上面菜单中的工具-->选择备份数据库<BR> 
        4、备份选项选择完全备份，目的中的备份到如果原来有路径和名称则选中名称点删除，然后点添加，如果原来没有路径和名称则直接选择添加，接着指定路径和文件名，指定后点确定返回备份窗口，接着点确定进行备份 <BR> 
        <BR> 
        <B>二、还原数据库</B><BR> 
        <BR> 
        1、打开SQL企业管理器，在控制台根目录中依次点开Microsoft SQL Server<BR> 
        2、SQL Server组-->双击打开你的服务器-->点图标栏的新建数据库图标，新建数据库的名字自行取<BR> 
        3、点击新建好的数据库名称（如论坛数据库Forum）-->然后点上面菜单中的工具-->选择恢复数据库<BR> 
        4、在弹出来的窗口中的还原选项中选择从设备-->点选择设备-->点添加-->然后选择你的备份文件名-->添加后点确定返回，这时候设备栏应该出现您刚才选择的数据库备份文件名，备份号默认为1（如果您对同一个文件做过多次备份，可以点击备份号旁边的查看内容，在复选框中选择最新的一次备份后点确定）-->然后点击上方常规旁边的选项按钮<BR> 
        5、在出现的窗口中选择在现有数据库上强制还原，以及在恢复完成状态中选择使数据库可以继续运行但无法还原其它事务日志的选项。在窗口的中间部位的将数据库文件还原为这里要按照你SQL的安装进行设置（也可以指定自己的目录），逻辑文件名不需要改动，移至物理文件名要根据你所恢复的机器情况做改动，如您的SQL数据库装在D:\Program Files\Microsoft SQL Server\MSSQL\Data，那么就按照您恢复机器的目录进行相关改动改动，并且最后的文件名最好改成您当前的数据库名（如原来是bbs_data.mdf，现在的数据库是forum，就改成forum_data.mdf），日志和数据文件都要按照这样的方式做相关的改动（日志的文件名是*_log.ldf结尾的），这里的恢复目录您可以自由设置，前提是该目录必须存在（如您可以指定d:\sqldata\bbs_data.mdf或者d:\sqldata\bbs_log.ldf），否则恢复将报错<BR> 
        6、修改完成后，点击下面的确定进行恢复，这时会出现一个进度条，提示恢复的进度，恢复完成后系统会自动提示成功，如中间提示报错，请记录下相关的错误内容并询问对SQL操作比较熟悉的人员，一般的错误无非是目录错误或者文件名重复或者文件名错误或者空间不够或者数据库正在使用中的错误，数据库正在使用的错误您可以尝试关闭所有关于SQL窗口然后重新打开进行恢复操作，如果还提示正在使用的错误可以将SQL服务停止然后重起看看，至于上述其它的错误一般都能按照错误内容做相应改动后即可恢复<BR> 
        <BR> 
        <B>三、收缩数据库</B><BR> 
        <BR> 
        一般情况下，SQL数据库的收缩并不能很大程度上减小数据库大小，其主要作用是收缩日志大小，应当定期进行此操作以免数据库日志过大<BR> 
        1、设置数据库模式为简单模式：打开SQL企业管理器，在控制台根目录中依次点开Microsoft SQL Server-->SQL Server组-->双击打开你的服务器-->双击打开数据库目录-->选择你的数据库名称（如论坛数据库Forum）-->然后点击右键选择属性-->选择选项-->在故障还原的模式中选择“简单”，然后按确定保存<BR> 
        2、在当前数据库上点右键，看所有任务中的收缩数据库，一般里面的默认设置不用调整，直接点确定<BR> 
        3、<font color=blue>收缩数据库完成后，建议将您的数据库属性重新设置为标准模式，操作方法同第一点，因为日志在一些异常情况下往往是恢复数据库的重要依据</font> <BR> 
        <BR> 
        <B>四、设定每日自动备份数据库</B><BR> 
        <BR> 
        <font color=red>强烈建议有条件的用户进行此操作！</font><BR> 
        1、打开企业管理器，在控制台根目录中依次点开Microsoft SQL Server-->SQL Server组-->双击打开你的服务器<BR> 
        2、然后点上面菜单中的工具-->选择数据库维护计划器<BR> 
        3、下一步选择要进行自动备份的数据-->下一步更新数据优化信息，这里一般不用做选择-->下一步检查数据完整性，也一般不选择<BR> 
        4、下一步指定数据库维护计划，默认的是1周备份一次，点击更改选择每天备份后点确定<BR> 
        5、下一步指定备份的磁盘目录，选择指定目录，如您可以在D盘新建一个目录如：d:\databak，然后在这里选择使用此目录，如果您的数据库比较多最好选择为每个数据库建立子目录，然后选择删除早于多少天前的备份，一般设定4－7天，这看您的具体备份要求，备份文件扩展名一般都是bak就用默认的<BR> 
        6、下一步指定事务日志备份计划，看您的需要做选择-->下一步要生成的报表，一般不做选择-->下一步维护计划历史记录，最好用默认的选项-->下一步完成<BR> 
        7、完成后系统很可能会提示Sql Server Agent服务未启动，先点确定完成计划设定，然后找到桌面最右边状态栏中的SQL绿色图标，双击点开，在服务中选择Sql Server Agent，然后点击运行箭头，选上下方的当启动OS时自动启动服务<BR> 
        8、这个时候数据库计划已经成功的运行了，他将按照您上面的设置进行自动备份 <BR> 
        <BR> 
        修改计划：<BR> 
        1、打开企业管理器，在控制台根目录中依次点开Microsoft SQL Server-->SQL Server组-->双击打开你的服务器-->管理-->数据库维护计划-->打开后可看到你设定的计划，可以进行修改或者删除操作 <BR> 
        <BR> 
        <B>五、数据的转移（新建数据库或转移服务器）</B><BR> 
        <BR> 
        一般情况下，最好使用备份和还原操作来进行转移数据，在特殊情况下，可以用导入导出的方式进行转移，这里介绍的就是导入导出方式，导入导出方式转移数据一个作用就是可以在收缩数据库无效的情况下用来减小（收缩）数据库的大小，本操作默认为您对SQL的操作有一定的了解，如果对其中的部分操作不理解，可以咨询动网相关人员或者查询网上资料<BR> 
        1、将原数据库的所有表、存储过程导出成一个SQL文件，导出的时候注意在选项中选择编写索引脚本和编写主键、外键、默认值和检查约束脚本选项<BR> 
        2、新建数据库，对新建数据库执行第一步中所建立的SQL文件<BR> 
        3、用SQL的导入导出方式，对新数据库导入原数据库中的所有表内容<BR> 
      </blockquote></td> 
  </tr> 
  <TR>
      <TD><TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
          <TBODY>
            <TR>
              <TD width=3><IMG height=3 src="images/O_angle_down_left.gif" width=3></TD>
              <TD width="100%" background=images/O_angle_bottom.gif height=3></TD>
              <TD width=3><IMG height=3 src="images/O_angle_down_right.gif" width=3></TD>
            </TR>
          </TBODY>
        </TABLE></TD>
    </TR>
</table> 
<%End Sub%> 
