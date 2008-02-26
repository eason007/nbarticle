<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：/Manager/language_files/Comm.asp
'= 摘    要：后台-语言包-公用文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-02-20
'====================================================================

Dim str_LeftMenu(8,11)

str_LeftMenu(0,0)="系统管理"
str_LeftMenu(0,1)="<a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_config.htm\');"">系统设定</a>"
str_LeftMenu(0,2)="<a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_friend_list.htm\');"">联盟管理</a> | <a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_friend_option.htm\', \'pageAdd\');"">添加</a>"
str_LeftMenu(0,3)="<a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_placard_list.htm\');"">公告管理</a> | <a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_placard_option.htm\', \'pageAdd\');"">添加</a>"
str_LeftMenu(0,4)="<a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_vote_list.htm\');"">投票管理</a> | <a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_vote_option.htm\', \'pageAdd\');"">添加</a>"
str_LeftMenu(0,5)="<a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_insidelink_list.htm\');"">站内连接</a> | <a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_insidelink_option.htm\', \'pageAdd\');"">添加</a>"
str_LeftMenu(0,6)="<a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_adsense_list.htm\');"">广告管理</a> | <a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_adsense_option.htm\', \'pageAdd\');"">添加</a>"
str_LeftMenu(0,7)="<a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/reload.htm\');"">数据更新</a>"

str_LeftMenu(1,0)="内容管理"
str_LeftMenu(1,1)="<a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_column_list.htm\');"">栏目管理</a> | <a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_column_option.htm\', \'pageAdd\');"">添加</a>"
str_LeftMenu(1,2)="<a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_content_list.htm\');"">文章管理</a> | <a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_content_option.htm\', \'pageAdd\');"">添加</a>"
str_LeftMenu(1,3)="<a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_columnbath_move.htm\');"">文章批量移动</a> | <a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_columnbath_del.htm\');"">删除</a>"
str_LeftMenu(1,4)="<a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_review_list.htm\');"">评论管理</a>"
str_LeftMenu(1,5)="<a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_arttemplate_list.htm\');"">文章模版管理</a> | <a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_arttemplate_option.htm\', \'pageAdd\');"">添加</a>"

str_LeftMenu(2,0)="会员管理"
str_LeftMenu(2,1)="<a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_usergroup_list.htm\');"">用户组管理</a> | <a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_usergroup_option.htm\', \'pageAdd\');"">添加</a>"
str_LeftMenu(2,2)="<a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_member_list.htm\');"">用户管理</a>"

str_LeftMenu(3,0)="安全管理"
str_LeftMenu(3,1)="<a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_master_list.htm\');"">管理员管理</a> | <a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./admin_master.asp\',\'\',\'\',false,\'action=add\');"">添加</a>"
str_LeftMenu(3,2)="<a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_ip_list.htm\');"">IP管理</a> | <a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_ip_option.htm\', \'pageAdd\');"">添加</a>"

str_LeftMenu(4,0)="风格管理"
str_LeftMenu(4,1)="<a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_makeindex_main.htm\');"">生成首页</a>"
str_LeftMenu(4,2)="<a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_makelist_main.htm\');"">生成列表</a>"
str_LeftMenu(4,3)="<a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_makeview_main.htm\');"">生成内容页</a>"
str_LeftMenu(4,4)="<a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_theme_list.htm\');"">风格管理</a>"
str_LeftMenu(4,5)="<a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_loadskin_main.htm\');"">导出风格</a> | <a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_loadskin_input.htm\');"">导入</a>"
str_LeftMenu(4,6)="<a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_makejs_list.htm\');"">Js文件管理</a> | <a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_makejs_option.htm\', \'pageAdd\');"">添加</a>"

str_LeftMenu(5,0)="数据库管理"
str_LeftMenu(5,1)="<a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./admin_data.asp\',\'\',\'\',false,\'action=backupdata\');"">备份数据库</a>"
str_LeftMenu(5,2)="<a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_data_execute.htm\');"">数据库高级管理</a>"
str_LeftMenu(5,3)="<a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_data_diskview.htm\');"">系统空间占用</a>"
str_LeftMenu(5,4)="<a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_mailout.htm\');"">导出邮件地址列表</a>"

str_LeftMenu(6,0)="上传文件管理"
str_LeftMenu(6,1)="<a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_upfile_list.htm\');"">上传文件管理</a>"

str_LeftMenu(7,0)="接口管理"
str_LeftMenu(7,1)="<a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_interface_list.htm\');"">外部接口管理</a> | <a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_interface_option.htm\', \'pageAdd\');"">添加</a>"

Const str_Comm_Reduce_Input="缩小文本框"
Const str_Comm_Zoom_Input="放大文本框"
Const str_Comm_Yes="是"
Const str_Comm_No="否"
Const str_Comm_Bar_Operation="操作"
Const str_Comm_ListEmpty="暂无内容"
Const str_Comm_State_Pass="审核通过"
Const str_Comm_State_NoPass="审核不通过"
Const str_Comm_Add_Operation="添加"
Const str_Comm_View_Operation="查看"
Const str_Comm_Edit_Operation="编辑"
Const str_Comm_Del_Operation="删除"
Const str_Comm_Alert_Del_Operation="确定删除吗?"
Const str_Comm_Alert_Pass_Operation="确定通过审核吗?"
Const str_Comm_Submit_Button="确定"
Const str_Comm_Save_Button="保 存"
Const str_Comm_Reset_Button="清 除"
Const str_Comm_Return_Button="返回"
Const str_Comm_Alert_Enabled="确定启用?"
Const str_Comm_Alert_Disabled="确定禁用?"
Const str_Comm_Enabled="启用"
Const str_Comm_Disabled="禁用"
Const str_Comm_Today="今日"
Const str_Comm_Week="今周"
Const str_Comm_Month="今月"
Const str_Comm_RecycleBin="回收站"
Const str_Comm_Select="请选择"
Const str_Comm_SelectAll="全选/反选"
Const str_Comm_Back="上一步"
Const str_Comm_Next="下一步"
Const str_Comm_HelpAlt="点击开关帮助提示"
Const str_Comm_NotAccess="对不起,你没有进行$1操作的权限."
Const str_Comm_AllColumn="所有栏目"

Const str_PassDataError="传递数据时发生错误,操作已暂停."
Const str_BatchOperationMessageForError="有 $1 篇文章因没有权限而没有处理."
Const str_BatchOperationMessageForSucess="所有操作已经完成."

'admin_left.asp
Const str_Thruway="快速通道"

'admin_main.asp
Const str_SystemInformation="系统信息"
Const str_CurrentMasterName="当前管理员："
Const str_SystemStat="系统统计"
Const str_Column="栏目"
Const str_Article="文章"
Const str_AuditArticle="待审核文章"
Const str_RegUser="注册用户"
Const str_Review="评论"
Const str_SystemOwner="系统使用者"
Const str_SystemVersion="系统版本"
Const str_ServerWindow="服务器操作系统"
Const str_ScripEngine="脚本引擎"
Const str_SiteFolderPath="站点路径"
Const str_FSOEnable="FSO支持"
Const str_ADOEnable="ADO支持"
Const str_JMailEnable="Jmail支持"
Const str_CDONTSEnable="CDONTS支持"
Const str_MoreSystemInformation="更多服务器信息"
Const str_SystemManagerShortcut="系统管理快捷方式"
Const str_QuickSearchArticle="查找文章"
Const str_SearchNow="立刻查找"
Const str_QuickSearchUser="快速查找用户"
Const str_UserGroup="用户组"
Const str_FunctionShortcut="快捷功能链接"
Const str_ColumnAdmin="栏目管理"
Const str_ArticleAdmin="文章管理"
Const str_ReLoadCache="更新数据"
Const str_ProductInformation="产品信息"
Const str_ProductCopyright="产品版权"
Const str_ProductSales="产品销售"
Const str_AboutMe="关于我们"
Const str_OperationNotice="操作说明"
Const str_UseGuide="使用向导"
Const str_IconEnabled="<strong>√</strong>"
Const str_IconDisabled="<font color=""red""><strong>×</strong></font>"

'admin_config.asp
Const str_Config_Help="这里的设置代表了系统中的全局设置，如果需要对某个版块或用户组进行单独设置，请到各自的管理中心操作"
Const str_Config_Base="基本设置"
Const str_Config_WebSiteName="站点名称"
Const str_Config_WebSiteURL="站点URL"
Const str_Config_WebSiteState="站点状态"
Const str_Config_WebSiteState_Open="开"
Const str_Config_WebSiteState_Close="关"
Const str_Config_ClosedMsg="关闭时的公告语"
Const str_Config_PageKeyWord="页面关键字"
Const str_Config_PageKeyWord_Help="每个关键字以 | 号分割"
Const str_Config_PageDescription="页面描述"
Const str_Config_PageDescription_Help="请不要使用英文的,号"
Const str_Config_SystemTimer="系统定时打开"
Const str_Config_SystemTimer_Value="定时开关的时间"
Const str_Config_SystemTimer_Value_Help="起止小时用符号 | 分开</font>"
Const str_Config_SystemMode="开启调试模式"
Const str_Config_SystemStyle="全站使用HTML模式"
Const str_Config_SystemIndexMode="首页使用HTML模式"
Const str_Config_Article="文章发布设置"
Const str_Config_PageSize="分页字数"
Const str_Config_DefaultPoster="默认发表人"
Const str_Config_ManagerEditor="后台文章编辑器"
Const str_Config_MemberEditor="会员文章编辑器"
Const str_Config_AutoRemote="远程图片自动上传"
Const str_Config_AutoRemote_Help="只限于eWeb Editor编辑器"
Const str_Config_Reg="注册设置"
Const str_Config_UserRegEnable="是否开启注册"
Const str_Config_EmailById="EMail只能注册一次"
Const str_Config_RegWaitAdmin="注册需要管理员验证"
Const str_Config_EMail="邮件设置"
Const str_Config_EMailAddress="EMail地址"
Const str_Config_SMTPServerAddress="SMTP服务器地址"
Const str_Config_SMTPLoginAccout="SMTP服务器登陆帐号"
Const str_Config_SMTPLoginPassWord="SMTP服务器登陆密码"
Const str_Config_Other="其他设置"
Const str_Config_GhostPostReview="允许游客发表评论"
Const str_Config_ReviewWaitAdmin="游客评论需要审核"
Const str_Config_GhostPostVote="允许游客参加投票"
Const str_Config_BadWord="系统屏蔽字符"
Const str_Config_BadWord_Help="格式:<br>屏蔽字符1==屏蔽后内容1;屏蔽字符2==屏蔽后内容2"
Const str_Config_ArticleFrom="文章来源"
Const str_Config_ArticleFrom_Help="格式:<br>名称1==网址1;名称2==网址2"
Const str_Config_SEO="搜索引擎优化(SEO)"

'admin_friend.asp
Const str_Friend_Help="①排序号为从大到小排列<br>②相关风格标签说明。<a href='http://help.nbarticle.com/tags_friend.html' target='_blank'>查看</a>"
Const str_Friend_LinkList="联盟列表"
Const str_Friend_AddLink="添加联盟"
Const str_Friend_Order="排序"
Const str_Friend_SiteName="站点名称"
Const str_Friend_SiteLogo="站点Logo"
Const str_Friend_SiteURL="站点URL"
Const str_Friend_SiteInfo="站点简介"
Const str_Friend_Location="所属版块"
Const str_Friend_State="状态"
Const str_Friend_Style="显示风格"
Const str_Friend_Input_Info="输入联盟资料"
Const str_Friend_Style_Img="图片"
Const str_Friend_Style_Txt="文本"
Const str_Friend_Index="首页"

'admin_placard.asp
Const str_Placard_Help="①相关风格标签说明。<a href='http://help.nbarticle.com/tags_siteplacard.html' target='_blank'>查看</a>"
Const str_Placard_PlacardList="公告列表"
Const str_Placard_AddPlacard="添加公告"
Const str_Placard_Title="公告标题"
Const str_Placard_Content="公告内容"
Const str_Placard_OverTime="过期时间"
Const str_Placard_AddTime="发布时间"
Const str_Placard_Input_Placard="输入公告内容"

'admin_vote.asp
Const str_Vote_Help="①每行一个投票项目，最多<strong>10</strong>个投票项目，超过则会自动作废，空行自动过滤<br />②相关风格标签说明。<a href='http://help.nbarticle.com/tags_sitevote.html' target='_blank'>查看</a>"
Const str_Vote_VoteList="投票列表"
Const str_Vote_AddVote="添加投票"
Const str_Vote_Title="投票主题"
Const str_Vote_VotedTotal="投票总数"
Const str_Vote_Type="投票类型"
Const str_Vote_State="状态"
Const str_Vote_Input_Vote="输入投票内容"
Const str_Vote_Content="投票内容"
Const str_Vote_Type_Radio="单选投票"
Const str_Vote_Type_Check="多选投票"
Const str_Vote_CanNotEdit="对不起，投票已开始，不能修改!"

'admin_insidelink.asp
Const str_InsideLink_Help="①站内连接：将文章正文内指定的关键词替换为连接到指定URL的超连接。关键词可限定作用范围，如全站有效或某栏目有效，并可在不同的栏目内指定同一关键词连接到不同地址。<br />②连接地址必须基于HTTP协议的网址格式,如:<br>&nbsp;http://www.youdomain.com<br>&nbsp;ftp://account:password@ipaddress:port"
Const str_InsideLink_InsideLinkList="站内连接列表"
Const str_InsideLink_AddInsideLink="添加站内连接"
Const str_InsideLink_LinkWord="连接字符串"
Const str_InsideLink_LinkURL="连接URL"
Const str_InsideLink_Location="所属栏目"
Const str_InsideLink_Input_Info="输入站内连接资料"
Const str_InsideLink_All="全站"

'admin_column.asp
Const str_Column_Help="①新添加的栏目，必须先到安全管理-=>管理员管理中给管理员赋于相应的管理权限后，才能在添加文章时显示在所属栏目中出现。<br />②相关风格标签说明。<a href=""http://help.nbarticle.com/template_tags.html"" target=""_blank"">查看</a>"
Const str_Column_ColumnList="栏目列表"
Const str_Column_AddColumn="添加栏目"
Const str_Column_Title="栏目名称"
Const str_Column_ArticleTotal="已审核文章"
Const str_Column_ManagerArticleTotal="未审核文章"
Const str_Column_Order="排序"
Const str_Column_ResetBoard="复位版面"
Const str_Column_Confirm_ResetBoard="确认复位所有栏目吗?"
Const str_Column_MoveArticle="移动文章"
Const str_Column_DelArticle="清空文章"
Const str_Column_Input_Column="输入栏目内容"
Const str_Column_Type="栏目类型"
Const str_Column_Type_Normal="正常"
Const str_Column_Type_Diss="专题"
Const str_Column_Attrib="所属分类"
Const str_Column_Info="栏目简介"
Const str_Column_OutURL="外部连接"
Const str_Column_ViewPower="浏览权限"
Const str_Column_IsHide="VIP区"
Const str_Column_IsReview="允许发表评论"
Const str_Column_IsPost="允许会员投稿"
Const str_Column_IsTop="在导航菜单显示"
Const str_Column_Style="列表风格"
Const str_Column_Style_Txt="纯标题文本"
Const str_Column_Style_Info="标题+简介"
Const str_Column_Style_Img="纯标题图片"
Const str_Column_Style_Txt_Info_Img="标题+简介+图片"
Const str_Column_ListTemplate="栏目页风格"
Const str_Column_ArticleTemplate="内容页风格"
Const str_Column_PageSize="列表页记录数"
Const str_Column_EofOrBof="对不起,已经是尽头不能再移动了."
Const str_Column_ColumnIsNotEmpty="对不起,栏目内还有文章,请清空或转移后再进行删除操作."
Const str_Column_ColumnHaveUnder="对不起,栏目尚有下属版块."
Const str_Column_Root="根目录"

'admin_content.asp
Const str_Content_Help="①删除单个文章默认为直接删除，而不进入回收站<br>②当使用批量管理功能之后，建议更新一次系统及栏目数据<br>③相关风格标签说明。<a href='http://help.nbarticle.com/tags_getarticlelist.html' target='_blank'>查看</a>"
Const str_Content_Add_Help="①只有在“标题图片”处填写图片的url后，才为图片新闻，建议图片大小为120*90<br>②当填写外部连接之后，即代表文章将连接到外部网址，请确保填写的地址正确。连接地址必须基于HTTP协议的网址格式,如:<br>&nbsp;http://www.youdomain.com<br>&nbsp;ftp://account:password@ipaddress:port<br>③可设定多个文章搜索关键字，关键字之间请以英文,号分隔，例如：xp,sp2<br>④选择手动分页时，分页标记是[NextPage]，<font color=red>区分大小写</font>。如需要对每页指定标题，则为[NextPage=标题内容]"
Const str_Content_ArticleList="所有文章"
Const str_Content_AddArticle="添加文章"
Const str_Content_TColor="标题颜色"
Const str_Content_Title="文章标题"
Const str_Content_Keyword="文章关键字"
Const str_Content_Author="文章作者"
Const str_Content_Img="标题图片"
Const str_Content_ReviewImg="[预览图片]"
Const str_Content_Top="推荐文章"
Const str_Content_Column="所属栏目"
Const str_Content_OutURL="外部连接"
Const str_Content_CutArticle="内容分页"
Const str_Content_CutArticle_Not="不分页"
Const str_Content_CutArticle_Auto="自动分页"
Const str_Content_CutArticle_Manual="手动分页"
Const str_Content_Content="文章正文"
Const str_Content_Page="每页"
Const str_Content_Word="字"
Const str_Content_Source="文章来源"
Const str_Content_Summary="文章摘要"
Const str_Content_SummaryFromText="从正文截取"
Const str_Content_Date="日期"
Const str_Content_ViewNum="浏览数"
Const str_Content_State="状态"
Const str_Content_PassNow="立即发布"
Const str_Content_SaveAs="另存为.."
Const str_Content_Review="评论"
Const str_Content_Option="操作"
Const str_Content_Resume="恢复"
Const str_Content_Make="生成HTML"
Const str_Content_BatchMove="批量移动到-=>"
Const str_Content_Color_Red="红色"
Const str_Content_Color_Green="绿色"
Const str_Content_Color_Blue="蓝色"
Const str_Content_Open_AdvancedFunction="打开增强选项"
Const str_Content_Close_AdvancedFunction="关闭增强选项"
Const str_Content_Preview="预览文章"
Const str_Content_ArticleTemplate="文章模版"

'admin_columnbath.asp
Const str_Bath_Move_Help="①这里只是移动文章，而不是拷贝或者删除！ <br />②您可以将一个下属栏目的文章移动到上级栏目，也可以将上级栏目的文章移动到下级栏目<br />③当使用该功能之后，建议更新一次系统及栏目数据"
Const str_Bath_Del_Help="①如果选择了不入回收站，本操作将批量删除栏目文章，<font color=red>则所有操作不可恢复。</font>如需还原文章，请到回收站。<br />②如果您确定这样做，请仔细检查您输入的信息。<br />③当使用该功能之后，建议更新一次系统及栏目数据"
Const str_Bath_BathMove="批量移动"
Const str_Bath_From="从"
Const str_Bath_To="到"
Const str_Bath_Condition="筛选条件"
Const str_Bath_All="所有文章"
Const str_Bath_ByDate="按日期筛选"
Const str_Bath_ByKeyword="按关键字筛选"
Const str_Bath_In="包含在"
Const str_Bath_BathDel="批量删除"
Const str_Bath_DestColumn="删除栏目"
Const str_Bath_Option="选项"
Const str_Bath_RecycleBin="放入回收站"
Const str_Bath_NoRecycleBin="直接删除"
Const str_Bath_MoveMsg="共转移了 $1 篇文章，建议更新栏目统计数据！"
Const str_Bath_DelMsg="共删除了 $1 篇文章，建议更新栏目统计数据！"

'admin_review.asp
Const str_Review_User="发表用户(IP)"
Const str_Review_UnderArticle="所属文章"
Const str_Review_Content="评论内容"
Const str_Review_AddTime="发表时间"
Const str_Review_State="状态"
Const str_Review_ViewArticle="查看"
Const str_Review_Help=""

'admin_arttemplate.asp
Const str_ArticleTemplate_Help="文章模版用于在添加文章时，可快速调用固定格式的内容以填充到正文。"
Const str_ArticleTemplate_TemplateList="模版列表"
Const str_ArticleTemplate_AddTemplate="添加模版"
Const str_ArticleTemplate_TemplateName="模版名称"
Const str_ArticleTemplate_Input_Template="输入模版内容"

'admin_usergroup.asp
Const str_Group_Help="①在这里您可以设置各个用户组在系统中的默认权限，系统默认用户组不能<font color=red>删除</font><BR>②如将该组用户转移到其他用户组，请到用户列表中进行操作<BR>③可以删除和编辑新添加的用户组<BR>④<strong>如果删除用户组，则该用户组所包含的用户将自动转到临时用户组</strong>"
Const str_Group_List="会员组列表"
Const str_Group_Add="添加会员组"
Const str_Group_Name="组名"
Const str_Group_UserTotal="用户数量"
Const str_Group_ShowList="列出用户"
Const str_Group_Login_Option="登陆设置"
Const str_Group_IsLogin="允许登陆"
Const str_Group_Power="组权限"
Const str_Group_ViewHide="浏览VIP区"
Const str_Group_Timer="定时登陆"
Const str_Group_Timer_Option="定时登陆时段"
Const str_Group_Timer_Option_Help="起止小时用符号 | 分开"
Const str_Group_Power_Option="权限设置"
Const str_Group_PostReviewForRegLater="注册n分钟后允许评论"
Const str_Group_PostReviewForRegLater_Help="0为不限制"
Const str_Group_PostVotedForRegLater="注册n分钟后允许投票"
Const str_Group_PostVotedForRegLater_Help="0为不限制"
Const str_Group_IsReview="允许发表评论"
Const str_Group_ReviewForManager="评论需要审核"
Const str_Group_IsPostArticle="允许发表投稿"
Const str_Group_PostForManager="投稿需要审核"
Const str_Group_DayMaxPost="日投稿上限"
Const str_Group_FavMax="收藏夹上限"
Const str_Group_FavMax_Help="0为不允许收藏"

'admin_member.asp
Const str_Member_Help=""
Const str_Member_Account="帐号"
Const str_Member_GroupName="会员组"
Const str_Member_RegDate="注册日期"
Const str_Member_State="状态"
Const str_Member_MoveTo="移动到"
Const str_Member_MemberInfo="会员资料"
Const str_Member_Sex="性别"
Const str_Member_Sex_Man="男"
Const str_Member_Sex_Woman="女"
Const str_Member_NickName="真实姓名"
Const str_Member_LoginTotal="登陆次数"
Const str_Member_BirtDay="生日"
Const str_Member_IsLogin="允许登陆"
Const str_Member_ArticleTotal="发表文章总数"
Const str_Member_LoginPassword = "登陆密码"
Const str_Member_PasswordEditInfo = "如不修改，请留空"

'admin_mailout.asp
Const str_MailOut_Help="①如选择导出到数据库，请确认NB_MailList.mdb已于指定目录下。（默认在/manager/Databackup目录中）<br>②使用导出到文本的功能需要服务器端必须支持FSO，导出地址为/manager/mail.txt，关于FSO请查询微软的网站<br>③导出邮件列表可能非常耗费服务器资源，请尽量在本地或在网络不繁忙的时候执行"
Const str_MailOut_OutToDataBase="数据库"
Const str_MailOut_OutToTxt="文本"
Const str_MailOut_ExportList="导出邮件地址列表"
Const str_MailOut_ExportType="导出方式"
Const str_MailOut_Now="开始导出"
Const str_mailOut_SuccessMsg="导出邮件地址列表成功，共 $1 条记录，<a href=""$2"">点击这里下载</a>"

'admin_master.asp
Const str_Master_Help="这里可以设置后台管理员在后台菜单的访问权限及文章栏目的使用权限"
Const str_Master_Account="登陆帐号"
Const str_Master_Password="登陆密码"
Const str_Master_LastLoginTime="上次登陆时间"
Const str_Master_LastLoginIp="上次登陆IP"
Const str_Master_State="状态"
Const str_Master_EditInfo="设定帐户信息"
Const str_Master_AccountList="管理员列表"
Const str_Master_AddAccount="添加管理员"
Const str_Master_MasterTotal="位管理员"
Const str_Master_AddMaster_AccountInfo="50个字符长度，建议只使用英文"
Const str_Master_AddMaster_PasswordInfo="<font color=red>如只修改用户登陆名而不修改密码请留空</font>"
Const str_Master_ColumnPowerOption="设定栏目权限"
Const str_Master_MenuPowerOption="设定后台权限"
Const str_Master_Column_Add="发布"
Const str_Master_Column_Manager="审核"
Const str_Master_Column_Edit="编辑"
Const str_Master_Column_Del="删除"

'admin_ip.asp
Const str_IP_Help="①如屏蔽一段ip，只需输入该段的开始ip及结束ip即可。如屏蔽192.168.1.1至192.168.1.100的ip段，只需填入192.168.1.1和192.168.1.100即可<br>②如只屏蔽一个ip，则开始ip及结束ip都为该ip地址。"
Const str_IP_IPHead="IP头"
Const str_IP_IPFoot="IP尾"
Const str_IP_OverTime="过期日期"
Const str_IP_AddIp="添加IP规则"
Const str_IP_IpList="IP规则列表"
Const str_IP_InputIp="输入IP规则"

'admin_makeindex.asp
Const str_MakeIndex_Help = "在这里将会把首页保存为根目录下的default.htm文件。每次有新增文章后，应当重新生成一次。如无需将首页静态化，只需删除该文件即可。"
Const str_MakeIndex_Info = "生成HTML首页"
Const str_MakeIndex_StartNow = "马上生成"
Const str_Operation_Complate="个任务已经完成"

'admin_makelist.asp
Const str_MakeList_Help="①本系统提供八种列表排序方式。根据文章数量及服务器性能不同，生成的时间也不同。建议每天批量更新文章两次，每次按时间降序排列生成一次，其他排序方式可以每两天批量生成一次。<br />②每页显示的数量越多，服务器执行速度越快；系统第一次使用必须全部生成一次。"
Const str_MakeList_Title="批量生成静态列表页"
Const str_MakeList_Option_1="第一种：按 <font color=800000>发布时间</font> <font color=800000>降序</font> 排列"
Const str_MakeList_Option_2="第二种：按 <font color=800000>发布时间</font> <font color=800000>升序</font> 排列"
Const str_MakeList_Option_3="第三种：按 <font color=800000>文章标题</font> <font color=800000>降序</font> 排列"
Const str_MakeList_Option_4="第四种：按 <font color=800000>文章标题</font> <font color=800000>升序</font> 排列"
Const str_MakeList_Option_5="第五种：按 <font color=800000>浏览人数</font> <font color=800000>降序</font> 排列"
Const str_MakeList_Option_6="第六种：按 <font color=800000>浏览人数</font> <font color=800000>升序</font> 排列"
Const str_MakeList_Option_7="第七种：按 <font color=800000>评论人数</font> <font color=800000>降序</font> 排列"
Const str_MakeList_Option_8="第八种：按 <font color=800000>评论人数</font> <font color=800000>升序</font> 排列"
Const str_MakeList_Column="栏目"
Const str_MakeList_Task="任务"
Const str_MakeList_Page="页面"
Const str_MakeList_Now="现第"
Const str_MakeList_AllComplate="所有任务已完成"

'admin_makeview.asp
Const str_MakeView_Help="①文章ID，可以填写您想从哪一个ID号开始进行生成静态网页。具体ID号可以在<a href=""javascript:vod();"" onclick=""javascript:ajaxChangeRightContent(\'./templates/admin_content_list.htm\');"">文章管理</a>中查看。<br />②无论是选择""按ID生成""还是按""按日期生成""，相互之间的间隔最好不要选择过大。"
Const str_MakeView_Title="批量生成静态内容页"
Const str_MakeView_MakeById="按ID生成"
Const str_MakeView_MakeByDate="按日期生成"
Const str_MakeView_MakeByColumn="按栏目生成"
Const str_MakeView_StartId="开始Id:"
Const str_MakeView_EndId="结束Id:"
Const str_MakeView_StartDate="开始时间:"
Const str_MakeView_EndDate="结束时间:"
Const str_MakeView_NextPage="下一页"
Const str_MakeView_PreviousPage="上一页"

'reload.asp
Const str_ReLoad_Help="<li><font color='red'>下列这些操作可能会非常消耗服务器资源，请慎重使用。</font><li>本页面的操作将需要使用服务器Adodb.Stream组件，请确认是否支持.<li>在生成之前，建议先更新系统数据。"
Const str_ReLoad_UpDateSystem="更新系统数据"
Const str_ReLoad_MakeSitemaps="生成Sitemaps文件"
Const str_ReLoad_MakeSitemaps_Desc="sitemaps文件生成于网站根目录下的sitemap-index.xml，请在 <a href='https://www.google.com/webmasters/sitemaps/login?hl=zh_CN' target=_blank>Google Sitemap</a> 上提交，将可能非常有助于Google对您网站页面的收录。"
Const str_ReLoad_MakeBaiduNewsop="生成百度新闻开放协议文件"
Const str_ReLoad_MakeBaiduNewsop_Desc="协议文件生成于网站根目录下的baidu-newsop.xml，请在 <a href='http://news.baidu.com/newsop.html#ks5' target=_blank>这里</a> 提交，将可能有助于百度对您网站页面的收录。"

'admin_template.asp
Const str_Theme_Info="①在这里，您可以添加和编辑风格，操作时请按照相关页面提示完整填写表单信息。<BR>②默认风格不能删除"
Const str_Theme_Name="风格名称"
Const str_Theme_Default="默认风格"
Const str_Theme_Add="添加风格"
Const str_Theme_Edit="编辑风格"

Const str_Theme_ModuleInfo="①在这里，您可以添加和编辑模块，操作时请按照相关页面提示完整填写表单信息。<BR>②<a href=""http://help.nbarticle.com/template_tags.html"" target=_blank><font color=red><u>风格标签说明</u></font></a>"
Const str_Theme_ModuleName="模块名称"
Const str_Theme_ModuleDesc="模块简介"
Const str_Theme_ModuleType="模块类型"
Const str_Theme_ModuleCode="模块代码"
Const str_Theme_ModuleHome="首页模块"
Const str_Theme_ModuleCss="Css模块"
Const str_Theme_ModuleHead="头部模块"
Const str_Theme_ModuleFoot="底部模块"
Const str_Theme_ModulePage="模块"
Const str_Theme_ModuleContent="内页模块"
Const str_Theme_ModuleAdd="添加模块"
Const str_Theme_ModuleEdit="编辑模块"

'admin_loadskin.asp
Const str_Skin_Help="①导出功能可以将数据库内的风格复制到包含有特定结构的数据库内。你可以使用manager/databackup/NB_Template.mdb这个文件。<br />②导入功能可以将具有特定结构的数据库内的风格复制到网站数据库中。建议使用系统自带的manager/databackup/NB_Template.mdb作为相互间风格传递的载体。"
Const str_Skin_Database="数据库"
Const str_Skin_Choice="选择"
Const str_Skin_InputTemplate="导入风格"
Const str_Skin_OutputTemplate="导出风格"
Const str_Skin_Input="导入"
Const str_Skin_Output="导出"
Const str_Skin_CopyComplete="风格复制成功！"

'admin_makejs.asp
Const str_Js_Help="①在这里，您可以新建和修改Js生成文件，操作时请按照相关页面提示完整填写表单信息。<br>②图片显示参数格式：宽&高，如160&90<br />③<a href=""http://forum.nbarticle.com/forum_posts.asp?TID=138"" target=""_blank"">论坛使用介绍</a><br />④系统Js文件：指js目录下的 friend.js(友情连接)、menu.js(栏目导航)、searchbar.js(搜索栏)"
Const str_Js_List="Js文件列表"
Const str_Js_Add="添加Js文件"
Const str_Js_Title="文件说明"
Const str_Js_Detail="文件介绍"
Const str_Js_UpDate="更新"
Const str_Js_Preview="预览"
Const str_Js_UpDateAll="更新全部Js文件"
Const str_Js_UpDateSystemJs="更新系统Js文件"
Const str_Js_BaseSetting="基本设置"
Const str_Js_TransferSetting="调用设置"
Const str_Js_FilePath="文件路径"
Const str_Js_TransferColumn="调用栏目"
Const str_Js_IncludeChildColumn="包含子栏目"
Const str_Js_ListStyle="列表样式"
Const str_Js_ListStyle_Txt="纯文本"
Const str_Js_ListStyle_Detail="图文(上标题+左下图+右下摘要)"
Const str_Js_ListStyle_Mix="图文(左图右标题)"
Const str_Js_ListStyle_Img="纯图片"
Const str_Js_TransferTotal="调用文章数"
Const str_Js_TransferType="调用文章类型"
Const str_Js_Transfer_AllArticle="所有文章"
Const str_Js_Transfer_CommendArticle="推荐文章"
Const str_Js_Transfer_HotArticle="热门文章"
Const str_Js_Transfer_ImgArticle="图片文章"
Const str_Js_TitleLen="标题长度"
Const str_Js_ContentLen="摘要长度"
Const str_Js_ShowFields="显示项目"
Const str_Js_ShowFields_ColumnName="栏目名称"
Const str_Js_ShowFields_NewTag="新文章标记"
Const str_Js_ShowFields_AddTime="发布日期"
Const str_Js_ShowFields_TypeTag="类型标记"
Const str_Js_ShowFields_ReviewLink="评论连接"
Const str_Js_OpenWindowType="窗口打开方式"
Const str_Js_OpenWindowType_Parent="原窗口"
Const str_Js_OpenWindowType_New="新窗口"
Const str_Js_RowTotal="每行列数"
Const str_Js_ImgSize="图片大小"

'admin_data.asp
Const str_Data_Backup_Help="<li>您可以用这个功能来备份您的系统数据，以保证您的数据安全！<br><li><font color=800000><strong>注意：所有路径都是与当前程序运行目录的相对路径</strong></font><li><font color=red><strong>需要FSO支持，FSO相关帮助请看微软网站</strong></font>"
Const str_Data_Execute_Help="请填写需要执行的sql语句。<strong>请注意使用正确的语法。</strong><br><font color=800000><strong>注意：语句执行后将不能恢复，请慎重操作！</strong></font>"
Const str_Data_Backup_Title="备份系统数据库"
Const str_Data_Backup_BackupFolder="备份目录"
Const str_Data_Backup_BackupFileName="备份文件名称"
Const str_Data_Execute_Title="数据库高级管理"
Const str_Data_Execute_ExecuteText="执行语句"
Const str_Data_SpaceView_Title="系统空间占用情况"
Const str_Data_SpaceView_BackupFolder="备份数据占用空间"
Const str_Data_SpaceView_DatabaseFolder="数据库文件占用空间"
Const str_Data_SpaceView_UploadFolder="上传图片占用空间"
Const str_Data_SpaceView_ManagerFolder="后台数据占用空间"
Const str_Data_SpaceView_HTMLListFolder="HTML列表页文件占用空间"
Const str_Data_SpaceView_HTMLViewFolder="HTML内容页文件占用空间"
Const str_Data_SpaceView_All="系统占用空间总计"

'admin_upfile.asp
Const str_UpFile_Help="①、本功能必须服务器支持FSO权限方能使用，FSO使用帮助请浏览微软网站。如果您服务器不支持FSO请手动管理。<BR>②、上传目录强制定义为UploadFile"
Const str_UpFile_Type="类型"
Const str_UpFile_FileName="文件地址"
Const str_UpFile_Size="大小"
Const str_UpFile_LastTime="最后访问"
Const str_UpFile_UploadTime="上传日期"
Const str_UpFile_CurrentPath="当前目录"

'admin_adsense.asp
Const str_AdSense_Help="①相关风格标签说明。<a href='http://help.nbarticle.com/tags_adsense.html' target='_blank'>查看</a>"
Const str_AdSense_Title="广告标题"
Const str_AdSense_Content="广告内容"
Const str_AdSense_Input_Template="输入广告内容"
Const str_AdSense_List="广告列表"
Const str_AdSense_Add="添加广告"

'admin_Interface.asp
Const str_Interface_Help="①<a href=""http://help.nbarticle.com/superpassport.html"" target='_blank'>开发说明</a><BR>②<a href='http://forum.nbarticle.com/forum_posts.asp?TID=3885' target='_blank'>使用说明</a>"
Const str_Interface_List="接口列表"
Const str_Interface_Add="添加接口"
Const str_Interface_Title="接口说明"
Const str_Interface_RemoteURL="远程接口地址"
Const str_Interface_StructFile="接口文件"
Const str_Interface_Type="接口类型"
Const str_Interface_Type_UserRegister="会员注册"
Const str_Interface_Type_UserChanngePassword="会员修改密码"
Const str_Interface_Type_UserChanngeInfo="会员修改资料"
Const str_Interface_Type_UserPostArticle="会员发表文章"
Const str_Interface_Type_ManagerPostArticle="管理员发表文章"
Const str_Interface_SKey="接口密匙"
Const str_Interface_Input="输入接口信息"
%>