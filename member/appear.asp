<!--#Include File="../include/inc.asp"-->
<!--#Include File="cls_db.asp"-->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：Member/Appear.asp
'= 摘    要：会员-会员发布文章文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2008-03-22
'====================================================================

Dim EA_Mem_DBO
Set EA_Mem_DBO = New cls_Member_DBOperation

If Not EA_Pub.IsMember Then Call EA_Pub.ShowErrMsg(41, 2)
If EA_Pub.Mem_GroupSetting(10)="0" Then Call EA_Pub.ShowErrMsg(42, 2)
If CLng(EA_Mem_DBO.Get_MemberDayPostTotal(EA_Pub.Mem_Info(0))(0, 0)) >= CLng(EA_Pub.Mem_GroupSetting(13)) Then Call EA_Pub.ShowErrMsg(43, 2)

If LCase(Request.QueryString ("action"))="save" Then Call Save_MemberAppear()

Dim ColumnOption, ArticleInfo
Dim Title, Text, KeyWord,ColumnId,ImgPath,Source,SourceUrl,Summary
Dim i,Level,TempArray,TStr,Temp
Dim EA_Editor
Dim PostId

PostId=EA_Pub.SafeRequest(3,"postid",0,0,0)

ArticleInfo=EA_DBO.Get_Article_Info(PostId,0)
If IsArray(ArticleInfo) Then 
	If CStr(ArticleInfo(7,0))<>EA_Pub.Mem_Info(0) Then Call EA_Pub.ShowErrMsg(23,1)

	Title=ArticleInfo(3,0)
	Text=ArticleInfo(5,0)
	KeyWord=ArticleInfo(12,0)
	ColumnId=ArticleInfo(0,0)
	ImgPath=ArticleInfo(18,0)
	Source=ArticleInfo(15,0)
	SourceUrl=ArticleInfo(16,0)
	Summary=ArticleInfo(4,0)
End If
	
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="zh-cn">
<head>
<title>我要投稿 - <%=EA_Pub.SysInfo(0)%></title>
<meta name="generator" content="NB文章系统(NBArticle) - <%=SysVersion%>" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="Content-Language" content="zh-cn" />
<link href="style.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="../js/public.js"></script>
<script type="text/javascript" src="../plugins/fck_editor/fckeditor.js"></script>
</head>
<body id="center">

<form name="a1" method="post" action="?action=save&postid=<%=PostId%>" onsubmit="return checkData()">

<div style="background: #DBF2FF; border: #A9D5F4 1px solid; line-height: 25px;">&nbsp;<strong>会员投稿</strong></div>
<div class="left" style="border: #DBE1E9 1px solid;padding: 5px;">
	<table>
		<tr>
			<td><select name="column"> 
                <option value="0">请选择栏目...</option> 
                <%
					ColumnOption=EA_Mem_DBO.Get_MemberAppearColumnList()
					
					If IsArray(ColumnOption) Then 
						For i=0 To UBound(ColumnOption,2)
							Level=(Len(ColumnOption(2,i))/4-1)*3
							Response.Write "<option value="""&ColumnOption(0,i)&"|||"&ColumnOption(1,i)&"|||"&ColumnOption(2,i)&""""
							If ColumnOption(0,i)=ColumnId Then Response.Write " selected"
							Response.Write ">"
							If Len(ColumnOption(2,i))>4 Then Response.Write "├"
							Response.Write String(Level,"-")
							Response.Write ColumnOption(1,i)&"</option>"
						Next
					End If
				%> 
              </select>&nbsp;<input type="text" name="title" size="50" value="<%=Title%>" /></td>
		</tr>
		<tr>
			<td id="myFCKeditor"></td>
		</tr>
	</table>
</div>

<div style="border: #DBE1E9 1px solid;padding:5px">
	<table>
		<tr>
			<td valign="top">关键字：</td><td><textarea name="keyword" cols="28" rows="4" title="以,号分隔"></textarea></td>
		</tr>
		<tr>
			<td valign="top">摘要：</td><td><textarea name="summary" id="summary" cols="28" rows="8" wrap="VIRTUAL"></textarea></td>
		</tr>
		<tr>
			<td>副标题：</td><td><input type="text" name="subtitle" size="33"></td>
		</tr>
		<tr>
			<td>副标题连接：</td><td><input type="text" name="suburl" size="33"></td>
		</tr>
		<tr>
			<td>标题图片：</td><td><input type="text" name="img" value="<%=ImgPath%>">&nbsp;<a href="javascript: vod();" onclick="review_img();">[预览]</a></td>
		</tr>
		<tr>
			<td valign="top">文章来源：</td><td><select name="choosesource" onchange="inputsource(this.options[this.selectedIndex].value)">
			<option value="==">快捷选择</option>
			<%
			Temp=EA_DBO.Get_System_Info()
			If IsArray(Temp) Then
				If Not IsNull(Temp(6,0)) Then 
					TempArray=Split(Temp(6,0),";")
		
					For i=0 To UBound(TempArray)
						TStr=""
						TStr=Split(TempArray(i),"==")
						
						Response.Write "<option value='"&TempArray(i)&"'>"&TStr(0)&"</option>"
					Next
				End If
			End If
			%>
      </select><br /><input name="source" type="text" value="<%=Source%>"><br /><input name="sourceurl" type="text" value="<%=SourceUrl%>" size=30></td>
		</tr>
	</table>

	<input type="hidden" name="author" value="<%=EA_Pub.Mem_Info(0)%>"> 
	<input type="submit" name="Submit1" value=" 提交我的文章 ">&nbsp;<input type="reset" name="Submit2" value="重置">
</div>

</form>
<script type="text/javascript">
function checkData() {
	var f1 = document.a1;
	var wm = "\n";
	var noerror = 1;
	var isAutoRemote=<%=EA_Pub.SysInfo(22)%>;
	var iEditorType=<%=EA_Pub.SysInfo(24)%>;

	var t1 = a1.Column;
	if (t1.value == "" || t1.value == "0") {
		wm += "请选择添加的栏目\r\n";
		noerror = 0;
	}

	var t1 = a1.title;
	if (t1.value == "" || t1.value == " ") {
		wm += "请填写文章的标题\r\n";
		noerror = 0;
	}

	if (iEditorType==0 || iEditorType==2)
	{
		if (iEditorType==0)
		{
			var sHTML = content1.getHTML();
		}
		else{
			f1.content.value = oEdit1.getHTMLBody();
			var sHTML = f1.content.value;
		}
		if (sHTML == "" || sHTML == " ") {
			wm += "请填写文章的正文\r\n";
			noerror = 0;
		}
	}

	if (noerror == 0) {
		alert(wm);
	}
	else{
		document.a1.Submit1.value ="正在提交，请稍候...";
		document.a1.Submit1.disabled=true;

		if (isAutoRemote==1 && iEditorType==0){
			content1.remoteUpload("doSubmit()");
		}
		else{
			doSubmit();
		}
	}
	return false;
}

// 表单提交（当远程上传完成后，触发此函数）
function doSubmit(){
	document.a1.submit();
}

function inputsource(Str){
	var tmp
	
	tmp=Str.split("==");
	
	document.all.source.value=tmp[0];
	document.all.sourceurl.value=tmp[1];
	
}

function change(obj,i) {
he=parseInt(obj.style.height);
if (he>=80&&he<=400)
   obj.style.height=he+i+'px';
else 
   obj.style.height='80px';
}

function doChange(objText, objDrop){
	if (!objDrop) return;
	var str = objText.value;
	var arr = str.split("|");
	var nIndex = objDrop.selectedIndex;
	objDrop.length=1;
	for (var i=0; i<arr.length; i++){
		objDrop.options[objDrop.length] = new Option('上传图片'+[i+1], arr[i]);
	}
	objDrop.selectedIndex = nIndex;
}
  
function review_img(){
	if (document.all.img.value!=''){
		window.open(''+document.all.img.value+'','','');
	}
}

var div = $("myFCKeditor");
var fck = new FCKeditor("Content", 650, 480);
fck.BasePath	= "../plugins/fck_editor/";
fck.Value		= "<%=Text%>";
div.innerHTML	= fck.CreateHtml();
</script>
</body>
</html>
<%

Sub Save_MemberAppear()
	Call EA_Pub.Chk_Post
	
	Dim PostId,TempStr
	Dim Key
	Dim Feedback
	ReDim ArticleInfo(18)
	
	If Request.Form("column")="" Or Request.Form("column")="0" Then 
		Call EA_Pub.ShowErrMsg(2, 2)
	Else
		TempStr=Split(Request.Form ("Column"),"|||")
	End If
	
	FoundErr=False
	
	PostId=EA_Pub.SafeRequest(3,"postid",0,0,0)
	ArticleInfo(0)=EA_Pub.SafeRequest(2,"title",1,"",1)
	ArticleInfo(1)=EA_Pub.BadWords_Filter(EA_Pub.SafeRequest(2,"content",1,"",-1))
	ArticleInfo(2)=EA_Pub.SafeRequest(2,"keyword",1,"",1)
	ArticleInfo(3)=EA_Pub.SafeRequest(0,Trim(TempStr(0)),0,0,0)
	ArticleInfo(4)=EA_Pub.SafeRequest(0,Trim(TempStr(1)),1,"",1)
	ArticleInfo(5)=EA_Pub.SafeRequest(0,Trim(TempStr(2)),1,"",0)
	ArticleInfo(6)=EA_Pub.SafeRequest(2,"img",1,"",1)
	ArticleInfo(8)=EA_Pub.SafeRequest(2,"source",1,"",1)
	ArticleInfo(9)=EA_Pub.SafeRequest(2,"sourceurl",1,"",1)
	ArticleInfo(10)=EA_Pub.BadWords_Filter(EA_Pub.SafeRequest(2,"summary",1,"",2))
	ArticleInfo(11)=Lenb(Text)

	If Len(ArticleInfo(0))>150 Or Len(ArticleInfo(0))=0 Then 
		ErrMsg = 44
		FoundErr=True
	End If
	If Len(ArticleInfo(2))>20 Then 
		ErrMsg = 45
		FoundErr=True
	End If
	If Len(ArticleInfo(10))>250 Then 
		ErrMsg = 46
		FoundErr=True
	End If
	If Not EA_DBO.Get_Column_Info(ArticleInfo(3))(15,0) Then 
		ErrMsg = 47
		FoundErr=True
	End If
	
	If FoundErr Then Call EA_Pub.ShowErrMsg(ErrMsg, 2)

	If ArticleInfo(6)="" Then ArticleInfo(6)=EA_Pub.SafeRequest(2,"d_picture",1,"",1)
	If ArticleInfo(6)="" Then 
		ArticleInfo(7)=0
	Else
		ArticleInfo(7)=1
	End If
	If EA_Pub.Mem_GroupSetting(11)="1" Then 
		ArticleInfo(12)=0
	Else
		ArticleInfo(12)=1
	End If
	
	ArticleInfo(13)=EA_Pub.Mem_Info(0)
	ArticleInfo(14)=EA_Pub.Mem_Info(1)
	ArticleInfo(15)=EA_Pub.SysInfo(5)&"|"&EA_Pub.SysInfo(6)
	ArticleInfo(16)=Now()
	ArticleInfo(17)=EA_DBO.Get_Column_Info(ArticleInfo(3))(11,0)
	
	ArticleInfo(18)=Year(CDate(ArticleInfo(16)))
	ArticleInfo(18)=ArticleInfo(18)&Right("00"&Month(CDate(ArticleInfo(16))),2)
	ArticleInfo(18)=ArticleInfo(18)&Right("00"&Day(CDate(ArticleInfo(16))),2)
	ArticleInfo(18)=ArticleInfo(18)&Right("00"&Hour(CDate(ArticleInfo(16))),2)
	ArticleInfo(18)=ArticleInfo(18)&Right("00"&Minute(CDate(ArticleInfo(16))),2)
	ArticleInfo(18)=ArticleInfo(18)&Right("00"&Second(CDate(ArticleInfo(16))),2)
	
	Randomize Timer
	key="000000"&CStr(Int((999999-1+100000)*Rnd+1))
	ArticleInfo(18)=ArticleInfo(18)&Right(Key,6)
	
	Feedback=EA_Mem_DBO.Set_Article_Insert(PostId,ArticleInfo)
	
	Select Case Feedback
	Case -1
		Call EA_Pub.ShowErrMsg(48, 2)
	Case 1
		Application.Lock 
		Application(sCacheName&"IsFlush")=1
		Application.UnLock 

		Call EA_Pub.ShowErrMsg(49, 2)
	Case 0
		Application.Lock 
		Application(sCacheName&"IsFlush")=1
		Application.UnLock 

		Call EA_Pub.ShowErrMsg(50, 2)
	End Select
End Sub
%> 
