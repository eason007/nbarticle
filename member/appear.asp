<!--#Include File="../conn.asp" -->
<!--#Include File="../include/inc.asp"-->
<!--#Include File="../include/cls_editor.asp" -->
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
'= 最后日期：2006-09-12
'====================================================================

If Not EA_Pub.IsMember Then Call EA_Pub.ShowErrMsg(10,1)

If EA_Pub.Mem_GroupSetting(10)="0" Then Call EA_Pub.ShowErrMsg(14,1)

If CLng(EA_DBO.Get_MemberDayPostTotal(EA_Pub.Mem_Info(0))(0,0))>=CLng(EA_Pub.Mem_GroupSetting(13)) Then Call EA_Pub.ShowErrMsg(41,1)

If LCase(Request.QueryString ("action"))="save" Then Call Save_MemberAppear()

Dim ColumnList,ArticleInfo
Dim Title,Text,KeyWord,ColumnId,ImgPath,Source,SourceUrl,Summary
Dim i,Level,TempArray,TStr,Temp
Dim EA_Editor
Dim PostId

PostId=EA_Pub.SafeRequest(3,"postid",0,0,0)

Set EA_Editor=New cls_Editor

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
	
EA_Editor.EditorType=EA_Pub.SysInfo(24)
EA_Editor.Value=Text
%>
<html>
<head>
<title>我要投稿</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta http-equiv="Content-Language" content="zh-CN">
<link href="style.css" rel="stylesheet" type="text/css" />
<SCRIPT language=JavaScript>
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
	else{
		//alert(document.all.d_picture.value);
		if (document.all.d_picture.selectedIndex>0){
			window.open(''+document.all.d_picture.value+'','','');
		}
	}
}
</SCRIPT>
</head>
<body bgcolor="#FFFFFF" text="#000000"> 
<table width="762" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC"> 
  <tr> 
    <td bgcolor="#FFFFFF"><table width="100%" border="0" cellpadding="1" cellspacing="2" align="center"> 
        <form name=a1 method="post" action="?action=save&postid=<%=PostId%>" onsubmit="return checkData()"> 
          <tr valign="middle"> 
            <td bgcolor="#dddddd" height="30" colspan="4">&nbsp;<b>会员投稿</b></td>
          </tr> 
          <tr valign="middle" height="25"> 
            <td width="15%" align="center" bgcolor="#e6f0ff">所属栏目</td> 
            <td align="left" colspan="3" >&nbsp;<select name="Column" class="LoginInput"> 
                <option value="0">请选择...</option> 
                <%
					ColumnList=EA_DBO.Get_MemberAppearColumnList()
					
					If IsArray(ColumnList) Then 
						For i=0 To UBound(ColumnList,2)
							Level=(Len(ColumnList(2,i))/4-1)*3
							Response.Write "<option value="""&ColumnList(0,i)&"|||"&ColumnList(1,i)&"|||"&ColumnList(2,i)&""""
							If ColumnList(0,i)=ColumnId Then Response.Write " selected"
							Response.Write ">"
							If Len(ColumnList(2,i))>4 Then Response.Write "├"
							Response.Write String(Level,"-")
							Response.Write ColumnList(1,i)&ColumnList(3,i)&"</option>"
						Next
					End If
				%> 
              </select></td>
          </tr>
          <tr valign="middle" height="25"> 
            <td width="15%" align="center" bgcolor="#e6f0ff">标题</td> 
            <td colspan="3" align="left" bgcolor="#FFFFFF">&nbsp;<input name="title" type="text" size=50 class="LoginInput" value="<%=Title%>"></td> 
          </tr>
          <tr valign="middle" height="25"> 
            <td width="15%" align="center" bgcolor="#e6f0ff">关键字</td> 
            <td align="left" bgcolor="ffffff" width="50%">&nbsp;<input name="keyword" size=40 type="text" class="LoginInput" value="<%=Keyword%>">&nbsp;以,号分隔</td>
            <td width="20%" align="center" bgcolor="#e6f0ff">作者</td> 
            <td align="left" bgcolor="#FFFFFF">&nbsp;<font color=800000><%=EA_Pub.Mem_Info(1)%></font></td> 
          </tr>
          <tr valign="middle" height="25"> 
            <td width="15%" align="center" bgcolor="#e6f0ff">标题图片</td> 
            <td colspan="3" align="left" bgcolor="#FFFFFF">&nbsp;<select name="d_picture" class="LoginInput"> 
                <option value=''>暂无图片</option> 
              </select>&nbsp;<input name="img" type="text" size=30 class="LoginInput" value="<%=ImgPath%>">&nbsp;<a href="#" onclick="review_img();">[预览]</a></td> 
          </tr>
          <tr valign="middle" height="25"> 
            <td width="15%" align="center" bgcolor="#e6f0ff">文章来源</td> 
            <td align="left" bgcolor="ffffff" colspan="3">&nbsp;<select name="choosesource" class="LoginInput" onchange="inputsource(this.options[this.selectedIndex].value)">
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
      </select>&nbsp;=->&nbsp;<input name="source" type="text" value="<%=Source%>" class="LoginInput">&nbsp;<input name="sourceurl" class="LoginInput" type="text" value="<%=SourceUrl%>" size=30></td> 
          </tr>
          <tr valign="middle"> 
            <td width="15%" align="center" bgcolor="#e6f0ff">正文</td> 
            <td colspan="3" align="left" bgcolor="ffffff">&nbsp;<%=EA_Editor.Create%></td> 
          </tr>
          <tr valign="middle"> 
            <td width="15%" align="center" bgcolor="#e6f0ff">摘要</td> 
            <td colspan="3" align="left" bgcolor="ffffff">&nbsp;<textarea name="summary" cols="70" rows="5" wrap="VIRTUAL"><%=EA_Pub.Un_Full_HTMLFilter(Summary)%></textarea><br>&nbsp;<a href="javascript:change(document.all.summary,-20)"><img src="../images/public/minus.gif" border=0 title="缩小"></a>&nbsp;&nbsp;<a href="javascript:change(document.all.summary,20)"><img src="../images/public/plus.gif" border=0 title="放大"></a></td> 
          </tr> 
          <tr> 
            <td colspan="4" align="center" valign="middle" height="25" bgcolor="#efefef"><input type="hidden" name="author" value="<%=EA_Pub.Mem_Info(0)%>"> 
              <input type="submit" name="Submit1" value=" 提交我的文章 ">&nbsp;<input type="reset" name="Submit2" value="重置"></td> 
          </tr> 
        </form> 
      </table></td> 
  </tr> 
</table> 
<%

Sub Save_MemberAppear()
	Call EA_Pub.Chk_Post
	
	Dim PostId,TempStr
	Dim Key
	Dim Feedback
	ReDim ArticleInfo(18)
	
	If Request.Form("column")="" Or Request.Form("column")="0" Then 
		ErrMsg="传递错误的栏目数据！"
		Call EA_Pub.ShowErrMsg(0,2)
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
		ErrMsg="标题内容长度不符。"
		ErrMsg=ErrMsg&"\n大于150或者等于0个字符"
		FoundErr=True
	End If
	If Len(ArticleInfo(2))>20 Then 
		ErrMsg="关键字长度不符。"
		ErrMsg=ErrMsg&"\n大于20个字符"
		FoundErr=True
	End If
	If Len(ArticleInfo(10))>250 Then 
		ErrMsg="文章简介长度不符。"
		ErrMsg=ErrMsg&"\n大于250个字符"
		FoundErr=True
	End If
	If Not EA_DBO.Get_Column_Info(ArticleInfo(3))(15,0) Then 
		ErrMsg="该栏目不允许投稿"
		FoundErr=True
	End If
	
	If FoundErr Then Call EA_Pub.ShowErrMsg(0,2)

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
	
	Feedback=EA_DBO.Set_Article_Insert(PostId,ArticleInfo)
	
	Select Case Feedback
	Case -1
		ErrMsg="在更新数据库过程中发生错误，操作已取消。"

		Call EA_Pub.ShowErrMsg(0,2)
	Case 1
		ErrMsg="文章已发布，正等待管理员审核。"
		Application.Lock 
		Application(sCacheName&"IsFlush")=1
		Application.UnLock 

		Call EA_Pub.ShowSusMsg(6,"")
	Case 0
		ErrMsg="文章已发布，现可马上查看。"
		Application.Lock 
		Application(sCacheName&"IsFlush")=1
		Application.UnLock 

		Call EA_Pub.ShowSusMsg(7,"")
	End Select
End Sub
%> 
