function vod() {
	//空函数
}

/*处理脚本部份 begin*/
function _executeScript(scriptFrag) {
	var scriptContainerId = "_SCRIPT_CONTAINER";
	var obj = $(scriptContainerId);
	if (obj != null) {
		document.body.removeChild(obj);
	}
	var scriptContainer = document.createElement('SCRIPT');
	scriptContainer.setAttribute("id", scriptContainerId);
	scriptContainer.text = scriptFrag;
	document.body.appendChild(scriptContainer);
} 
/*处理脚本部份 end*/

String.prototype.trim = function() {
//去掉字符串头尾空格
	var result=this.replace(/(^\s*)/g, "");
	result=result.replace(/(\s*$)/g, "");
	return result;
}

ie = (document.all)? true:false
if (ie){
//快捷提交表单-仅限IE(Ctrl+Enter)
	function ctlent(eventobject){
		if(event.ctrlKey && window.event.keyCode==13){
			this.document.form1.submit();
		}
	}
}

function $() {
	var elements = new Array();

	for (var i = 0; i < arguments.length; i++) {
		var element = arguments[i];

		if (typeof element == 'string')
			element = document.getElementById(element);

		if (arguments.length == 1)
			return element;

		elements.push(element);
	}

	return elements;
}

function PageNav(vPageNo,vShowTypes){
	var Nav="";

	Nav = '<div id="pageList">';

	if (vPageNo-4<1){
		var RootPage=1;
	}
	else{
		var RootPage=vPageNo-4;
	}
	if (vPageNo+4>PUB_PageCount){
		var EndPage=PUB_PageCount;
	}
	else{
		var EndPage=vPageNo+4;
	}

	if (vPageNo > 1){
		Nav+='<a href="javascript:vod();" onclick="javascript:ShowContentList('+1+','+vShowTypes+')" title="首页" class="first">&laquo;</a>&nbsp;';
		Nav+='<a href="javascript:vod();" onclick="javascript:ShowContentList('+(vPageNo-1)+','+vShowTypes+')" title="上一页" class="list">&lt;</a>&nbsp;';
	}

	for (i=RootPage;i<=EndPage;i++){
		if (i==vPageNo){
			Nav+='<span class="current">['+i+']</span>&nbsp;';
		}
		else{
			Nav+='<a href="javascript:vod();" onclick="javascript:ShowContentList('+i+','+vShowTypes+')" class="list">['+i+']</a>&nbsp;';
		}
	}

	if (vPageNo < PUB_PageCount){
		Nav+='<a href="javascript:vod();" onclick="javascript:ShowContentList('+(vPageNo+1)+','+vShowTypes+')" title="下一页" class="list">&gt;</a>&nbsp;';
		Nav+='<a href="javascript:vod();" onclick="javascript:ShowContentList('+PUB_PageCount+','+vShowTypes+')" title="尾页" class=\"last\">&raquo;</a>&nbsp;';
	}

	Nav+='<span class="total">' + vPageNo+'&nbsp;/&nbsp;'+PUB_PageCount+' 页</span>&nbsp;';

	Nav+='</div>';

	return Nav;
}

function ShowContentList(vPage,ShowTypes){
	var StarId=(vPage-1)*PUB_PageSize;
	var Tmp,TmpStr;
	TmpStr="";

	if ((StarId+PUB_PageSize)-1>(RCount-1)){
		var EndId=RCount-1;
	}
	else{
		var EndId=(StarId+PUB_PageSize)-1;
	}

	for (i=StarId;i<=EndId;i++){
		Tmp=ReviewList[i].split("||");
		TmpStr+='<table width="100%" align="center" id="feedback">';
		if (ShowTypes==0)
		{
			TmpStr+='<tr>';
			TmpStr+='<td width="10%" align="right">发言人:</td>';
			TmpStr+='<td width="40%" align="left">&nbsp;'+Tmp[0]+'</td>';
			TmpStr+='<td width="10%" align="right">发言时间:</td>';
			TmpStr+='<td width="40%" align="left">&nbsp;'+Tmp[1]+'</td>';
			TmpStr+='</tr>';
			TmpStr+='<tr>';
			TmpStr+='<td align="right" valign="top">内容:</td>';
			TmpStr+='<td colspan="3">'+Tmp[2]+'</td>';
			TmpStr+='</tr>';
		}
		else{
			TmpStr+='<tr>';
			TmpStr+='<td width="10%" align="left" valign="top">'+Tmp[0]+'</td>';
			TmpStr+='<td width="30%" align="left" valign="top">'+Tmp[1]+'</td>';
			TmpStr+='<td>'+Tmp[2].substring(0,50)+'</td>';
			TmpStr+='</tr>';
		}
		TmpStr+='</table>';
	}

	TmpStr+=PageNav(vPage,ShowTypes);

	$('CommentList').innerHTML=TmpStr;
}

function submit_vote(vote_id){
//投票处理函数
	var vote_form=$('vote_'+vote_id);

	var target_url='vote.asp?votetype='+vote_form.votetype.value;
	target_url+='&voteid='+vote_form.voteid.value;
	target_url+='&vote=';

	for (var i=0;i<vote_form.vote.length ; i++)
	{
		if (vote_form.votetype.value==0)
		{
			if (vote_form.vote[i].checked==true)
			{
				target_url+=vote_form.vote[i].value+',';
			}
		}
		else{
			if (vote_form.vote[i].checked==true)
			{
				target_url+=vote_form.vote[i].value+',';
			}
		}
	}
	target_url=target_url.substring(0,target_url.length-1);

	return target_url;
}

function getHtml2Div (sUrl) {
	ajaxGetDateToPage(sUrl, "divAjax");
}

document.write ("<div id='divAjax' style='display:none;'></div>");