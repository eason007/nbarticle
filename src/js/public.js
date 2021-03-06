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

function getCookie(cName){
	var cValue="";
	var cName=cName+"=";

	if(document.cookie.length>0){ 
		offset=document.cookie.indexOf(cName);
		if(offset!=-1){ 
			offset+=cName.length;
			end=document.cookie.indexOf(";",offset);
			if(end==-1) {
				end=document.cookie.length;
			}
			cValue=decodeURI(document.cookie.substring(offset,end))
		}
	}

	return cValue;
}

function LoadJS(id, fileUrl)
{
	var scriptTag = $(id);
	var oHead = document.getElementsByTagName('HEAD').item(0);
	var oScript= document.createElement("script");

	if (scriptTag)oHead.removeChild(scriptTag);

	oScript.id = id;
	oScript.type = "text/javascript";
	oScript.src=fileUrl ;
	oHead.appendChild(oScript);
}

function submit_vote(vote_id){
//投票处理函数
	var vote_form=$('vote_'+vote_id);

	var target_url='&votetype='+vote_form.votetype.value;
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

function friend () {
	var html = "";

	html = "<form name=\"form1\" method=\"post\" action=\"" + EliteCMS.basePath + "action.asp?action=link\">"
	html += "<table>";
	html += "<tr>"
	html += "<td align=\"right\">站点名称：</td><td align=\"left\"><input type=\"text\" name=\"name\" /></td>";
	html += "</tr>"
	html += "<tr>"
	html += "<td align=\"right\">站点Logo：</td><td align=\"left\"><input type=\"text\" name=\"logo\" value=\"http://\" /></td>";
	html += "</tr>"
	html += "<tr>"
	html += "<td align=\"right\">站点URL：</td><td align=\"left\"><input type=\"text\" name=\"url\" value=\"http://\" /></td>";
	html += "</tr>"
	html += "<tr>"
	html += "<td align=\"right\">站点简介：</td><td align=\"left\"><textarea name=\"info\" wrap=\"VIRTUAL\" cols=\"50\" rows=\"5\"></textarea></td>";
	html += "</tr>"
	html += "<tr>"
	html += "<td align=\"right\">显示风格：</td><td align=\"left\"><input type=\"radio\" value=\"0\" name=\"style\" />文本<input type=\"radio\" value=\"1\" name=\"style\" />图片</td>";
	html += "</tr>"
	html += "<tr>"
	html += "<td align=\"left\" colspan=\"2\"><input type=\"submit\" name=\"Submit\" value=\"提交\" /></td>";
	html += "</tr>"
	html += "</table>";
	html += "</form>";

	return html;
}

function comment (iArticleID) {
	var html = "";

	html = "<form name=\"form1\" method=\"post\" action=\"" + EliteCMS.basePath + "action.asp?action=comment\">"
	html += "<table>";
	html += "<input type=\"hidden\" name=\"articleid\" value=\"" + iArticleID + "\" />"
	html += "<tr>"
	html += "<td align=\"right\">笔名：</td><td align=\"left\"><input type=\"text\" name=\"name\" /></td>";
	html += "</tr>"
	html += "<tr>"
	html += "<td align=\"right\">评论：</td><td align=\"left\"><textarea name=\"review\" wrap=\"VIRTUAL\" cols=\"30\" rows=\"5\"></textarea></td>";
	html += "</tr>"
	html += "<tr>"
	html += "<td align=\"left\" colspan=\"2\"><input type=\"submit\" name=\"Submit\" value=\"提交\" /></td>";
	html += "</tr>"
	html += "</table>";
	html += "</form>";

	return html;
}

function siupIn () {
	var html = "";

	html = "<form name=\"form1\" method=\"post\" action=\"" + EliteCMS.basePath + "member/login.asp?action=login\">"
	html += "<table style=\"border: #A9D5F4 1px solid; width: 250px\">";
	html += "<tr>"
	html += "<td height=\"25\" colspan=\"2\" bgcolor=\"#DBF2FF\" align=\"center\"><strong>会员登陆</strong></td>";
	html += "</tr>"
	html += "<tr>"
	html += "<td align=\"right\">帐号：</td><td align=\"left\"><input type=\"text\" name=\"UserName\" style=\"width: 130px;\" /></td>";
	html += "</tr>"
	html += "<tr>"
	html += "<td align=\"right\">密码：</td><td align=\"left\"><input type=\"password\" name=\"Password\" style=\"width: 130px;\" /></td>";
	html += "</tr>"
	html += "<tr>"
	html += "<td align=\"right\"><input type=\"submit\" name=\"Submit\" value=\"登陆\" /></td><td align=\"left\"><input type=\"checkbox\" name=\"SaveTimes\" value=\"10\" />自动登陆</td>";
	html += "</tr>"
	html += "</table>";
	html += "</form>";

	return html;
}

EliteCMS = {
	basePath : "",
	isMember : false,
	memberData : "",
	memberInfo : Array(),
	windows: "<div id=\"EliteBox\" style=\"z-index: 99; position: absolute; top: 50px; left: 100px; border: #DBE1E9 1px solid; background: #fff; padding: 15px 20px;\"><a href=\"javascript: vod();\" onclick=\"$('EliteBox').style.display = 'none';\" title=\"关闭\">[X]</a><div id=\"EliteWindow\"></div></div>",

	init : function () {
		this.memberData = getCookie('UserData');

		if (this.memberData.length > 0) {
			this.isMember = true;
			this.memberInfo = this.memberData.split("|");
		}
		else {
			this.isMember = false;
		}

		LoadJS("objAjax", this.basePath + "js/objAjax.js");
	},

	showMember : function () {
		if (this.isMember)
		{
			document.write (this.memberInfo[1] + "，你好 - [<a href=\"javascript: vod();\" onclick=\"window.open('" + EliteCMS.basePath + "member/appear.asp','_blank','scrollbars=yes,width=1030,height=580');\">投稿</a>] [<a href=\"javascript: vod();\" onclick=\"EliteCMS.showWindow2File('" + EliteCMS.basePath + "member/myappear.asp')\">投稿箱</a>] [<a href=\"javascript: vod();\" onclick=\"EliteCMS.showWindow2File('" + EliteCMS.basePath + "member/favlist.asp')\">收藏夹</a>] [<a href=\"javascript: vod();\" onclick=\"EliteCMS.showWindow2File('" + EliteCMS.basePath + "member/changepwd.asp')\">修改密码</a>] [<a href=\"javascript: vod();\" onclick=\"EliteCMS.showWindow2File('" + EliteCMS.basePath + "member/changecase.asp')\">修改资料</a>] [<a href=\"" + this.basePath + "member/login.asp?action=logout\">退出</a>]");
		}
		else{
			document.write ("<a href=\"javascript: vod();\" onclick=\"EliteCMS.showWindow(siupIn())\">登陆</a> | <a href=\"javascript: vod();\" onclick=\"EliteCMS.showWindow2File('" + EliteCMS.basePath + "member/register.asp')\">注册</a>");
		}
	},

	showWindow : function (op) {
		if (!$("EliteBox"))
		{
			document.body.innerHTML += this.windows;
		}
		else{
			$("EliteWindow").innerHTML = "";
		}

		$("EliteWindow").innerHTML = op;
		$("EliteBox").style.display = "";
	},

	showWindow2File : function (sURL) {
		if (!$("EliteBox"))
		{
			document.body.innerHTML += this.windows;
		}
		else{
			$("EliteWindow").innerHTML = "";
			$("EliteBox").style.display = "";
		}

		ajaxGetDateToPage(sURL, "EliteWindow");
	}
}

EliteCMS.init();
