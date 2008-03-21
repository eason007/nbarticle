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

function friend (siteUrl) {
	var html = "";

	html = "<table>";
	html += "<form name=\"form1\" method=\"post\" action=\"" + siteUrl + "action.asp?action=link\">"
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
	html += "</form>";
	html += "</table>";

	return html;
}

function comment (iArticleID, siteUrl) {
	var html = "";

	html = "<table>";
	html += "<form name=\"form1\" method=\"post\" action=\"" + siteUrl + "action.asp?action=comment\">"
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
	html += "</form>";
	html += "</table>";

	return html;
}

function siupIn (siteUrl) {
	var html = "";

	html = "<table>";
	html += "<form name=\"form1\" method=\"post\" action=\"" + siteUrl + "member/login.asp?action=login\">"
	html += "<tr>"
	html += "<td align=\"right\">用户名：</td><td align=\"left\"><input type=\"text\" name=\"UserName\" style=\"width: 130px;\" /></td>";
	html += "</tr>"
	html += "<tr>"
	html += "<td align=\"right\">密码：</td><td align=\"left\"><input type=\"password\" name=\"Password\" style=\"width: 130px;\" /></td>";
	html += "</tr>"
	html += "<tr>"
	html += "<td align=\"right\"><input type=\"submit\" name=\"Submit\" value=\"提交\" /></td><td align=\"left\"><input type=\"checkbox\" name=\"SaveTimes\" value=\"10\" />自动登陆</td>";
	html += "</tr>"
	html += "<tr>"
	html += "<td align=\"left\" colspan=\"2\">[<a href=\"javascript: vod();\" onclick=\"window.open('" + siteUrl + "member/register.asp','','width=790,height=400')\">注册</a>] - [<a href=\"javascript: vod();\" onclick=\"window.open('" + siteUrl + "member/getpass.asp','','scrollbars=no,width=650,height=150')\">忘记密码</a>]</td>";
	html += "</tr>"
	html += "</form>";
	html += "</table>";

	return html;
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


EliteCMS = {
	isMember : false,
	memberData : "",
	memberInfo : Array(),
	windows: "<div id=\"EliteWindow\" style=\"z-index: 99; position: absolute; top: 200px; left: 400px; border: #DBE1E9 1px solid; background: #fff; padding: 10px 30px;\"></div>",

	init : function () {
		this.memberData = getCookie('UserData');

		if (this.memberData.length > 0) {
			this.isMember = true;
			this.memberInfo = this.memberData.split("|");
		}
		else {
			this.isMember = false;
		}
	},

	showMember : function () {
		if (this.isMember)
		{
			document.write (this.memberInfo[1] + "，你好 - <a href=\"\">退出");
		}
		else{
			document.write ("<a href=\"javascript: vod();\" onclick=\"EliteCMS.showWindow(siupIn())\">登陆</a> | 注册");
		}
	},

	showWindow : function (op) {
		if (!$("EliteWindow"))
		{
			document.body.innerHTML += this.windows;
		}
		else{
			$("EliteWindow").innerHTML = "";
		}

		$("EliteWindow").innerHTML = "<a href=\"javascript: vod();\" onclick=\"$('EliteWindow').style.display = 'none';\" class=\"left\">[X]</a>" + op;
		$("EliteWindow").style.display = "";
	}
}

EliteCMS.init();
