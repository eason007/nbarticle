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

function postFriend () {
	var html = "";

	html = "<table>";
	html += "<form name=\"form1\" method=\"post\" action=\"action.asp?action=link\">"
	html += "<tr>"
	html += "<td align=\"right\">站点名称：</td><td align=\"left\"><input type=\"text\" name=\"name\"></td>";
	html += "</tr>"
	html += "<tr>"
	html += "<td align=\"right\">站点Logo：</td><td align=\"left\"><input type=\"text\" name=\"logo\" value=\"http://\"></td>";
	html += "</tr>"
	html += "<tr>"
	html += "<td align=\"right\">站点URL：</td><td align=\"left\"><input type=\"text\" name=\"url\" value=\"http://\"></td>";
	html += "</tr>"
	html += "<tr>"
	html += "<td align=\"right\">站点简介：</td><td align=\"left\"><textarea name=\"info\" wrap=\"VIRTUAL\" cols=\"50\" rows=\"5\"></textarea></td>";
	html += "</tr>"
	html += "<tr>"
	html += "<td align=\"right\">显示风格：</td><td align=\"left\"><input type=\"radio\" value=\"0\" name=\"style\">文本<input type=\"radio\" value=\"1\" name=\"style\">图片</td>";
	html += "</tr>"
	html += "<tr>"
	html += "<td align=\"left\" colspan=\"2\"><input type=\"submit\" name=\"Submit\" value=\"提交\"></td>";
	html += "</tr>"
	html += "</form>";
	html += "</table>";

	return html;
}