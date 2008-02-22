// 帮助实际内容
var help_topic_desc = "";
// 帮助收缩后显示的内容
var help_topic_desc0 = "<div class='help_topic_desc_hidden' onclick='sw_help_topic()'>——" + helpTitle + "&nbsp;&nbsp</div>";
// 帮助title的内容
var help_topic_title = "";

function sw_help_topic_init(title)
{
	var obj = $("help_topic");
	var obj_icon = $("help_topic_icon");

	help_topic_desc = obj.innerHTML;
	help_topic_title = obj_icon.innerHTML;
	
	// 标题样式更新
	obj_icon.className = "help_topic_title";

	sw_help_topic_ex(0);
}

function sw_help_topic()
{
	var obj = $("help_topic");
	var obj_icon = $("help_topic_icon");

	if (obj.style.height == "20px")
	{
		b_status = 1;
	}
	else
	{
		b_status = 0;
	}

	sw_help_topic_ex(b_status);
}

function sw_help_topic_ex(b_status)
{
	var obj = $("help_topic");
	var obj_icon = $("help_topic_icon");

	if (b_status == 1)
	{ 
		// 打开
		obj_icon.innerHTML = "- " + help_topic_title;
		obj.style.height = "";
		obj.innerHTML = help_topic_desc;
		obj.className = "help_topic_desc";
	}       
	else
	{
		// 关闭
		obj_icon.innerHTML = "+ " + help_topic_title;
		obj.style.height = "20px";
		obj.innerHTML = help_topic_desc0;
		obj.className = "help_topic_desc_hidden";
	}

	try 
	{
		parent.iframeResize();       
	}
	catch(e)
	{
	}
}


function CheckAll(form,objTag)
{
	for (var i=0;i<form.elements.length;i++)
	{
		var e = form.elements[i];
		if (e.type=='checkbox' && e.name !='chkall'){
			e.checked = objTag.checked;
		}
	}
}


/*更改右边内容*/
function ajaxChangeRightContent(pageUrl, pageType, actionMessage, isPageForceLoad, postData) {
	if (pageType == "pageAdd")
	{
		delete tabFormID[currentTab];
	}

	$("divLoading").style.display = "";
	$('divDialog').style.display = "none";

	aoControl = new ajaxObject();
	aoControl.URL = pageUrl;
	aoControl.doResponseMethod = ajaxPageContent;
	aoControl.actionMessage = actionMessage;

	if (typeof(postData) != "undefined")
	{
		aoControl.dateVal = postData;
		aoControl.ajaxPostDate(aoControl);
	}
	else{
		aoControl.ajaxGetDate(aoControl);
	}
}

//内容插入页面中
function ajaxPageContent(obj) {
	message = obj.returnMessage;
	if(message != "") {
		if (message == "500")
		{
			alertMessage(noExtentOfAuthority);
			return false;
		}

		actionMessage = obj.actionMessage;

		openTab(message);

		if (actionMessage != "" && actionMessage)
		{
			alertMessage(actionMessage);
		}
	}	
}

//插入内部脚本
function executeScript (message){
	var regexp2 = /<script(.|\n)*?>((.|\n|\r\n)*)?<\/script>/im;

	if (message) {
		for (var i = 0; i < message.length; i++) {
			/* Note: do not try to write more than one <script> in your view.*/
			/* break;  process only one script element */
			var realScript = message[i].match(regexp2);
			_executeScript(realScript[2]);
		}
	}
	else{
		message = "";
	}

	var argc = executeScript.arguments.length;
	var argv = executeScript.arguments;

	if (argc > 1)
	{
		tabScriptCache[currentTab] = message;

		if (argv[1])
		{
			initTabContent();
		}
	}
}

/*修改信息*/
function edit(ID, url) {
	tabFormID[currentTab] = ID;
    ajaxChangeRightContent(url);

    return false;
}

/*删除信息*/
function del(url, formid, name) {
    if(!formid){
		formid = 'form1';
	}
	var form=$(formid);

	if(!name){
		name="checkbox";
	}
	var j=0;var str="";
	for(var i=0;i<form.elements.length;i++){
		if(form.elements[i].type == "checkbox" && form.elements[i].checked && form.elements[i].name == name) {
			j++;
			str += form.elements[i].value + ",";
		}
	}

	if (j==0){
		alertMessage(noChoose);
		return false;
	}
	else{
		str = str.substr(0, str.length-1);
		if (confirmMessage(delConfim))
		{
			objForm.formName = formid;
			objForm.postUrl = url;
			objForm.dateVal = "action=del&ID="+str;
			objForm.doResponseMethod = getDeletePostResult;
			objForm.ajaxPost();
			objTable.start();
		}
	}

    return false;
}

/*自定义表单提交*/
function postChooseForm(url, action, resultFunction, formid, name) {
    if(!formid){
		var form=$('form1');
	}else{
		var form=$(formid);
	}

	if(!name){
		name="checkbox";
	}
	var j=0;var str="";
	for(var i=0;i<form.elements.length;i++){
		if(form.elements[i].type == "checkbox" && form.elements[i].checked && form.elements[i].name == name) {
			j++;
			str += form.elements[i].value + ",";
		}
	}

	if (j==0){
		alertMessage(noChoose);
		return false;
	}
	else{
		str = str.substr(0, str.length-1);

		objForm.postUrl = url;
		objForm.dateVal = "action="+action+"&ID="+str;
		objForm.doResponseMethod = resultFunction;
		objForm.ajaxPost();
		objTable.start();
	}

    return false;
}

/*修改状态结果*/
function getStatePostResult(returnResult) {
	if(returnResult >= 0) {
		alertMessage(channgeSuccess);
		if(objTable){
			objTable.start();
		}
	}
	else {
		alertMessage(channgeFail);
	}
}

/*删除内容结果*/
function getDeletePostResult(returnResult) {
	if(returnResult >= 0) {
		alertMessage(delSuccess);
		if(objTable){
			if(objTable.total-objTable.prePage*(objTable.nowPage-1)==1 && objTable.nowPage>1){
				objTable.nowPage--;
			}
			objTable.start();
		}
	}
	else {
		alertMessage(delFail);
	}
}

/*修改内容结果*/
function getEditPostResult(returnResult) { 
	if(returnResult >= 0) {
		ajaxChangeRightContent(editSuccessUrl, 'pageContent', editSuccess);	
	}
	else {
		alertMessage(editFail);
	}
}

/*添加信息结果*/
function getAddPostResult(returnResult) {
	if(returnResult >= 0) {
		alertMessage(addSuccess);
	}
	else{
		alertMessage(addFail);
	}
	objForm.resetForm();
}

function alertMessage(sMsg){
	//alert(sMsg);
	//loadDialog(sMsg);
	$('dialogBox').innerHTML = sMsg;
	$('divDialog').style.display = "";
}

function confirmMessage(sMsg){
	flag = confirm(sMsg);

	if (flag == true)
	{
		return 1;
	}
	else{
		return 0;
	}
}

function change(obj,i) {
	he=parseInt(obj.style.height);

	if (he>=80&&he<=400){
	   obj.style.height=he+i+'px';
	}
	else {
	   obj.style.height='150px';
	}
}


/*
冷情圣郎
2004年10月9日
HTML代码过滤涵数
*/
function cleanHtml(ReContents)
{
	//清理多余HTML代码
	ReContents = ReContents.replace(/<p>&nbsp;<\/p>/gi,"")
	ReContents = ReContents.replace(/<p><\/p>/gi,"")
	ReContents = ReContents.replace(/<p>/,"")
	ReContents = ReContents.replace(/<\/p>/,"")
	ReContents = ReContents.replace(/<li>/,"")
	ReContents = ReContents.replace(/<lu>/,"")
	ReContents = ReContents.replace(/(<(meta|iframe|frame|span|tbody|layer)[^>]*>|<\/(iframe|frame|meta|span|tbody|layer)>)/gi, "");
	ReContents = ReContents.replace(/<\\?\?xml[^>]*>/gi, "") ;
	ReContents = ReContents.replace(/o:/gi, "");
	ReContents = ReContents.replace(/ /gi, "");
	ReContents = ReContents.replace(/&nbsp;/gi, " ");
	ReContents = ReContents.replace(/(<(style|strong)[^>]*>|<\/(style|strong)>)/gi, "");
	//验证空白行
	ReContents = ReContents.replace(/^\[ \t]*$/,"")
	//表格也要过滤！
	ReContents = ReContents.replace(/(<(table|tbody|tr|td|th|)[^>]*>|<\/(table|tbody|tr|td|th|)>)/gi, "");
	//图片过滤
	ReContents = ReContents.replace(/(<(img)[^>]*>|<\/(img)>)/gi, "");
	//<div>过滤
	ReContents = ReContents.replace(/(<(div|blockquote|fieldset|legend)[^>]*>|<\/(div|blockquote|fieldset|legend)>)/gi, "");
	//<font>过滤
	ReContents = ReContents.replace(/(<(font|i|u|h[1-9]|s)[^>]*>|<\/(font|i|u|h[1-9]|s)>)/gi, "");
	//过滤脚本
	ReContents = ReContents.replace(/(<script[^>]*>|<\/script>)/gi, "");
	//去掉任何标记中的任何事件！
	RegExp = /<(\w[^>|\s]*)([^>]*)(on(finish|mouse|Exit|error|click|key|load|change|focus|blur))(.[^>]*)/gi;
	ReContents = ReContents.replace(RegExp, "<$1")
	RegExp = /<(\w[^>|\s]*)([^>]*)(&#|window\.|id|javascript:|js:|about:|file:|Document\.|vbs:|cookie| name| id)(.[^>]*)/gi;
	ReContents = ReContents.replace(RegExp, "<$1")

	ReContents = ReContents.replace(/\n/gi, "");
	ReContents = ReContents.replace(/\r/gi, "");

	return ReContents;
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