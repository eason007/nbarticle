// 帮助实际内容
var help_topic_desc = "";
// 帮助收缩后显示的内容
var help_topic_desc0 = "<div class='help_topic_desc_hidden' onclick='sw_help_topic()'>——" + helpTitle + "&nbsp;&nbsp</div>";
// 帮助title的内容
var help_topic_title = "";

function sw_help_topic_init(title)
{
	var obj = document.getElementById("help_topic");
	var obj_icon = document.getElementById("help_topic_icon");

	help_topic_desc = obj.innerHTML;
	help_topic_title = obj_icon.innerHTML;
	
	// 标题样式更新
	obj_icon.className = "help_topic_title";

	sw_help_topic_ex(0);
}

function sw_help_topic()
{
	var obj = document.getElementById("help_topic");
	var obj_icon = document.getElementById("help_topic_icon");

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
	var obj = document.getElementById("help_topic");
	var obj_icon = document.getElementById("help_topic_icon");

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

	document.getElementById("divLoading").style.display = "";

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
		var form=document.getElementById('form1');
	}else{
		var form=document.getElementById(formid);
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
		if (confirmMessage(delConfim))
		{
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
		var form=document.getElementById('form1');
	}else{
		var form=document.getElementById(formid);
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
		objForm.resetForm();
	}
	else{
		alertMessage(addFail);
	}
}

function alertMessage(sMsg){
	//alert(sMsg);
	loadDialog(sMsg);
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


var NUMBER_OF_REPETITIONS = 40;
var nRepetitions = 0;
var g_oTimer = null;

function loadDialog(sDialogContent){
	var dTitle = document.getElementById("DialogTitle");
	dTitle.innerHTML = "<img src=\"images/spacer.gif\" alt=\"\" />&nbsp;" + DialogTitle;

	document.getElementById("DialogContent").innerHTML = sDialogContent;
	document.getElementById("DialogCancel").value = " " + closeDialogValue + " ";

	var Dialog = document.getElementById("divDialog");
	Dialog.style.display = "";
	Dialog.focus();
	resizeModal();

	// Add a resize handler for the window
	window.onresize = resizeModal;
	// Add a warning in case anyone tries to navigate away or refresh the page
	window.onbeforeunload = showWarning;

	if (!isMozilla){
		Dialog.filters.alpha.opacity = 0 + nRepetitions*5;
		showDialog();
	}
}

function showDialog()
{
	if (nRepetitions < NUMBER_OF_REPETITIONS)
	{
		// Set the timeout somewhere between 0 and .25 seconds
		var nTimeoutLength = Math.random() * 200;

		document.getElementById("divDialog").filters.alpha.opacity=0 + nRepetitions*10;
		g_oTimer = window.setTimeout("showDialog();", nTimeoutLength);
		nRepetitions++;
	}
	else
	{
		var nTimeoutLength = Math.random() * 250;

		document.getElementById("divDialog").filters.alpha.opacity-=10;

		g_oTimer = window.setTimeout("showDialog();", nTimeoutLength);
		nRepetitions++;

		if (nRepetitions==(NUMBER_OF_REPETITIONS*2) || document.getElementById("divDialog").filters.alpha.opacity==0)
		{
			closeDialog();
		}
	}
}

function closeDialog()
{
	if (g_oTimer != null)
	{
		// Clear the timer so we don't get called back an extra time
		window.clearTimeout(g_oTimer);
		g_oTimer = null;
	}

	// Hide the fake modal DIV
	document.getElementById("divModal").style.width = "0px";
	document.getElementById("divModal").style.height = "0px";
	document.getElementById("divDialog").style.display = "none";

	// Remove our event handlers
	window.onresize = null;
	window.onbeforeunload = null;

	nRepetitions = 0;
}

function showWarning()
{
	//Warn users before they refresh the page or navigate away
	return "";
}

function resizeModal()
{
	pSize = getPageSize();

	if (pSize[1] > pSize[3])
	{
		var divHeight = pSize[1];
	}
	else{
		var divHeight = pSize[3];
	}
	var divWidth = pSize[2];
	
	// Resize the DIV which fakes the modality of the dialog DIV
	document.getElementById("divModal").style.width = divWidth + "px";
	document.getElementById("divModal").style.height = divHeight + "px";

	// Re-center the dialog DIV
	document.getElementById("divDialog").style.left = ((divWidth - 350) / 2) + "px";
	document.getElementById("divDialog").style.top = ((divHeight - 200) / 2) + "px";
}

function getPageSize(){
	var xScroll, yScroll;

	if (window.innerHeight && window.scrollMaxY) {
		xScroll = document.body.scrollWidth;
		yScroll = window.innerHeight + window.scrollMaxY;
	}
	else if (document.body.scrollHeight > document.body.offsetHeight){ // all but Explorer Mac
		xScroll = document.body.scrollWidth;
		yScroll = document.body.scrollHeight;
	} 
	else { // Explorer Mac...would also work in Explorer 6 Strict, Mozilla and Safari
		xScroll = document.body.offsetWidth;
		yScroll = document.body.offsetHeight;
	}

	var windowWidth, windowHeight;
	if (self.innerHeight) {  // all except Explorer
		windowWidth = self.innerWidth;
		windowHeight = self.innerHeight;
	}
	else if (document.documentElement && document.documentElement.clientHeight) { // Explorer 6 Strict Mode
		windowWidth = document.documentElement.clientWidth;
		windowHeight = document.documentElement.clientHeight;
	}
	else if (document.body) { // other Explorers
		windowWidth = document.body.clientWidth;
		windowHeight = document.body.clientHeight;
	}  

	// for small pages with total height less then height of the viewport
	if(yScroll < windowHeight){
		pageHeight = windowHeight;
	}
	else { 
		pageHeight = yScroll;
	}

	if(xScroll < windowWidth){  
		pageWidth = windowWidth;
	}
	else {
		pageWidth = xScroll;
	}

	arrayPageSize = new Array(pageWidth,pageHeight,windowWidth,windowHeight);
	return arrayPageSize;
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