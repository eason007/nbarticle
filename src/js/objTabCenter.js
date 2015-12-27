var isMozilla = (typeof document.implementation != 'undefined') && (typeof document.implementation.createDocument != 'undefined') && (typeof HTMLDocument!='undefined');

if (isMozilla) {
	HTMLElement.prototype.__defineGetter__("outerHTML", function () {
		var str = '';
		str += elementDetail(this);

		var inp = this.getElementsByTagName("INPUT");
		for (var i=0; i<inp.length; i++) {
			var chld = inp[i];
			str += elementDetail(chld);
		}

		var tet = this.getElementsByTagName("TEXTAREA");
		for (var i=0; i<tet.length; i++) {
			var chld = tet[i];
			str += elementDetail(chld);
		}

		str += "<" +this.tagName+ ">";
		return str;
		}
	);

	function elementDetail (element) {
		if (!element.attributes) return "";

		var attrs = element.attributes;
		var str = "<" + element.tagName;

		for (var i=0; i<attrs.length; i++) {
			if (element.tagName=="INPUT" || element.tagName=="TEXTAREA") {
				if (element.tagName=="INPUT"){
					switch (element.type.toLowerCase()) {
					case "radio" :
					case "checkbox" :
						if (element.checked){
							element.setAttribute("checked", "checked");
						}
						else{
							element.removeAttribute("checked");
						}
						break;
					case "hidden" :
					case "text" :
					case "submit" :
					case "file" :
					case "button" :
					case "reset" :
					default : 
						element.setAttribute("value", element.value);
					}
				}
				else{
					element.innerHTML = element.value;
				}

				str += ' value="' +element.value+ '"';
			}
			else {
				str += ' ' +attrs[i].name+ '="' +attrs[i].value+ '"';
			}
		}

		str += ">";

		return str;
	}
}

function cleanSpace(span){
	var tabBox = $("tabBox").innerHTML + span;

	while (tabBox.indexOf("&nbsp;&nbsp;") > -1){
		tabBox = tabBox.replace("&nbsp;&nbsp;", "&nbsp;");
	}

	return tabBox;
}

function initTab(obj) {
	if (obj.getAttribute("userData") == "0")
	{
		return false;
	}

	if(document.all){
		obj.attachEvent("onmouseover", overTab);
		obj.attachEvent("onmouseout", outTab);
	}
	else{
		obj.addEventListener("mouseover", overTab, true);
		obj.addEventListener("mouseout", outTab, true);
	}

	obj.setAttribute("userData", "0");
}

function newTab(sContent){
	var span = "&nbsp;<span id='Tab_" + tabIndex+"' onclick='javascript: switchTab("+tabIndex+");' userData='1' blank='1'>"+emptyTab+tabCloseIcon(tabIndex)+"</span>";

	$("tabBox").innerHTML = cleanSpace(span);

	initTab($("Tab_" + tabIndex));

	tabIndex++;

	switchTab(tabIndex-1, sContent);
}

function tabCloseIcon(iTabIndex) {
	return "&nbsp;<a href=\"javascript:vod();\" onclick=\"javascript:removeTab("+iTabIndex+");\" title='"+closeTab+"'>X</a>";
}

function setTabTitle(iTabIndex, sTabTitle){
	$("Tab_"+iTabIndex).innerHTML = sTabTitle + tabCloseIcon(iTabIndex);
	$("Tab_"+iTabIndex).title = sTabTitle;
}

function removeTab(tabNum){
	rTab += "Tab_" + tabNum + ",";

	var isSwitch = $("Tab_"+tabNum).getAttribute("userData");

	$("tabBox").removeChild($("Tab_"+tabNum));

	tabCache[tabNum] = "";
	tabScriptCache[tabNum] = "";

	$("tabBox").innerHTML = cleanSpace("");

	if (isSwitch == "1")
	{
		for (var i=tabNum-1;i>=0 ;i-- )
		{
			if (rTab.indexOf(",Tab_"+i+",") == -1)
			{
				if(document.all){
					$("Tab_"+i).click();
				}
				else{
					$("Tab_"+i).onclick(i);
				}
				break;
			}
		}
	}
}

function overTab(event){
	if(event == null){
		event = window.event; // For IE
	}
	var eventObj = event.srcElement? event.srcElement : event.target;  // IE use srcElement, Firefox use target
	
	eventObj.className="hover";
}

function outTab(event){
	if(event == null){
		event = window.event; // For IE
	}
	var eventObj = event.srcElement? event.srcElement : event.target;  // IE use srcElement, Firefox use target

	eventObj.className="";
}

function switchTab(eventObjIndex){
	$("contentDiv").outerHTML;
	tabCache[currentTab] = $("contentDiv").innerHTML;

	if (eventObjIndex != currentTab)
	{
		//切换标签
		if (rTab.indexOf(",Tab_"+currentTab+",") == -1){
			var tabObj = $("Tab_"+currentTab);
			tabObj.className = "";
			initTab(tabObj);
		}

		tabObj = $("Tab_"+eventObjIndex);
		tabObj.className = "abg";
		if(document.all){
			tabObj.detachEvent("onmouseover", overTab);
			tabObj.detachEvent("onmouseout", outTab);
		}
		else{
			tabObj.removeEventListener("mouseover", overTab, true);
			tabObj.removeEventListener("mouseout", outTab, true);
		}

		tabObj.setAttribute("userData", "1");
		tabObj.setAttribute("blank", "0");
		currentTab = eventObjIndex;
	}

	var argc = switchTab.arguments.length;
	var argv = switchTab.arguments;
	if (argc > 1)
	{
		var sContent = argv[1];

		if (typeof(sContent) == "undefined"){
			//新建标签
			returnMessage = "<div style='height: 300px;'></div>";
			sContent = "";

			tabObj.setAttribute("blank", "1");
			isExeJS = false;
		}
		else{
			var regexp1 = /<script(.|\n)*?>(.|\n|\r\n)*?<\/script>/ig;
			var returnMessage = sContent.replace(regexp1, "");
			sContent = sContent.match(regexp1);
			isExeJS = true;
		}

		tabCache[eventObjIndex] = returnMessage;
	}
	else{
		//切换标签
		var returnMessage = tabCache[eventObjIndex];
		var sContent = tabScriptCache[eventObjIndex];
		isExeJS = false;
	}

	$("contentDiv").innerHTML = returnMessage;
	
	executeScript(sContent, isExeJS);
}

function openTab(sContent){
	if (openNewTab)
	{
		for (var i = 0;i<tabIndex ;i++ )
		{
			if (rTab.indexOf(",Tab_"+i+",") == -1 && $("Tab_"+i).getAttribute("blank") == "1")
			{
				switchTab(i, sContent);
				return false;
			}
		}

		newTab(sContent);
	}
	else{
		if (currentTab != 0)
		{
			switchTab(currentTab, sContent);
		}
		else{
			for (var i = 1;i<tabIndex ;i++ )
			{
				if (rTab.indexOf(",Tab_"+i+",") == -1)
				{
					switchTab(i, sContent);
					return false;
				}
			}

			newTab(sContent);
		}
	}
}