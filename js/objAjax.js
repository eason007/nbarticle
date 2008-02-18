function ajaxObject() {
	/*
		gReturnMessage：返回内容类型，XML|TEXT
		targetArea：目标区域ID或对象
		URL：提交的地址
		dateVal：POST提交的提交参数
		actionMessage：提示框中的提示内容
		returnMessage：处理成功返回的内容
		doResponseMethod：处理方法
	*/
	this.objType = "POST";
	this.gReturnMessage = "TEXT";
	this.targetArea = "";
	this.URL = "";
	this.dateVal = "";
	this.actionMessage = "";
	this.returnMessage = "";
	this.doResponseMethod = "";
	this.xmlreq = false;
	this.returnMessageText ="";
	this.returnMessageXML = "";
	this.pageType = "";
	this.debug = false;
	var errCount = 0;

	this.newXMLHttpRequest = function(){
	//获取xmlhttp控件对象
		xmlreq = this.getMsObject();

		if(typeof(xmlreq)!="object"){
			try {
				xmlreq = new XMLHttpRequest();
			}
			catch (ex){
				if (this.debug == true)
				{
					alert("Create XMLHTTP Object Fail");
				}
				else{
					return false;
				}
			}
		}

		return xmlreq;
	}

	this.getMsObject = function(){
		var msxmls = ["MSXML3", "MSXML2", "Microsoft"];

		for (var i=0; i < msxmls.length; i++) {
			try {
				xmlreq = new ActiveXObject(msxmls[i] + ".XMLHTTP");
				break;
			}
			catch (ex)
			{
				if (i==(msxmls.length-1))
				{
					return false;
				}
			}
		}

		return xmlreq;
	}
	
	this.getReadyStateHandler = function (xmlreq, obj) {
		return function () {
			if (xmlreq.readyState == 4) {
				if (xmlreq.status == 200) {
					if (obj.debug == true){
						alert(xmlreq.responseText);
					}
					if(xmlreq.responseText == "-100") {
						top.location = "./";
					}
					if (obj.gReturnMessage == "TEXT") {	
						obj.returnMessage = xmlreq.responseText.toString();
						obj.returnMessageText = xmlreq.responseText;
						obj.returnMessageXML = xmlreq.responseXML;
						obj.doResponseMethod(obj);
					}
					else {
						obj.returnMessage = xmlreq.responseXML; 
						obj.returnMessageText = xmlreq.responseText;
						obj.returnMessageXML = xmlreq.responseXML;
						obj.doResponseMethod(obj);
					}
				}
				else {
					if (errCount < 4)
					{
						if(obj.objType == "GET") {
							errCount++;
							obj.ajaxGetDate(obj);
						}
						else if(obj.objType == "POST") {
							errCount++;
							obj.ajaxPostDate(obj);
						}
					}
					else{
						if (obj.debug == true){
							alert("Request Fail\nMethod:"+obj.objType+"\nUrl:"+obj.URL+"\nData:"+obj.dateVal);
						}
						
						errCount = 0;
						return false;
					}
				}
			}
		};
	}

	/*get方式取得数据*/
	this.ajaxGetDate = function(obj) {
		obj.objType = "GET";

		obj.xmlreq = obj.newXMLHttpRequest();
		obj.xmlreq.onreadystatechange = obj.getReadyStateHandler(obj.xmlreq, obj);
		obj.xmlreq.open("GET", obj.URL, true);
		obj.xmlreq.send(null);

		delete (obj.xmlreq);
	}

	/*post方式取得数据*/
	this.ajaxPostDate = function(obj) {
		obj.objType = "POST";

		obj.xmlreq = obj.newXMLHttpRequest();
		obj.xmlreq.onreadystatechange = obj.getReadyStateHandler(obj.xmlreq, obj);
		obj.xmlreq.open("POST", obj.URL, true);

		obj.xmlreq.setRequestHeader("Method", "POST " + obj.URL + " HTTP/1.1");
		obj.xmlreq.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
		obj.xmlreq.setRequestHeader("Cache-Control", "no-cache");
		obj.xmlreq.setRequestHeader("Pragma", "no-cache");

		obj.xmlreq.send(obj.dateVal);

		delete (obj.xmlreq);
	}
	
	/*内容插入页面中*/
	this.dataInsertPage = function (obj) {
	    if(obj.targetArea) {
		    var regexp1 = /<script(.|\n)*?>(.|\n|\r\n)*?<\/script>/ig;
    		var regexp2 = /<script(.|\n)*?>((.|\n|\r\n)*)?<\/script>/im;
    		
    		var message = obj.returnMessage;
    		/* draw the html first */
    		var returnMessage = message.replace(regexp1, "");

    		$(obj.targetArea).innerHTML = returnMessage;
    		var result = message.match(regexp1);
    		if (result) {
    			for (var i = 0; i < result.length; i++) {
					/* Note: do not try to write more than one <script> in your view.*/
    				/* break;  process only one script element */

    				var realScript = result[i].match(regexp2);
    				_executeScript(realScript[2]);
    			}
    		}
    		$(obj.targetArea).style.display="";
		}
		else {
		    return true;
		}
	}

	/*处理element*/
	this.replaceElement = function (mtObj) {
		var xmldoc = mtObj.returnMessage;
		if(xmldoc.getElementsByTagName("elements").length > 0) {
			var rowNames = xmldoc.getElementsByTagName("elements");
			for(var ri=rowNames[0].childNodes.length-1; ri>=0; ri--) {	
				if(rowNames[0].childNodes[ri].nodeType == 1) {
					elementValue = rowNames[0].childNodes[ri].firstChild.nodeValue;
					elementID = rowNames[0].childNodes[ri].getAttribute("elementID");
					elementType = rowNames[0].childNodes[ri].getAttribute("elementType");
					if($(elementID)) {
						if(elementType == "innerText") {
							if(navigator.appName.indexOf("Explorer") > -1){
								$(elementID).innerText = elementValue;
							} 
							else{
								$(elementID).textContent = elementValue;
							}
						} 
						else {
							if($(elementID)) {
								if($(elementID).type == "hidden" || $(elementID).type == "text"  || $(elementID).type == "textArea") {
									$(elementID).value = elementValue;
								} 
								else {
									$(elementID).innerHTML = elementValue;
								}
							}
						}
					}
				}
			}
		}
		if($("divLoading")) {
			$("divLoading").style.display="NONE";
		}
	}
}

/*Get方式取得数据并插入页面*/
function ajaxGetDateToPage(URL,targetArea,actionMessage,isDebug) {
	aoNew = new ajaxObject();

	aoNew.debug = isDebug;
	aoNew.URL = URL;
	aoNew.targetArea = targetArea;
	aoNew.doResponseMethod = aoNew.dataInsertPage;
	aoNew.actionMessage = aoNew.actionMessage;
	aoNew.ajaxGetDate(aoNew);
}

//Post方式取得数据并插入页面
function ajaxPostDateToPage(URL, dateVal, targetArea, actionMessage) {
	aoNew = new ajaxObject();

	aoNew.URL = URL;
	aoNew.dateVal = dateVal;
	aoNew.targetArea = targetArea;
	aoNew.doResponseMethod = aoNew.dataInsertPage;
	aoNew.actionMessage = aoNew.actionMessage;
	aoNew.ajaxPostDate(aoNew);
}

/*Get方式取得数据*/
function ajaxGetDate(URL, responseMethod, returnMessage) {
	aoNew = new ajaxObject();

	aoNew.URL = URL;
	aoNew.targetArea = null;
	aoNew.doResponseMethod = responseMethod;
	aoNew.gReturnMessage = returnMessage;
	aoNew.actionMessage = null;
	aoNew.ajaxGetDate(aoNew);
}

//Post方式取得数据
function ajaxPostDate(URL, dateVal, responseMethod, returnMessage) {
    aoNew = new ajaxObject();

	aoNew.URL = URL;
	aoNew.dateVal = dateVal;
	aoNew.targetArea = null;
	aoNew.doResponseMethod = responseMethod;
	aoNew.gReturnMessage = returnMessage;
	aoNew.ajaxPostDate(aoNew);
}

/*处理element*/
function ajaxReplaceElement(URL, dateVal, isDebug) {
	aoNew = new ajaxObject();

	aoNew.debug = isDebug;
	aoNew.URL = URL;
	aoNew.gReturnMessage = "XML";
	aoNew.dateVal = dateVal;
	aoNew.doResponseMethod = aoNew.replaceElement;
	aoNew.ajaxPostDate(aoNew);
}
