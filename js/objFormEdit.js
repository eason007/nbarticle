String.prototype.trim = function() {
	var result=this.replace(/(^\s*)/g, "");
	result=result.replace(/(\s*$)/g, "");
	return result;
}

function formEditObject() {
	this.debug = false;
	this.formName = "";
	this.postUrl = "";
	this.dateVal = "";

	//提交后的处理方法
	this.doResponseMethod = null;
	//其他表单数据填充操作
	this.doResponseMethodOther = '';

	/*是否在在线编辑器*/
	this.fckName = "";
	
	this.start = function(){
	    if(this.dateVal != "" && this.postUrl!="") {
			aoForm = new ajaxObject();
			aoForm.URL = this.postUrl;
			aoForm.dateVal = this.dateVal;
			aoForm.targetArea = this;
			aoForm.doResponseMethod = this.getReturnValue;
			aoForm.gReturnMessage = "XML";
			aoForm.ajaxPostDate(aoForm);
		}
		else {
			return false;
		}
	}

	this.getReturnValue = function(obj) {
		if(obj.targetArea.debug) {
	        alert(obj.returnMessageText);
	    }
		if (obj.returnMessageText == "505")
		{
			alertMessage(noExtentOfAuthority);
			removeTab(currentTab);
			return false;
		}
		if(!obj.targetArea.doResponseMethod) {
	        obj.targetArea.editForm(obj.returnMessage, obj.targetArea);
	    }
		else {
	        obj.targetArea.doResponseMethod(obj.returnMessage);
	    }
		if(obj.targetArea.doResponseMethodOther!=''){
			eval(obj.targetArea.doResponseMethodOther);
		}

		document.getElementById("divLoading").style.display = "none";
	}
	
	this.editForm = function(infoXmldoc, formObj) {
	    /*处理XML*/
	    if(infoXmldoc.getElementsByTagName("elements").length > 0) {
			var rowNames = infoXmldoc.getElementsByTagName("elements");
			for(var ri=rowNames[0].childNodes.length-1; ri>=0; ri--) {
				if(rowNames[0].childNodes[ri].nodeType == 1) {
					elementValue = rowNames[0].childNodes[ri].firstChild.nodeValue;
					elementID = rowNames[0].childNodes[ri].getAttribute("elementID");
					elementType = rowNames[0].childNodes[ri].getAttribute("elementelementType");
					if(document.getElementById(elementID) && elementValue != "") {
						if (document.getElementById(elementID).type == "button" || document.getElementById(elementID).type == "hidden")
						{
							document.getElementById(elementID).value = elementValue;
						}
						else{
							document.getElementById(elementID).innerHTML = elementValue;
						}
					}
				}
				if(rowNames[0].childNodes[ri].nodeType == 3) {
				}
			}
		}

        if(infoXmldoc.getElementsByTagName("info").length>0 && infoXmldoc.getElementsByTagName("info")[0].childNodes.length > 0) {
			var rowNames = infoXmldoc.getElementsByTagName("info")[0];
			for(var ri=rowNames.childNodes.length-1; ri>=0; ri--) {	
				if(rowNames.childNodes[ri].nodeType == 1) {
					xmlValue = rowNames.childNodes[ri].firstChild.nodeValue;
					xmlName = rowNames.childNodes[ri].tagName;
					if(xmlName.toLowerCase().indexOf("_password") == -1) {
						formObj.doFormValue(xmlName, xmlValue, formObj);
					}
				}
			}  
        }

		if(this.fckName != "") {
			oFCKeditor.BasePath	= "/html/FCKeditor/";
			oFCKeditor.ReplaceTextarea() ;
			__FCKeditorNS = null;
	    }
		
	    return false;
	}

	this.doFormValue = function(targetID, value, formObj) {
		//value = decodeURIComponent(value);//相对于objForm编码的解码
	    vForm = document.getElementById(formObj.formName);

	    var elementsCount = vForm.elements.length;
		var elements = vForm.elements;
		var postStr = "";

		for (var i = 0; i < elementsCount;i++) {
			var element = elements[i];
			if(element.name == targetID ) {
    	        switch (element.type.toLowerCase()) { 
    				case "hidden" :
   						element.setAttribute('value',value);
    					break;
    				case "radio" :
    					if (element.value == value) { 
    						element.setAttribute('checked',"true");
    					}
						else {
    					    element.removeAttribute("checked");
    					}
    					break;
    				case "checkbox" : 
    				    valueArray = value.split(",");
    				    element.checked = false;
    				    for(var ij=0; ij<valueArray.length; ij++) {
        					if (element.value == valueArray[ij]) { 
        						element.setAttribute('checked',"true");
        					}
    				    }
    					break;
    				case "select-one" : 
    				    if(value.toLowerCase().indexOf("(build-select)") != -1){
							var tmp = element.length;
							for (var ii=0;ii<=tmp ;ii++ )
							{
								element.remove(0);
							}

        				    var items=new Array();
							var str=value.trim();
							
							if (str.indexOf(' ')!=-1) {
								items=str.split(" ");
								
								var current=items[0].split(",");
								if(current[1]!=null && current[1]!=''){
									current=current[1];
								}
								else{
									current=-1;
								}
								
								var c=0;
								for(var ii=1;ii<items.length;ii++){
									var a=items[ii].split(",");
									if(a[0]!=null && a[0]!='' && a[1]!=null && a[1]!=''){
										var newOpt=new Option(a[0],a[1]);
										element.options[element.options.length]=newOpt;
										element.options[element.options.length-1].selected=(a[1]==current)?true:false;
										if(a[1]==current) c++;
									}
								}
								if(c==0){
									if(element.options[0]!=null){
										element.options[0].selected=true;
									}
								}
							}
						}
						else {
        				    for(var is=0; is<element.options.length; is++) {
        				        opt = element.options[is];
        				        evalue = opt.value;
        						if (!evalue && !('value' in opt)) {
        						    evalue = opt.text;
        						} 
        				        if(evalue == value) {
        				            opt.selected=true;
        				        }
								else {
        				            opt.selected=false;
        				        }
        				    }
    				    }
        				break;
    				case "select-multiple" :
    				    valueArray = value.split(",");
    					for (var j = 0; j < element.options.length; j++) {
                            var opt = element.options[j];
                            opt.selected = false; 
                            var optValue = opt.value;

                            if (!optValue && !('value' in opt))  {
                                optValue = opt.text;
                            }

                            for(var ijs=0; ijs<valueArray.length; ijs++) {
								if(optValue == valueArray[ijs]) {
									opt.selected = true;
								}
							} 
    					}
    					break;
    				case "textarea" :
						if (element.id.substr(0, 4) == "FCK_") {
   							this.fckName = element.name;
   						}
						else {
							value = value.replace(/&quot;/g, '"');
							value = value.replace(/&#039;/g, "'");
							value = value.replace(/&#92;/g, "\\");
							value = value.replace(/&lt;/g, "<");
							value = value.replace(/&gt;/g, ">");
							value = value.replace(/&#123;/g, "{");
							value = value.replace(/&#125;/g, "}");
							element.value = value;
						}
    					break;
    				case "text" :
    					value = value.replace(/&quot;/g, '"'); 
    					value = value.replace(/&#039;/g, "'");
    					value = value.replace(/&#92;/g, "\\");
    					value = value.replace(/&lt;/g, "<");
    					value = value.replace(/&gt;/g, ">");
    					value = value.replace(/&#123;/g, "{");
    					value = value.replace(/&#125;/g, "}");
    					element.setAttribute('value',value);
    					break;
					case "submit" :
					case "file" :
					case "button" :
					case "reset" :
						element.disabled = false;
						break;
    				default : 
    					element.setAttribute('value',value);
    					break;
    			}
			}
	    }	    
	}
}