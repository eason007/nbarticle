String.prototype.trim = function() {
	var result=this.replace(/(^\s*)/g, "");
	result=result.replace(/(\s*$)/g, "");
	return result;
}

function formObject() {
	/*显示内容*/
	this.debug = false;
	/*指定form*/
	this.formName = "";
	/*post地址*/
	this.postUrl = "";
	/*post内容*/
	this.dateVal = "";
	
	/*提交后的处理方法*/
	this.doResponseMethod = "";	
	this.doResponseOther = "";

	this.ajaxPost = function(returnStr,appendArg){
		$("divLoading").style.display = "";

		vForm = $(this.formName);
		if(vForm) {
		    if(!Validator.Validate(vForm, 3)) {
				$("divLoading").style.display = "none";
    			return false;
		    }
		}
		if(this.dateVal == "") {
    		var elementsCount = vForm.elements.length;
    		var elements = vForm.elements;
    		var postStr = "";
    		for (var i = 0; i < elementsCount;i++) {
    			var element = elements[i];
    			switch (element.type) {
    				case "radio" : 
    					if (element.checked) { 
    						postStr+="&"+element.name+"="+encodeURIComponent(element.value);
    					} 
    					break;
    				case "checkbox" :
    					if (element.checked) {
    						postStr+="&"+element.name+"="+encodeURIComponent(element.value);
    					}
    					else{
    						postStr+="&"+element.name+"=";
    					}
    					break;
    				case "select-one" : 
    					var value = '', opt, index = element.selectedIndex;
    					if (index >= 0) {
    						opt = element.options[index];
    						value = opt.value;
    						if (!value && !('value' in opt)) value = opt.text;
    					}
    					postStr+="&"+element.name+"="+encodeURIComponent(value);
    					break;
    				case "select-multiple" :
    					for (var j = 0; j < element.options.length; j++) {
    						var opt = element.options[j];
    						if (opt.selected) {
    							var optValue = opt.value;
    							if (!optValue && !('value' in opt)) optValue = opt.text;
    							postStr+="&"+element.name+"="+encodeURIComponent(optValue);
    						}
    					}
    					break;
    				default : 
    					if (element.type != "submit" && element.type != "button" && element.type != "reset" && typeof(element) == "object") {
    						postStr+="&"+element.name+"="+encodeURIComponent(element.value);
    					} else {
    						element.disabled = true;
    					}
    					break;
    			}
    		}	
		} else {
		    postStr = this.dateVal;
		}

		if(!returnStr) {
			if(this.postUrl && this.postUrl!="") {
				aoForm = new ajaxObject();
				aoForm.URL = this.postUrl;
				if(appendArg){
					aoForm.dateVal = postStr+appendArg;
				}else{
					aoForm.dateVal = postStr;
				}
				aoForm.targetArea = this;
				aoForm.doResponseMethod = this.getReturnValue;
				aoForm.ajaxPostDate(aoForm);
			} else {
				return postStr;
			}
			this.dateVal = "";
		} else {
			return postStr;
		}
	}
	this.getReturnValue = function(obj) {
		if(obj.targetArea.debug) {
	        alert(obj.returnMessageText);
	    }
        if(obj.targetArea.doResponseMethod) {
            obj.targetArea.doResponseMethod(obj.returnMessage);
        }
        if(obj.targetArea.doResponseOther) {
            eval(obj.targetArea.doResponseOther);
        }

		$("divLoading").style.display = "none";
	}
	this.resetForm = function(){
	   vForm = $(this.formName);
	   if(vForm) {
			vForm.reset();
			var elementsCount = vForm.elements.length;
			var elements = vForm.elements;
			for (var i = 0; i < elementsCount;i++) {
				var element = elements[i];
				switch (element.type) {
    				case "checkbox" :
						element.checked = false;
    				break;
    				default : 
    					if (element.type == "submit" || element.type == "button" || element.type == "reset") {
    						element.disabled = false;
    					}
						switch(element.getAttribute('default')){
							case 'undisabled':
								element.disabled = false;
								break;
						}
    				break;
    			}
    		}
		}
	}
}