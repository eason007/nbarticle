function trim(text) {
	if (text == null) return null
	return text.replace(/^(\s+)?(.*\S)(\s+)?$/, '$2');
}

function tableObject() {
	this.debug = false;
	this.total = 0;	
	this.nowPage = 1;
	this.prePage = 10;
	this.pageTargetID = "pageListDiv";
	this.tableID = "listTable";
	this.url = "";
	this.dateVal = "";
	this.noneValue = "";
	this.action = "";
	var className = 'tdlist';

	/*获得数据后的后续操作*/
	this.doResponseOther=null;
	//自动释放后续操作
	this.doResponseOtherAutoCancel=true;

	this.start = function() {
		aoTable = new ajaxObject();
		aoTable.targetArea = this;
		aoTable.doResponseMethod = this.makeTable;
		aoTable.gReturnMessage = "XML";
		aoTable.debug = this.debug;

		if(this.dateVal != "") {
			document.getElementById("divLoading").style.display = "";

			aoTable.URL = this.url;
			aoTable.dateVal=this.dateVal+"&nowPage="+this.nowPage+"&prePage="+this.prePage;
			aoTable.ajaxPostDate(aoTable);
		}
		else {
			document.getElementById("divLoading").style.display = "";

			if(this.url.indexOf("?")>-1){
				aoTable.URL = this.url+"&nowPage="+this.nowPage+"&prePage="+this.prePage;
			}else{
				aoTable.URL = this.url+"?nowPage="+this.nowPage+"&prePage="+this.prePage;
			}
			aoTable.ajaxGetDate(aoTable);
		}
	}

	this.setURL = function (sURL){
		if (this.url != sURL)
		{
			this.nowPage=1;
		}

		this.url = sURL;
	}

	this.makeTable = function(mtObj) {
		var tableXmldoc = mtObj.returnMessage;
		var obj = mtObj.targetArea;
		var	table = document.getElementById(obj.tableID);
		var elementType = "innerText";

		if (mtObj.returnMessageText == "505")
		{
			alertMessage(noExtentOfAuthority);
			removeTab(currentTab);
			return false;
		}

		document.getElementById("divLoading").style.display = "";

		/*删除旧表格*/
		for(ti=0;ti<100;ti++) {
			if(document.getElementById(obj.tableID+"row"+ti)) {
				var row = document.getElementById(obj.tableID+"row"+ti);
				var tbl = row.parentNode;
				tbl.removeChild(row);
			}
			else {
				break;
			}
		}

		if(tableXmldoc.getElementsByTagName("list").length>0) {
			var listContents = tableXmldoc.getElementsByTagName("list")[0];
			if(listContents) {
				total = listContents.getAttribute("total");
				if(total) {	
					obj.total = total;
					if(table.rows) {
						nowRows = table.rows.length;
					}
					else {
						nowRows = 0;
					}
					if(total < 1 && table) {
						if(obj.noneValue != "") {
							var row = table.insertRow(nowRows);
							hr = row.insertCell(0);
							className = 'nonetdlist';
							row.setAttribute("id", obj.tableID+"row0");
							hr.setAttribute("align", "center");
							hr.setAttribute("colSpan", "100");
							hr.setAttribute("className", className);
							hr.setAttribute("class", className);
							hr.innerHTML = trim(obj.noneValue);	
						}
					}
				}	
				
				if(listContents.getElementsByTagName("item").length > 0 && table) {
					if(table.rows) {
						nowRows = table.rows.length;
					}
					else {
						nowRows = 0;
					}

					contents = listContents.getElementsByTagName("item");

					for(var i=0; i<contents.length;i++) {
						var s = contents[i];

						var row = table.insertRow(nowRows+i);
						row.setAttribute("id", obj.tableID+"row"+i);
						
						var cName = s.getAttribute("className");
						if (!cName) {
							row.setAttribute("className", className);
							row.setAttribute("class", className);
						}

						if(s.getAttribute("ID")) {
							var ID = s.getAttribute("ID");
						}
						else {
							var ID = 0;
						}

						for(var j=s.childNodes.length-1;j>=0;j--) {
							if(s.childNodes[j].nodeType == 1) {
								var rowName = s.childNodes[j].tagName.toLowerCase();

								if(rowName.substr(0,1) != "_") {
									xmlValue = s.childNodes[j].firstChild.nodeValue;

									hr = row.insertCell(0);
									hr.setAttribute("className", className);
									hr.setAttribute("class", className);

									var rowWidth = s.childNodes[j].getAttribute("rowWidth");
									if(rowWidth) {
										hr.setAttribute("width", trim(rowWidth));
									}
									
									switch(rowName) {
										case "checkbox":
											hr.innerHTML = trim("<input type=\"checkBox\" name=\"checkbox\" value=\""+xmlValue+"\"/>");
											break;

										case "orderid":
											hr.innerHTML = trim("<input type=\"text\" name=\"place\" value=\""+xmlValue+"\" size=\"3\"/>");
											break;	

										case "action":
											hr.innerHTML = trim(replaceAll(obj.action, "$ID", ID));
											break;

										case "isclose":
											if(xmlValue == "Y") {
												hr.innerHTML = "<font color=red>已关闭</font>";
										    }
											else {
										        hr.innerHTML = "使用中";
										    }
											break;

										case "pageturn":
											eval('var pageTurnStr=obj.'+rowName+';');
											if(pageTurnStr==undefined){
												pageTurnStr=null;
												hr.innerHTML = trim(xmlValue);
												break;
											}
											var hrTemp=replaceAll(pageTurnStr, "$ID", ID);
											hrTemp=replaceAll(hrTemp, "$Value", xmlValue);
											hr.innerHTML = hrTemp;
											hrTemp=null;
											pageTurnStr=null;
											break;

										default :
											hr.innerHTML = trim(xmlValue);
											break;
									}
								}
							}
						}
					}
				}
			}
		}	
		

		if(obj.pageTargetID) {
			var total	= obj.total ? obj.total:0;
			if(document.getElementById(obj.pageTargetID)) {
				if(total > 0) {
					document.getElementById(obj.pageTargetID).innerHTML= obj.page();
					document.getElementById(obj.pageTargetID).style.display = "";
				}
				else {
					document.getElementById(obj.pageTargetID).innerHTML= "";
				}
			}
		}

		if(document.form1!=undefined && document.form1.selectAll!=undefined && document.form1.selectAll.checked){
			form1.selectAll.checked=false;
		}

		FormEdit.formName = "form1";
		FormEdit.editForm(tableXmldoc,FormEdit);
		
		if(obj.doResponseOther){
			eval(obj.doResponseOther);
			if(obj.doResponseOtherAutoCancel){
				obj.doResponseOther=null;
			}
		}

		document.getElementById("divLoading").style.display = "none";
	}


	this.page = function() {
		var total	= this.total ? this.total:0;
		
		if(total > 0) {
    		var nowPage	= this.nowPage ? this.nowPage:1;
    		var prePage	= this.prePage ? this.prePage:10;
    		var targetID = this.targetID;
    		var pageNum = (total/prePage)>parseInt(total/prePage)?(parseInt(total/prePage)+1):(parseInt(total/prePage));
    		
    		var pageContent = "";

    		pageContent += "&nbsp;&nbsp;&nbsp;";
    		pageContent += "<span id='pageOrange'>"+total+"</span>&nbsp;条记录&nbsp;&nbsp;<span id='pageOrange' class='orange'>"+nowPage+"</span>/"+pageNum+"页&nbsp;&nbsp;";
    		pageContent += "去第<input type='text' size='4' style=\"width:25px;height:17px;\" value='"+nowPage+"' onChange='javascript:checkPageNum(this, "+pageNum+", "+nowPage+", objTable);' /><input type=button value='页' style=\"height:17px;\" />&nbsp;&nbsp;";
    		if(nowPage > 1) {
    			pageContent += "<a href=\"#\" onclick=\"javascript:pageGo(1);\">首页</a>&nbsp;&nbsp;<a href=\"#\" onclick=\"javascript:pageGo("+(nowPage-1)+");\">上一页</a>";
    		}

    		pageContent += "&nbsp;&nbsp;";
    		if( pageNum > nowPage >= 1) {
    			pageContent += "<a href=\"#\" onclick=\"pageGo("+(nowPage+1)+");\">下一页</a>&nbsp;&nbsp;<a href=\"#\" onclick=\"javascript:pageGo("+pageNum+");\">尾页</a>&nbsp;&nbsp;";
    		}
    		
    		pageContent += "";
    
    		return pageContent;
		}
	} 
}

function pageGo(nowPage) {
	document.getElementById("divLoading").style.display = "";

	objTable.nowPage = nowPage;
	objTable.start();
}

function checkPageNum(now, pageNum, nowPage) {
	if(now.value > pageNum) {
		alert('输入的页码不能大于'+pageNum);
		now.value = nowPage;
		return false;
	}
	else {
		pageGo(parseInt(now.value));
		return false;
	}
}

function replaceAll(strOrg,strFind,strReplace){
	var index = 0;
	while(strOrg.indexOf(strFind,index) != -1){
		strOrg = strOrg.replace(strFind,strReplace);
		index = strOrg.indexOf(strFind,index);
	}
	return strOrg
}
