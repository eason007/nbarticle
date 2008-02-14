
var sPop = null;
var postSubmited = false;

var userAgent = navigator.userAgent.toLowerCase();
var is_opera = (userAgent.indexOf('opera') != -1);
var is_saf = ((userAgent.indexOf('applewebkit') != -1) || (navigator.vendor == 'Apple Computer, Inc.'));
var is_webtv = (userAgent.indexOf('webtv') != -1);
var is_ie = ((userAgent.indexOf('msie') != -1) && (!is_opera) && (!is_saf) && (!is_webtv));
var is_ie4 = ((is_ie) && (userAgent.indexOf('msie 4.') != -1));
var is_moz = ((navigator.product == 'Gecko') && (!is_saf));
var is_kon = (userAgent.indexOf('konqueror') != -1);
var is_ns = ((userAgent.indexOf('compatible') == -1) && (userAgent.indexOf('mozilla') != -1) && (!is_opera) && (!is_webtv) && (!is_saf));
var is_ns4 = ((is_ns) && (parseInt(navigator.appVersion) == 4));
var is_mac = (userAgent.indexOf('mac') != -1);

document.write("<style type='text/css' id='defaultPopStyle'>");
document.write(".cPopText { font-family: Tahoma, Verdana; background-color: #FFFFCC; border: 1px #000000 solid; font-size: 12px; line-height: 18px; padding: 2px 4px 2px 4px; visibility: hidden; filter: Alpha(Opacity=80)}");
document.write("</style>");
document.write("<div id='popLayer' style='position:absolute;z-index:1000' class='cPopText'></div>");

function showPopupText(event) {	
	if(event.srcElement) o = event.srcElement; else o = event.target;
	if (!o) return;
	MouseX = event.clientX;
	MouseY = event.clientY;
	if(o.alt != null && o.alt!="") { o.pop = o.alt;o.alt = "" }
	if(o.title != null && o.title != ""){ o.pop = o.title;o.title = "" }
	if(o.pop != sPop) {
		sPop = o.pop;
		if(sPop == null || sPop == "") {
			document.getElementById("popLayer").style.visibility = "hidden";
		} else {
			if(o.dyclass != null) popStyle = o.dyclass; else popStyle = "cPopText";
			document.getElementById("popLayer").style.visibility = "visible";
			showIt();
		}
	}
}

function showIt() {
	document.getElementById("popLayer").className = popStyle;
	document.getElementById("popLayer").innerHTML = sPop.replace(/<(.*)>/g,"&lt;$1&gt;").replace(/\n/g,"<br>");;
	popWidth = document.getElementById("popLayer").clientWidth;
	popHeight = document.getElementById("popLayer").clientHeight;
	if(MouseX + 12 + popWidth > document.body.clientWidth) popLeftAdjust = -popWidth - 24; else popLeftAdjust = 0;
	if(MouseY + 12 + popHeight > document.body.clientHeight) popTopAdjust = -popHeight - 24; else popTopAdjust = 0;
	document.getElementById("popLayer").style.left = MouseX + 12 + document.documentElement.scrollLeft + popLeftAdjust + "px";
	document.getElementById("popLayer").style.top = MouseY + 12 + document.documentElement.scrollTop + popTopAdjust + "px";
}

if(!document.onmouseover) {
	document.onmouseover = function(e) {
		if(!e) showPopupText(window.event); else showPopupText(e);
	};
}