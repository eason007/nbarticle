<!--#Include File="../conn.asp" -->
<!--#Include File="comm/inc.asp" -->
<%
'====================================================================
'= Team Elite - Elite Article System
'= Copyright (c) 2005 - 2007 Eason Chan All Rights Reserved.
'=-------------------------------------------------------------------
'= 版权协议：
'=	GPL (The GNU GENERAL PUBLIC LICENSE Version 2, June 1991)
'=-------------------------------------------------------------------
'= 文件名称：/Manager/Admin_Left.asp
'= 摘    要：后台-左边控制菜单文件
'=-------------------------------------------------------------------
'= 最后更新：eason007
'= 最后日期：2007-02-01
'====================================================================

Call EA_Manager.Chk_IsMaster

Dim i,j

Call EA_Pub.Close_Obj
Set EA_Pub=Nothing

%>
<script type="text/javascript">
function switchLi(iLi){
	var oImg = document.getElementById("Img"+iLi);
	if (oImg.src.indexOf("menu_title_up.gif") > -1)
	{
		oImg.src = "images/menu_title_down.gif";

		var bStat = "none";
	}
	else{
		oImg.src = "images/menu_title_up.gif";

		var bStat = "";
	}

	var aLi=getElementsByName_iefix("li","Li"+iLi);
	for (var i=0;i<aLi.length ;i++ )
	{
		aLi[i].style.display = bStat;
	}
}

function getElementsByName_iefix(tag, name) {
	var elem = document.getElementsByTagName(tag);
	var arr = new Array();
	for(i = 0,iarr = 0; i < elem.length; i++) {
		att = elem[i].getAttribute("name");
		if(att == name) {
			arr[iarr] = elem[i];
			iarr++;
		}
	}
	return arr;
}
</script>

<ul>
	<li class="title" style="border-top:1px #c5d7e2 solid; cursor:pointer;" onclick="javascript: switchLi(-1);"><img src="images/menu_title_up.gif" class="right" style="cursor:pointer;vertical-align: middle;" id="Img-1"><%=str_Thruway%></li>
	<li class="item" name="Li-1">&nbsp;&nbsp;&nbsp;<%=str_LeftMenu(0,1)%></li>
	<li class="item" name="Li-1">&nbsp;&nbsp;&nbsp;<%=str_LeftMenu(1,2)%></li>
	<li class="item" name="Li-1">&nbsp;&nbsp;&nbsp;<%=str_LeftMenu(4,4)%></li>
	<%For i=0 To Ubound(str_LeftMenu)-1%>
	<li class="title" style="cursor:pointer;" onclick="javascript: switchLi(<%=i%>);"><img src="images/menu_title_down.gif" class="right" style="cursor:pointer;vertical-align: middle;" id="Img<%=i%>"><%=str_LeftMenu(i,0)%></li>
	<%
	For j=1 To Ubound(str_LeftMenu,2)
		If IsEmpty(str_LeftMenu(i,j)) Then Exit For
	%>
	<li class="item" name="Li<%=i%>" style="display: none;">&nbsp;&nbsp;&nbsp;<%=str_LeftMenu(i,j)%></li>
	<%Next%>
	<%Next%>
</ul>