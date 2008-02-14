function MD(d){
	//取到本月的最大日期
	for(var i=31;i>=28;i--){
		var tempDate = new Date(d.getFullYear(),d.getMonth(),i)
		if(tempDate.getFullYear()==d.getFullYear()&&d.getMonth()==tempDate.getMonth())return i;
	}
}

function WC(d,tobj){
//写日历内容
	var mstr="一月,二月,三月,四月,五月,六月,七月,八月,九月,十月,十一月,十二月";
	var ccd=new Date();//当日日期
	if(d==""||typeof(d)=="undefined")d=new Date();
	//日历表头
	var ss="<table cellpadding=1 bgColor='#021853'><tr><td bgcolor='#ECF2FC'>";
	ss+="<table border=0 cellspacing=0 cellpadding=0 width='' bgColor='#ECF2FC'>";
	ss+="<tr height=22 bgColor='#0A246A'>";
	ss+="<td colspan=2>"
	//item1
	//Year List Mode
	ss+="<select style='heidth:20;width=70px'"
	ss+=" onChange=\"MC(this.value+'-"+(d.getMonth()+1)+"-"+d.getDate()+"','"+tobj+"')\">";
	var yy=parseInt(d.getFullYear());
	for(var j=(yy-50);j<=(yy+50);j++){
		ss+="<option value='"+j+"'";
		ss+=((j==yy)?" selected":"");
		ss+=">"+j+"年</option>";
	}
	ss+="</select>"
	ss+="</td>";
	ss+="<td colspan=2>";
	//item2
	//Month List
	ss+="<select style='height:20px;width:70px'"
	ss+=" onChange=\"MC('"+d.getFullYear()+"-'+this.value+'-"+d.getDate()+"','"+tobj+"')\">";
	var msa=mstr.split(",");
	for(var i=0;i<12;i++){
		ss+="<option value="+(i+1);
		ss+=(d.getMonth()==i)?" selected ":"";
		ss+=">"+msa[i]+"</option>";
	}
	ss+="</select>";
	//month list end
	ss+="</td>";
	ss+="<td colspan=3 onClick='CC()' style='cursor:hand;' align='center'>"
	ss+="<font color='#ffffff' style=font-size=9pt></b>关闭日历窗口</b></font></td></tr>";
	ss+="<tr>";
	ss+="<td align=center><font style='color:#0A246A;font-size:9pt'>星期日</font></td>";
	ss+="<td align=center><font style='color:#0A246A;font-size:9pt'>星期一</font></td>";
	ss+="<td align=center><font style='color:#0A246A;font-size:9pt'>星期二</font></td>";
	ss+="<td align=center><font style='color:#0A246A;font-size:9pt'>星期三</font></td>";
	ss+="<td align=center><font style='color:#0A246A;font-size:9pt'>星期四</font></td>";
	ss+="<td align=center><font style='color:#0A246A;font-size:9pt'>星期五</font></td>";
	ss+="<td align=center><font style='color:#0A246A;font-size:9pt'>星期六</font></td>";
	ss+="</tr>";
	ss+="<tr><td colspan=7 bgColor='#0A246A'></td></tr>"
	var mbd = new Date(d.getFullYear(),d.getMonth(),1)//指定日期的1号
	var bd=mbd.getDay();//月份开始的星期
	maxdate = MD(d);//指定日期当月的最大天数
	var cd=d.getDate();//当前日期
	var cr=((maxdate+bd)%7==0)?((maxdate+bd)/7):parseInt((maxdate+bd)/7+1)
	//写日历主体
	for(var i=0;i<cr;i++){
		ss+="<tr>";
		for(var j=0;j<7;j++){
			var tv=(i*7+j-bd+1);
			if((i==0&&j<bd)||tv>maxdate){ss+="<td>&nbsp;</td>"}else{
				ss+="<td align='center' valign='middle' style='cursor:hand;font-size:9pt;' ";
				ss+="onclick=RD('"+tobj+"','"+d.getFullYear()+"-"+(d.getMonth()+1)+"-"+tv+"')>";
				if(d.getFullYear()==ccd.getFullYear()&&d.getMonth()==ccd.getMonth()&&ccd.getDate()==tv){
					//当前日期的显示
					ss+="<font color=red><b>"+tv+"</b></font>";
					//ss+="</div>";	
				}else{
					ss+=((cd==tv)?("<font color='green'><b>"+tv+"</b></font>"):tv);
				}
				ss+="</td>";
			}
		}
		ss+="</tr>";
	}
	//表格结束
	ss+="<tr><td colspan=7 align='center'>&nbsp;<font style=font-size=9pt><b>今天:"+ccd.getFullYear()+"-"+(ccd.getMonth()+1)+"-"+ccd.getDate()+"</b></font></td></tr>";
	ss+="</table>";
	ss+="</td></tr></table>"

	return ss;	
}

function SD(sobj,tobj){
	var tt = document.getElementById(tobj);
	var ds=tt.value;
	var d;
	if(ds==""){d=new Date()}else{var da=ds.split("-");var d=new Date(da[0],da[1]-1,da[2])}

	if(document.getElementById('calendar')==null){
		document.body.innerHTML += "<div id='calendar' style='position:absolute;width:280px;height:160px;display:none;z-index:99'></div>";
	}else{
		document.getElementById('calendar').style.display="none";
	}

	var ttop = (screen.height - 160) / 2;
	var ttleft = (screen.width - 280) / 2;

	document.getElementById('calendar').innerHTML=WC(d,tobj);
	document.getElementById('calendar').style.left=ttleft;
	document.getElementById('calendar').style.top=ttop;
	document.getElementById('calendar').style.display="";
}

function RD(obj,d){
	document.getElementById(obj).value = d;
	CC();
}

function CC(){
	document.getElementById('calendar').style.display="none";
}

function MC(dstr,objstr){
	var da=dstr.split("-");var d=new Date(da[0],da[1]-1,da[2]);
	document.getElementById('calendar').innerHTML=WC(d,objstr);
}