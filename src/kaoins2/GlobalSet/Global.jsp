<%@page contentType="text/html;charset=Big5" language="java" import="java.sql.*,java.util.*"%>
<%
int secerror=0;
String sec=(String)session.getAttribute("Security");
if (sec==null) secerror=1;
else if (!(sec.equals("1") || sec.equals("2") || sec.equals("3"))) secerror=1;
if (secerror==1) response.sendRedirect("../Logout.htm");
Class.forName("sun.jdbc.odbc.JdbcOdbcDriver");
int bpfun;
bpfun=0;
%>
<%!
public String BackLink(String Link,String Mes)
{
	return "<tr><th align=center colspan=4><a href="+Link+">"+Mes+"</a></th></tr>";
}


//將日期欄位內容正規化
public String StandardDate(String od)
{
	if(od==null) od=""; else if(Integer.parseInt(od.substring(0,4))<=1901) od=""; 
	if(od.length()>=10) od=od.substring(0,10);
	return od;
}

//將有效日期轉換成顏色
public String TranDateToColor(String vd)
{
	Calendar nc=Calendar.getInstance();
	String color="",vds=vd;
	int year=0,ivd=0,ind=0;
		
	if (vd!= null && vds.length()>=10)
        {
        	ivd=(Integer.parseInt(vds.substring(0,4))-1900)*365+Integer.parseInt(vds.substring(5,7))*30+Integer.parseInt(vds.substring(8,10));
        	ind=(nc.get(Calendar.YEAR)-1900)*365+((nc.get(Calendar.MONTH)+1)*30)+nc.get(Calendar.DAY_OF_MONTH);
        	year=ivd-ind;
		if(year<0)
			//過期->紅色
			color="FF0000";
		else if(year>365)
			//還有一年以上的有效期限
			color="00FF00";
		else
			//一年以內的有效期
			color="FF7F00";
	}
	return color;
}
//將數字轉換成處理狀況
public String TranNumToProgress(String num)
{
	String res="未知";
	
	if(num.equals("0")) res="未處理";
	else if(num.equals("1")) res="處理中";
	else if(num.equals("2")) res="已處理";
	return res;
}
%>
<style type="text/css">
A:hover {COLOR:#008080;FONT-FAMILY:新細明體;FONT-SIZE:15px;text-decoration:underline}
A	{COLOR:#6666bb;FONT-FAMILY:新細明體;FONT-SIZE:15px;text-decoration:none}
TD	{COLOR:#222222;FONT-FAMILY:新細明體;FONT-SIZE:15px}
TH	{COLOR:#222222;FONT-FAMILY:新細明體;FONT-SIZE:18px}
</style>
<center>
<SCRIPT Language="Javascript">
document.onkeypress = getKey;

function printit()
{ 
	if (window.print) 
	{
		window.print() ; 
	} 
	else 
	{
		var WebBrowser = '<OBJECT ID="WebBrowser1" WIDTH=0 HEIGHT=0 CLASSID="CLSID:8856F961-340A-11D0-A96B-00C04FD705A2"></OBJECT>';
		document.body.insertAdjacentHTML('beforeEnd', WebBrowser);
		WebBrowser1.ExecWB(6, 2);
		WebBrowser1.outerHTML = ""; 
	}
}

function getKey(keyStroke) 
{
	if (navigator.appName == "Netscape")
	{
    		eventChooser = keyStroke.which;
  	}
  	else
  	{
  		eventChooser = event.keyCode; 
  	}
  	which = String.fromCharCode(eventChooser).toLowerCase();
  	if (eventChooser == 13) ;
}

function ETab(input, len, index) 
{
	if(input.value.length >= len)
  	{
    		if (index == input.form.length) index = 0;
    		input.form[index].focus();
  	}
}

</script>
