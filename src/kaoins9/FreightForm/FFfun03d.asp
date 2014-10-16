<!-- #include file = ../GlobalSet/conn.asp -->
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title>攬貨商資料</title>
	<!-- #include file = ../head_bundle.html -->
	<script type="text/javascript" src="FFfun03d.js"></script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="10" marginwidth="0" marginheight="0" onload="initialiseMenu(), CLocation();">
	<!-- #include file = ../title.htm -->
<%
    dim sql, rs
    set rs = nothing
    
    '===========接參數===========
    dim szID, szName, szStatus, szVesselLine
    szID     = request ("ID")
    szName   = request ("Name")
    szVesselLine = request ("VesselLine")
    szStatus = request ("Status")
    
    
    if szStatus = "" then
        szStatus = "Add"    'Default value
    end if
    
    '===========查詢欲修改的資料===========
    dim szIDTmp, szNameTmp, szFaxNo_1Tmp, szFaxNo_2Tmp, szContact_1Tmp, szContact_2Tmp, szPhone_1Tmp, szPhone_2Tmp, szAddressTmp
    dim szMailAddrTmp, szMailAddrCCTmp
    if szStatus = "Modify" then
        sql = "select * from FreightOwner where ID = '" + szID + "'"
        set rs = conn.execute(sql)
        
        if not rs.eof then
            szIDTmp        = rs("ID")
            szNameTmp      = rs("Name")
            szFaxNo_1Tmp   = rs("FaxNo_1")
            szFaxNo_2Tmp   = rs("FaxNo_2")
            szContact_1Tmp = rs("Contact_1")
            szContact_2Tmp = rs("Contact_2")
            szPhone_1Tmp   = rs("Phone_1")
            szPhone_2Tmp   = rs("Phone_2")
            szMailAddrTmp  = rs("MailAddr")
            szMailAddrCCTmp= rs("MailAddrCC")
            szAddressTmp   = rs("Address")
            
            '清除參數
            szID    = ""
            szName  = ""
        end if
    end if
    
%>

<form name="form" method="post" action="FOfun01b.asp" onsubmit=" javascript: return checkform();"  OnKeyDown="CheckHotKey()">
<table cellspacing=0 cellpadding=0 width="100%" border="0" align="center" >
    <tr> 
        <td > 
            <table cellspacing=0 cellpadding=0 width="100%" border=0>
                <tbody> 
                    <tr bgcolor=#3366cc > 
                        <td width=1><img height=26 src="../image/coin2ltb.gif" width=20></td>
                        <td align=middle width="100%" ><div align="center"><font color="#FFFFFF" >攬貨商資料</font></div></td>
                        <td width=1><img height=26 src="../image/coin2rtb.gif" width=20></td>
                    </tr>
                </tbody> 
            </table>
        </td>
    </tr>
    <tr> 
        <td > 
            <table cellspacing=0 cellpadding=1 width="100%" bgcolor=#000000 border=0 height="8">
                <tbody> 
                    <tr>
                        <td height="24"> 
                            <table cellspacing=0 cellpadding=0 width="100%" bgcolor=#ebebeb border=0 height="38">
                                <tbody>
    	                            <tr> 
                                        <td align=right bgcolor=#C9E0F8 height="2" width="17%">編　　　號：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2" width="31%"> 
                                            <input type="text" readonly name="ID" size="10" maxlength="10" value="<%=szIDTmp%>" onfocus=SelectText(0) OnKeyDown="ChangeFocus(1)"  onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>
                                        <td align=right bgcolor=#C9E0F8 height="2" width="20%"></td>
                                        <td align=left bgcolor=#C9E0F8 height="2" width="32%"></td>
                                    </tr>
                                    <tr>
                                        <td align=right bgcolor=#C9E0F8 height="2" ><span style="letter-spacing: 4pt">公司名稱：</span></td>
                                        <td align=left bgcolor=#C9E0F8 height="2" colspan="3"> 
                                            <input type="text" readonly name="Name" size="30" maxlength="20" value="<%=szNameTmp%>" onfocus=SelectText(1) OnKeyDown="ChangeFocus(2)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>
                                    </tr>
            
                                    <tr> 
                                        <td align=right bgcolor=#C9E0F8 height="2"><span style="letter-spacing: 4pt">傳真電話：</span></td>
                                        <td align=left bgcolor=#C9E0F8 height="2"> 
                                            <input type="text" readonly name="FaxNo_1" size="15" maxlength="10" value="<%=szFaxNo_1Tmp%>" onfocus=SelectText(2) OnKeyDown="ChangeFocus(3)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>
                                        <td align=right bgcolor=#C9E0F8 height="2">傳真電話：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2"> 
                                            <input type="text" readonly name="FaxNo_2" size="15" maxlength="10" value="<%=szFaxNo_2Tmp%>" onfocus=SelectText(3) OnKeyDown="ChangeFocus(4)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>
                                    </tr>
            
                                    <tr> 
                                        <td align=right bgcolor=#C9E0F8 height="2">聯　絡　人：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2"> 
                                            <input type="text" readonly name="Contact_1" size="15" maxlength="10" value="<%=szContact_1Tmp%>" onfocus=SelectText(4) OnKeyDown="ChangeFocus(5)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>
                                        <td align=right bgcolor=#C9E0F8 height="2">聯絡電話：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2" > 
                                            <input type="text" readonly name="Phone_1" size="15" maxlength="10" value="<%=szPhone_1Tmp%>" onfocus=SelectText(5) OnKeyDown="ChangeFocus(6)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>
                                    </tr>
            
                                    <tr> 
                                        <td align=right bgcolor=#C9E0F8 height="2">聯　絡　人：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2"> 
                                            <input type="text" readonly name="Contact_2" size="15" maxlength="10" value="<%=szContact_2Tmp%>" onfocus=SelectText(6) OnKeyDown="ChangeFocus(7)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>
                                        <td align=right bgcolor=#C9E0F8 height="2">聯絡電話：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2"> 
                                            <input type="text" readonly name="Phone_2" size="15" maxlength="10" value="<%=szPhone_2Tmp%>" onfocus=SelectText(7) OnKeyDown="ChangeFocus(8)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>
                                    </tr>
                                    
                                    <tr> 
                                        <td align=right bgcolor=#C9E0F8 height="2"><span style="letter-spacing: 4pt">E-Mail：</span></td>
                                        <td align=left bgcolor=#C9E0F8 height="2" colspan="3"> 
                                            <input type="text" readonly name="MailAddr" size="50" maxlength="50" value="<%=szMailAddrTmp%>" onfocus=SelectText(8) OnKeyDown="ChangeFocus(9)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>
                                    </tr>
                                    
                                    <tr> 
                                        <td align=right bgcolor=#C9E0F8 height="2"><span style="letter-spacing: -1pt">E-Mail副本：</span></td>
                                        <td align=left bgcolor=#C9E0F8 height="2" colspan="3"> 
                                            <input type="text" readonly name="MailAddrCC" size="50" maxlength="50" value="<%=szMailAddrCCTmp%>" onfocus=SelectText(8) OnKeyDown="ChangeFocus(9)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>
                                    </tr>
                                    
                                    <tr> 
                                        <td align=right bgcolor=#C9E0F8 height="2">地　　　址：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2" colspan="3"> 
                                            <input type="text" readonly name="Address" size="50" maxlength="50" value="<%=szAddressTmp%>" onfocus=SelectText(9) OnKeyDown="ChangeFocus(10)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
                                        </td>
                                    </tr>
    
                                </tbody> 
                            </table>
                            </td>
                        </tr>
                </tbody> 
            </table>
        </td>
    </tr>
 
    <tr> 
        <td> 
            <table cellspacing=0 cellpadding=0 width="100%" border=0>
                <tbody> 
                    <tr> 
                        <td width=1><img height=38 src="../image/box1.gif" width=20></td>
                        <td width=1><img height=38 src="../image/box2.gif" width=9></td>
                        <td valign=center align=middle width=1 background="../image/box3.gif">&nbsp; </td>
                        <td width=1><img height=38 src="../image/box4.gif" width=27></td>
                        <td valign=center align=middle width="100%" bgcolor=#3366cc> 
                            <div align="center">
                                <input type="button" style="background-color:#C9E0F8;border-style: outset; border-width: 2" name="Back" value="回上一頁" OnKeyDown="BackByKeyPress()" OnMouseUp="OnBack()" onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">
                            </div>
                        </td>
                        <td width=1><img height=38 src="../image/box5.gif" width=20></td>
                    </tr>
                </tbody>
            </table> 
        </td>
    </tr>
</table>   

<input type="hidden" name="VesselLine" value=<%=szVesselLine%>>

<%
    rs.close
    conn.close 
    
    set rs=nothing
    set conn=nothing
%>  


</form>
</body>
</html>
