<!-- #include file = ../GlobalSet/conn.asp -->
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title>攬貨報告書列印/傳真</title>
	<!-- #include file = ../head_bundle.html -->
	<script type="text/javascript" src="FSfun01b.js"></script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="10" marginwidth="0" marginheight="0" onload="initialiseMenu(), CLocation();">
	<!-- #include file = ../title.htm -->
<%
    dim sql, rs
    set rs = nothing
    
    '===========接參數===========
    dim szVesselListID
    
    szVesselListID = request ("VesselListID")            '航次
    dim szVesselNo, szDate, szVesselName, szID, szOwner, szCheckInID, szLineName, szVesselLine
    'sql = "select VesselList.*, VesselLine.Name from VesselList, VesselLine where VesselList.ID = '" + szVesselListID + "'"
    'sql = sql + " and VesselLine.ID = VesselList.VesselLine"
    sql = "select VesselList.* from VesselList where VesselList.ID = '" + szVesselListID + "'"
    
    set rs = conn.execute(sql)
    
    if not rs.eof then
        szID         = rs("ID")
        szVesselName = rs("VesselName")
        szVesselNo   = rs("VesselNo")
        szOwner      = rs("Owner")
        szCheckInID  = rs("CheckInID")
        szDate       = rs("Date")
        szVesselLine = rs("VesselLine")
        
        sql = "select VesselLine.Name from VesselLine where VesselLine.ID = '" + szVesselLine + "'"
    	
    	set rs1 = conn.execute(sql)
    	
    	if not rs1.eof then
        	szLineName = rs1("Name")
        	rs1.close
        end if
        
        set rs1 = nothing
    end if

%>

<form name="form" method="post" action="FFfun02c.asp" onsubmit=" javascript: return checkform();"  OnKeyDown="CheckHotKey()">
<table cellspacing=0 cellpadding=0 width="100%" border="0" align="center" >
<tr> 
    <td > 
        <table cellspacing=0 cellpadding=0 width="100%" border=0>
            <tbody> 
                <tr bgcolor=#3366cc > 
                    <td width=1><img height=26 src="../image/coin2ltb.gif" width=20></td>
                    <td align=middle width="100%" ><div align="center"><font color="#FFFFFF" >攬貨報告書列印/傳真</font></div></td>
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
                                <!--
                                <tr> 
        	                        <td align=right bgcolor=#C9E0F8 height="2" width="30%">編　　號：</td>
        	                        <td align=left bgcolor=#C9E0F8 height="2" width="70%"><%=szVesselListID%></td>                                  
                                </tr>
                                -->                               
                                <tr> 
        	                        <td align=right bgcolor=#C9E0F8 height="2" width="30%">船　　名：</td>
        	                        <td align=left bgcolor=#C9E0F8 height="2" width="70%"><%=szVesselName%></td>                                  
                                </tr>
                                <tr> 
        	                        <td align=right bgcolor=#C9E0F8 height="2"><span style="letter-spacing: 7pt">船公司</span>：</td>
        	                        <td align=left bgcolor=#C9E0F8 height="2"><%=szOwner%></td>                                  
                                </tr>
                                <tr> 
        	                        <td align=right bgcolor=#C9E0F8 height="2">航　　次：</td>
        	                        <td align=left bgcolor=#C9E0F8 height="2"><%=szVesselNo%></td>                                  
                                </tr>  
                                <tr> 
        	                        <td align=right bgcolor=#C9E0F8 height="2">船　　掛：</td>
        	                        <td align=left bgcolor=#C9E0F8 height="2"><%=szCheckInID%></td>                                  
                                </tr> 
                                <tr> 
        	                        <td align=right bgcolor=#C9E0F8 height="2"><span style="letter-spacing: 7pt">結關日</span>：</td>
        	                        <td align=left bgcolor=#C9E0F8 height="2"><%=szDate%></td>                                  
                                </tr>
                                <tr> 
        	                        <td align=right bgcolor=#C9E0F8 height="2">航　　線：</td>
        	                        <td align=left bgcolor=#C9E0F8 height="2"><%=szLineName%></td>                                  
                                </tr> 
                                <tr> 
        	                        <td align=right bgcolor=#C9E0F8 height="2">公司名稱：</td>
        	                        <td align=left bgcolor=#C9E0F8 height="2">
                                        <select size="1" name="CompanyID" OnKeyDown="ChangeFocusByPressedKey(7,2,2)">             
                                        <%   
                                            set rs1 = nothing                                                
                                            
                                            '查詢公司資料
                                            sql = "select ID, ChineseName from CompanyInfo"
                                            
                                            set rs1 = conn.execute(sql)
                                        
                                            while not rs1.eof
                                                response.write "<option value=" & rs1("ID") & ">" & rs1("ID") & "-" & rs1("ChineseName") & "</option>"                                                    
                                        	    rs1.movenext
                                            wend
                                            
                                            rs1.close
                                            set rs1 = nothing
                                        %> 	    
                                        </select>
                                    </td>                                  
                                </tr>                                         
        	                                              
                                <tr> 
        	                        <td align=right bgcolor=#C9E0F8 height="2">
        	                            <input checked type="radio" name="SelectType" value="1" onselect=AddReportTypeItem() onfocus=OnSelectType0() style="border-style: solid; border-color: #C9E0F8" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">單　號：
        	                        </td>
        	                        <td align=left bgcolor=#C9E0F8 height="2"> 
        	                            <input type="text" name="FormID" size="40" value="" onfocus=SelectText(2) OnKeyDown="OnFormID()" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
        	                        </td>              
                                </tr> 
                                
                                <tr> 
        	                        <td align=right bgcolor=#C9E0F8 height="2">
        	                            <input type="radio" name="SelectType" value="2" onselect=RemoveReportTypeItem() onfocus=OnSelectType1() style="border-style: solid; border-color: #C9E0F8" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">攬貨商：
        	                        </td>
        	                        <td align=left bgcolor=#C9E0F8 height="2">
                                        <input type="text" name="FreightOwnerID" size="40" value="" onfocus=SelectText(4) OnKeyDown="OnFreightOwnerID()" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">                                        
                                    </td>                                  
                                </tr>  
                                
                                <tr> 
        	                        <td align=right bgcolor=#C9E0F8 height="2">報表格式：</td>
        	                        <td align=left bgcolor=#C9E0F8 height="2">
                                        <select size="1" name="ReportType" OnKeyDown="ChangeFocusByPressedKey(4,6,6)">             
                                            <option value="0">總表格式 (全部)</option>
                                            <option value="1">總表格式 (僅核對過的)</option>
                                            <option value="2">詳細尺寸資料</option>
                                        </select>
                                    </td>                                  
                                </tr>                                
                                <!--
                                <tr> 
        	                        <td align=right bgcolor=#C9E0F8 height="2"><span style="letter-spacing: 7pt">貨櫃場</span>：</td>
        	                        <td align=left bgcolor=#C9E0F8 height="2"> 
        	                            <input type="text" name="Container" size="10" value="A.C.T" onfocus=SelectText(6) OnKeyDown="ChangeFocusByPressedKey(5,7,7)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)">
        	                        </td>              
                                </tr>
                                -->
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
                            <input type="button" style="background-color:#C9E0F8;border-style: outset; border-width: 2" value="查詢資料" OnKeyDown="SearchByKeyPress()" OnMouseUp="Search()" onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">
                            <!--<input type="button" style="border-style: outset; border-width: 2" value="傳真" OnKeyDown="FaxByKeyPress()" OnMouseUp="Fax()">
                            <input type="button" style="border-style: outset; border-width: 2" value="Mail" OnKeyDown="MailByKeyPress()" OnMouseUp="Mail()">
                            <!--<input type="button" value="重置" OnKeyDown="ResetByKeyPress()" OnMouseUp="Reset()">-->
                            <input type="button" style="background-color:#C9E0F8;border-style: outset; border-width: 2" value="回上一頁" name="ReSearch" OnKeyDown="ReSearchByKeyPress()" OnMouseUp="OnReSearch()" onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">
                        </div>
                    </td>
                    <td width=1><img height=38 src="../image/box5.gif" width=20></td>
                </tr>
            </tbody>
        </table> 
    </td>
</tr>

<input type="hidden" name="VesselListID" value=<%=szVesselListID%>>  <!--航次-->
<input type="hidden" name="VesselLine" value=<%=szVesselLine%>>  <!--航次-->


<%
    rs.close
    conn.close
    
    set rs=nothing
    set conn=nothing
%> 
</table>   
</form>
</body>
</html>
