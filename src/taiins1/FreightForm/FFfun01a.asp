<!-- #include file = ../GlobalSet/conn.asp -->
<!-- #include file = FFShareFun.asp -->
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title>進倉單資料設定</title>
	<!-- #include file = ../head_bundle.html -->
	<script type="text/javascript" src="FFfun01.js"></script>
	<script type="text/javascript" src="FFShareFun.js"></script>
	<script type="text/vbscript" src="FFShareFun.vbs"></script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="10" marginwidth="0" marginheight="0" onload="initialiseMenu(), CLocation();">
	<!-- #include file = ../title.htm -->
<%
    dim sql, rs
    set rs = nothing
    
    '===========接參數===========
    dim szVesselID, szStatus
    szVesselID      = request ("VesselID")            '編號
    szStatus        = request ("Status")
        
    '20-Nov2004: 記住要keep的form ID
    dim szPrevFoundID
    szPrevFoundID = request("PrevFoundID")
        
    if szStatus = "" then
        szStatus = "Add"    'Default value
    end if
    
        
    '===========查詢欲修改的資料===========    
    dim nStoreSum_PieceTmp, fStoreSum_WeightTmp, fStoreSum_ForestryTmp, fStoreSum_VolumeTmp
    dim szIDTmp(50), szStorehouseTmp(50), szPackageStyleIDTmp(50)
    dim nPageNoTmp(50), nBoardTmp(50), nPieceTmp(50)
    dim fLengthTmp(50), fWidthTmp(50), fHeightTmp(50), fVolumeTmp(50), fWeightTmp(50), fTotalWeightTmp(50)
    dim bIsPLTmp(50)
    dim nSNTmp(50)
  
    dim nDataCounter
    
    
    dim szVesselNoTmp, szVesselNameTmp, szVesselDateTmp, szOwnerTmp, szCheckInIDTmp
    
    
    szVesselIDTmp = szVesselID            '船碼
    
    '查詢船名, 航次, 結關日 
    sql = "select * from VesselList where ID ='" + szVesselID + "'"
    set rs = conn.execute(sql)
    if not rs.eof then
        szVesselNoTmp       = rs("VesselNo")        '航次 
        szOwnerTmp          = rs("Owner")           '船公司
        szVesselNameTmp     = rs("VesselName")      '船名
        szCheckInIDTmp      = rs("CheckInID")       '船掛
        szVesselDateTmp     = rs("Date")            '結關日            
    end if
    
    if szStatus = "Add" then
        nDataCounter = 50    
    '修改資料              
    else
        
        nStoreSum_PieceTmp     = 0          '總件數
        fStoreSum_WeightTmp    = 0          '總重量
        fStoreSum_ForestryTmp  = 0          '總才積
        fStoreSum_VolumeTmp    = 0          '總體積
        
    
        '查細項資料
        set rs = nothing
        nDataCounter = 0
        sql = "select * from FreightForm where VesselID = '" + szVesselID + "'"
        set rs = conn.execute(sql)
        
        while not rs.eof
            nSNTmp(nDataCounter)              = rs("SN")                'Primary key: Serial no
            nPageNoTmp(nDataCounter)          = rs("PageNo")            '頁次
            szIDTmp(nDataCounter)             = rs("ID")                '單號
            szStorehouseTmp(nDataCounter)     = rs("Storehouse")        '倉位
            nBoardTmp(nDataCounter)           = rs("Board")             '板數    
            bIsPLTmp(nDataCounter)            = rs("IsPL")              '堆量 ( P/L )
            nPieceTmp(nDataCounter)           = rs("Piece")             '件數
            szPackageStyleIDTmp(nDataCounter) = rs("PackageStyleID")    '包裝    
            fLengthTmp(nDataCounter)          = rs("Length")            '長
            fWidthTmp(nDataCounter)           = rs("Width")             '寬
            fHeightTmp(nDataCounter)          = rs("Height")            '高
            fVolumeTmp(nDataCounter)          = rs("Volume")            '體積
            fWeightTmp(nDataCounter)          = rs("Weight")            '單重
            fTotalWeightTmp(nDataCounter)     = rs("TotalWeight")       '總重
            
            '計算總量
            nStoreSum_PieceTmp = nStoreSum_PieceTmp + nPieceTmp(nDataCounter)
            fStoreSum_WeightTmp = fStoreSum_WeightTmp + fTotalWeightTmp(nDataCounter)
            fStoreSum_ForestryTmp = fStoreSum_ForestryTmp
            fStoreSum_VolumeTmp = fStoreSum_VolumeTmp + fVolumeTmp(nDataCounter)
            
            nDataCounter = nDataCounter + 1
            rs.movenext            
        wend
        
        '清除參數
        szID = ""
    end if
    
    '檢查資料是否超出範圍
    if nDataCounter > 50 then
        nDataCounter = 50
    end if
    
    '計算PackageStyle的總數
    dim nPackageStyleCount
    sql="select Count(ID) as TotalCount from PackageStyle"
    
    set rs=conn.execute(sql)
    if not rs.eof then
        nPackageStyleCount = rs("TotalCount")
    end if

    
%>



<form name="form" method="post" action="FFfun01b.asp" onsubmit=" javascript: return checkform();"  OnKeyDown="CheckHotKey(); CheckSaveHotKey1();SwitchToSearchFForm(); return CheckForbiddenKey();">
<table cellspacing=0 cellpadding=0 width="100%" border="0" align="center" >
<tr> 
    <td > 
        <table cellspacing=0 cellpadding=0 width="100%" border=0>
            <tbody> 
                <tr bgcolor=#3366cc > 
                    <td width=1><img height=26 src="../image/coin2ltb.gif" width=20></td>
                    <td align=middle width="30%"><div align="center"><font color="#FFFFFF" >儲存(F2)　　倉單查詢(F8)</font></div></td>
                    <td align=middle width="45%" ><div align="center"><font color="#FFFFFF" >進倉單資料設定</font></div></td>                    
                    <td align=middle width="25%"><div align="center"><font color="#FFFFFF" >計算機(=)</font></div></td>
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
                                <tr bgcolor=#C9E0F8 height="2"> 
                                    <td align="left" width="1%"></td>
                                    <td align="left" width="18%">船公司:<%=szOwnerTmp%></td>
                                    <td align="left" width="17%">船名:<%=szVesselNameTmp%></td>                                    
                                    <td align="left" width="18%">航次:<%=szVesselNoTmp%></td>
                                    <td align="left" width="21%">船掛:<%=szCheckInIDTmp%></td>
                                    <td align="left" width="25%">結關日:<%=szVesselDateTmp%></td>                                                     
                                </tr> 
                                
                        <%
                            if szStatus = "Modify" then
                        %>
                                <tr align="center" bgcolor=#C9E0F8 height="2"> 
                                    <td align="right">總件數:</td>
                                    <td align="left" ><%=nStoreSum_PieceTmp%></td>  
                                    <td align="right">總重量:</td>
                                    <td align="left" ><%=fStoreSum_WeightTmp%></td>
                                    <td align="right">總才積:</td>
                                    <td align="left" ><%=fStoreSum_ForestryTmp%></td>
                                    <td align="right"><span style="letter-spacing: 4pt">總體積:</span></td>
                                    <td align="left"><%=fStoreSum_VolumeTmp%></td>  
                                </tr>
                                <input type="hidden" name="StoreSum_Piece" value=<%=nStoreSum_PieceTmp%>>         <!--總件數-->
                                <input type="hidden" name="StoreSum_Weight" value=<%=fStoreSum_WeightTmp%>>       <!--總重量-->
                                <input type="hidden" name="StoreSum_Forestry" value=<%=fStoreSum_ForestryTmp%>>   <!--總才積-->
                                <input type="hidden" name="StoreSum_Volume" value=<%=fStoreSum_VolumeTmp%>>       <!--總體積-->
                                    
                        <%
                            end if
                        %>
                            </tbody>
                        </table>
                        <table cellspacing=0 cellpadding=0 width="100%" bgcolor=#ebebeb border=0 height="38">
                            <tbody>        
                                <tr align="center" bgcolor=#C9E0F8 height="1">
                                    <td width="3%"></td>
                                    <td width="8%"></td>
                                    <td width="8%"></td>
                                    <td width="8%"></td>
                                    <td width="6%"></td>
                                    <td width="5%"></td>
                                    <td width="8%"></td>
                                    <td width="6%"></td>
                                    <td width="8%"></td>
                                    <td width="7%"></td>
                                    <td width="8%"></td>
                                    <td width="8%"></td>
                                    <td width="8%"></td>
                                    <td width="8%"></td> 
                                    <td width="1%"></td>                      
                                </tr> 
                                
                                
                                <tr> 
                                    <td align="center" bgcolor=#C9E0F8 height="2" colspan=15><hr></td>                            
                                </tr> 
                                
                                <tr align="center" bgcolor=#C9E0F8 height="2">
                                    <td></td> 
                                    <td><font size="3">頁次</font></td>
                                    <td><font size="3">單 號</font></td>
                                    <td><font size="3">倉 位</font></td>
                                    <td><font size="3">板 數</font></td>
                                    <td><font size="3">堆 量</font></td>
                                    <td><font size="3">件 數</font></td>
                                    <td><font size="3">包 裝</font></td>
                                    <td><font size="3">長</font></td>
                                    <td><font size="3">寬</font></td>  
                                    <td><font size="3">高</font></td>
                                    <td><font size="3">體 積</font></td>
                                    <td><font size="3">單 重</font></td> 
                                    <td><font size="3">總 重</font></td>
                                    <td></td>                       
                                </tr> 
                            <%
                                dim nLineCnt, nCurIndex
                                nLineCnt = 0
                                nCurIndex = 0
                                
                                '查詢包裝
                                sql="select * from PackageStyle order by ID"
                                set rs=conn.execute(sql)
                                
                                for i = 0 to nDataCounter-1
                            %> 
                                <%
                                    nCurIndex = nLineCnt
                                %>
                                <input type="hidden" name="SN_<%=i%>" value=<%=nSNTmp(i)%>>         <!--Serial number-->
                                <tr align="center" bgcolor=#C9E0F8 height="2">
                                    <%
                                        nCurIndex = nCurIndex + 1 
                                    %>
                                    <td><input type="checkbox" name="Del_<%=i%>" value="1" style="border-color: #C9E0F8" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)"></td>
                                    <%
                                        nCurIndex = nCurIndex + 1 
                                    %>   
                                    <td><input type="text" name="PageNo_<%=i%>"      size="3" maxlength="3" value="<%=nPageNoTmp(i)%>" style="direction: rtl" onfocus="SelectText(<%=nCurIndex%>)" OnKeyDown="OnDirectionKey_LField(<%=nCurIndex%>, <%=i%>);ChangeFocus(<%=nCurIndex+1%>)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)"></td>  
                                    <%
                                        nCurIndex = nCurIndex + 1 
                                    %> 
                                    <td><input type="text" name="ID_<%=i%>"          size="4" maxlength="5" value="<%=szIDTmp(i)%>" onfocus="OnIDFocus(<%=nCurIndex%>); SelectText(<%=nCurIndex%>)" OnKeyDown="OnDirectionKey_MField(<%=nCurIndex%>, <%=i%>); ChangeFocus(<%=nCurIndex+3%>)" onblur="Repeat(<%=nCurIndex%>)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)"></td>  
                                    <%
                                        nCurIndex = nCurIndex + 1 
                                    %> 
                                    <td><input type="text" name="Storehouse_<%=i%>"  size="3" maxlength="4" value="<%=szStorehouseTmp(i)%>" onfocus=SelectText(<%=nCurIndex%>) OnKeyDown="OnDirectionKey_MField(<%=nCurIndex%>, <%=i%>); ChangeFocus(<%=nCurIndex+2%>)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)"></td>
                                    <%
                                        nCurIndex = nCurIndex + 1 
                                    %> 
                                    <td><input type="text" name="Board_<%=i%>"       size="1" maxlength="2" value="<%=nBoardTmp(i)%>" style="direction: rtl" onfocus=SelectText(<%=nCurIndex%>) OnKeyDown="OnDirectionKey_MField(<%=nCurIndex%>, <%=i%>); ChangeFocus(<%=nCurIndex+1%>)" onblur="VolumeCalculator(<%=i%>)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)"></td>
                                    <%
                                        nCurIndex = nCurIndex + 1 
                                    %>
                                    <td>
                                        <select size="1" name="IsPL_<%=i%>" onmousewheel="javascript: return false;" OnKeyDown="OnIsPL(<%=i%>, <%=nCurIndex%>, <%=nCurIndex+1%>); OnDirectionKey_MField_OnlyRL(<%=nCurIndex%>, <%=i%>)" onblur="VolumeCalculator(<%=i%>)">>
                                        <%
                                            if (bIsPLTmp(i) = 0) then 
                                                response.write "<option selected value=0>" + "" + "</option>"
                                                response.write "<option value=1>" + "是" + "</option>"                                                
                                            else
                                                response.write "<option value=0>" + "" + "</option>"
                                                response.write "<option selected value=1>" + "是" + "</option>"                                                
                                            end if     
                        	  	        %>                                	  	
                                  	    </select>
                                    </td>
                                    <%
                                        nCurIndex = nCurIndex + 1 
                                    %>
                                    <td><input type="text" name="Piece_<%=i%>" size="3" maxlength="5" value="<%=nPieceTmp(i)%>" onblur="SimpleCalculator(<%=nCurIndex%>); VolumeCalculator(<%=i%>)" onfocus="OnPieceFocus(<%=nCurIndex%>); SelectText(<%=nCurIndex%>)" OnKeyDown="OnDirectionKey_MField(<%=nCurIndex%>, <%=i%>); ChangeFocus(<%=nCurIndex+2%>)" onkeyup="CallInputBox(<%=nCurIndex%>)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)"></td>
                                    <%
                                        nCurIndex = nCurIndex + 1 
                                    %>
                                    <td>                                        
                                        <select size="1" name="PackageStyleID_<%=i%>"  onmousewheel="javascript: return false;" OnKeyDown="OnPackageStyleID(<%=nPackageStyleCount%>, <%=i%>, <%=nCurIndex%>, <%=nCurIndex+1%>); OnDirectionKey_MField_OnlyRL(<%=nCurIndex%>, <%=i%>)">             
                            	  	        <option value=0></option>
                                    	  	<% 
                                    	  	    if nPackageStyleCount <> 0 then
                                    	 		    rs.movefirst
                                    	 		end if
                                    	        
                                    	 		while not rs.eof
                                    	 		    if rs("ID") = szPackageStyleIDTmp(i) then
                                    	 		        response.write "<option selected value=" & rs("ID") & ">" & rs("Name") & "</option>" 
                                    	 		    else           		 	
                                                        response.write "<option value=" & rs("ID") & ">" & rs("Name") & "</option>" 
                                                    end if
                                    	 			rs.movenext
                                    	 		wend
                                    	 		
                                    	 	%> 	    
                                  	    </select>
                                    </td>
                                    <%
                                        nCurIndex = nCurIndex + 1 
                                    %>
                                    <td><input type="text" style="direction: rtl" name="Length_<%=i%>"          size="3" maxlength="4" value="<%=fLengthTmp(i)%>"      onfocus=SelectText(<%=nCurIndex%>) OnKeyDown="OnDirectionKey_MField(<%=nCurIndex%>, <%=i%>); ChangeFocus(<%=nCurIndex+1%>)" onblur="VolumeCalculator(<%=i%>)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)"></td>
                                    <%
                                        nCurIndex = nCurIndex + 1 
                                    %>
                                    <td><input type="text" style="direction: rtl" name="Width_<%=i%>"           size="3" maxlength="4" value="<%=fWidthTmp(i)%>"       onfocus=SelectText(<%=nCurIndex%>) OnKeyDown="OnDirectionKey_MField(<%=nCurIndex%>, <%=i%>); ChangeFocus(<%=nCurIndex+1%>)" onblur="VolumeCalculator(<%=i%>)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)"></td>
                                    <%
                                        nCurIndex = nCurIndex + 1 
                                    %>
                                    <td><input type="text" style="direction: rtl" name="Height_<%=i%>"          size="3" maxlength="4" value="<%=fHeightTmp(i)%>"      onfocus=SelectText(<%=nCurIndex%>) OnKeyDown="OnDirectionKey_MField(<%=nCurIndex%>, <%=i%>); ChangeFocus(<%=nCurIndex+2%>)" onblur="VolumeCalculator(<%=i%>)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)"></td>
                                    <%
                                        nCurIndex = nCurIndex + 1 
                                    %>
                                    <td><input type="text" style="direction: rtl" name="Volume_<%=i%>"          size="5" maxlength="6" value="<%=fVolumeTmp(i)%>"      onfocus=SelectText(<%=nCurIndex%>) OnKeyDown="OnDirectionKey_MField(<%=nCurIndex%>, <%=i%>); ChangeFocus(<%=nCurIndex+6%>)" onblur="VolumeCalculator(<%=i%>)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)"></td>
                                    <%
                                        nCurIndex = nCurIndex + 1 
                                    %>
                                    <td><input type="text" style="direction: rtl" name="Weight_<%=i%>"          size="3" value="<%=fWeightTmp(i)%>"      onfocus="SelectText(<%=nCurIndex%>);UpdateWeightHasFocus()" OnKeyDown="OnDirectionKey_MField(<%=nCurIndex%>, <%=i%>); ChangeFocus(<%=nCurIndex+1%>)" onblur="WeightCalculator(<%=i%>);UpdateWeightLoseFocus()" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)"></td>
                                    <%
                                        nCurIndex = nCurIndex + 1 
                                    %>
                                    <td><input type="text" style="direction: rtl" name="TotalWeight_<%=i%>"     size="4" value="<%=fTotalWeightTmp(i)%>" onfocus="SelectText(<%=nCurIndex%>);UpdateTotalWeightHasFocus()" onblur=UpdateTotalWeightLoseFocus() OnKeyDown="OnDirectionKey_RField(<%=nCurIndex%>, <%=i%>); ChangeFocus(<%=nCurIndex+4%>)" onblur="WeightCalculator(<%=i%>);UpdateTotalWeightLoseFocus()" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)"></td>                 
                                    <td></td>
                                </tr>  
                                
                            <%
                                    nLineCnt = nLineCnt + 15
                                next
                            %>
                                
                            </tbody> 
                        </table>
                    </td>
                </tr>
            </tbody> 
        </table>
    </td>
</tr>
 
<%
    '查核定過的單號
    dim TextString
    call ShowText(conn,TextString)	
    
%>

<script lanuage="javascript">
<%= TextString %>
</script>

<tr> 
    <td> 
        <table cellspacing=0 cellpadding=0 width="100%" border=0>
            <tbody> 
                <tr> 
                    <td width=1><img height=38 src="../image/box1.gif" width=20></td>
                    <td width=1><img height=38 src="../image/box2.gif" width=9></td>
                    <td valign=center align=middle width=1 background="../image/box3.gif">&nbsp; </td>
                    <td width=1><img height=38 src="../image/box4.gif" width=27></td>
                    <td valign="middle" align="middle" width="100%" bgcolor=#3366cc>                    
                    <input type="hidden" name="">         <!--*只為了加一個欄位*--> 
                    <input type="hidden" name="">         <!--*只為了加一個欄位*-->
                    <input type="hidden" name="">         <!--*只為了加一個欄位*-->
                        <div align="center">
                    <%    
                        if szStatus = "Add" then
                    %>                              
                            <input type="button" name="Save" style="background-color:#C9E0F8; border-style: outset; border-width: 2;" value="儲　　存" OnKeyDown="AddByKeyPress()" OnMouseUp="OnSave()" onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">                              
                    <%
                        end if
                    %> 
                            　<input type="button" name="BackBtn" style="background-color:#C9E0F8; border-style: outset; border-width: 2" value="回航次資料設定" OnKeyDown="BackByKeyPress()" OnMouseUp="Back()" onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">
                    <%    
                        if szStatus = "Add" then
                    %>                              
                            　<input type="button" name="ResetPartBtn" style="background-color:#C9E0F8; border-style: outset; border-width: 2" value="清除打勾者" OnKeyDown="ResetPartByKeyPress()" OnMouseUp="ResetPart()" onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">
                            　<input type="button" name="ResetBtn" style="background-color:#C9E0F8; border-style: outset; border-width: 2" value="重置" OnKeyDown="ResetByKeyPress()" OnMouseUp="Reset()" onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">                            
                    <%
                        end if
                    %> 
                        </div>
                    </td>
                    <td width=1><img height=38 src="../image/box5.gif" width=20></td>
                </tr>
            </tbody>
        </table> 
    </td>
</tr>
</table> 

<%
    rs.close
    conn.close
    set rs=nothing
    set conn=nothing
%>

<input type="hidden" name="Status" value=<%=szStatus%>>         <!--狀態-->
<input type="hidden" name="VesselID" value=<%=szVesselIDTmp%>>  <!--船碼-->
<input type="hidden" name="PrevFoundID" value=<%=szPrevFoundID%>>  <!--要keep的倉單-->

</form>
</body>
</html>
