<!-- #include file = ../GlobalSet/conn.asp -->
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title>攬貨商-單號對照資料</title>
	<!-- #include file = ../head_bundle.html -->
	<script type="text/javascript" src="FFfun03b.js"></script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="10" marginwidth="0" marginheight="0" onload="initialiseMenu(), CLocation();">
	<!-- #include file = ../title.htm -->
<%
function AddOne(Num)
  DIM TEMP(3)
	Dim WithDigital
	Dim Index
	WithDigital=False
	Num=ucase(Num)
	'response.write "NUM="&Num&"<BR>"

  for i=0 to 3
    if (not IsNumeric(Mid(Num,i+1,1))) then 
    	WithDigital=True
    	Index=i
    end if
  next

  if (WithDigital) then
    for i=0 to 3
      TEMP(i)=asc(mid(num,i+1,1))
    next
    TEMP(Index)=TEMP(Index)+1
    AddOne=chr(TEMP(0))+chr(TEMP(1))+chr(TEMP(2))+chr(TEMP(3))  
  else
  	Num=cint(Num)+1  	
  	Num=cstr(Num)
  	for l =0 to 3- len(Num)	'補零
      Num="0"+Num
    next 
    AddOne=Num
  end if	  
end function
 
%>
<%
	 '02-May2005: 權限檢查
	 if Session.Contents("GroupType@AGCMT") < 0 then
	 	response.redirect "../Login/Login.asp"
	 end if
 
 
    dim sql, rs
    set rs = nothing
    
    '===========接參數===========
    '21-Nov2004: 新增航線欄位, 因為不同航線可能會有同樣的單號
    dim szFormID, szOwnerID, szVesselLine, szStatus, szFromIDTmp, szOwnerIDTmp, szVesselLineTmp
    szFormID     = request ("FormID")        '單號
    szOwnerID    = request ("OwnerID")        '攬貨商
    szVesselLine = request ("VesselLine")     '航線
    szStatus     = request ("Status")         '狀態
   
    if szStatus = "" then
        szStatus = "Add"
    end if
    
    if szStatus = "Modify" then
    
        sql = "select * from FormToOwner where FormID = '" + szFormID + "' and VesselLine = '" + szVesselLine + "' order by OwnerID, FormID"
        
        set rs = conn.execute(sql)
        
        if not rs.eof then
            szFromIDTmp = Ucase(rs("FormID"))
            szOwnerIDTmp = rs("OwnerID")
            szVesselLineTmp = rs("VesselLine")
            
            szFormID = ""
            szOwnerID = ""
            szVesselLine = ""
        end if
        
    else
        szVesselLineTmp = szVesselLine             
    end if
           
    set rs = nothing                                                
    
    '查詢航線
    dim szVesselLineName
    sql = "select Name from VesselLine where ID = '" + szVesselLineTmp + "' order by ID"
    
    set rs = conn.execute(sql)

    if not rs.eof then
        szVesselLineName = rs("Name")        
    end if
%>


<form name="form" method="post" OnKeyDown="CheckHotKey()">
<table cellspacing=0 cellpadding=0 width="100%" border="0" align="center" >
    <tr> 
        <td > 
            <table cellspacing=0 cellpadding=0 width="100%" border=0>
                <tbody> 
                    <tr bgcolor=#3366cc > 
                        <td width=1><img height=26 src="../image/coin2ltb.gif" width=20></td>
                        <td align=middle width="100%" ><div align="center"><font color="#FFFFFF" >攬貨商-單號對照資料</font></div></td>
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
                                        <td width="40%" align=right bgcolor=#C9E0F8 height="2">航　線：</td>
                                        <td width="60%" align=left bgcolor=#C9E0F8 height="2"><%=szVesselLineName%></td>
                                    </tr>
            	                    <tr> 
            	                        <td align=right bgcolor=#C9E0F8 height="2" width="30%">單　號：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2" width="70%"> 
                                            <input type="text" name="FormID" size="30" value="<%=szFromIDTmp%>" onfocus=SelectText(0) OnKeyDown="ChangeFocus(1)" onfocusin="SetFocusStyle(this, true, false)" onfocusout="SetFocusStyle(this, false, false)"
                                                <%
                                                    if szStatus = "Modify" then 
                                                        response.write "readonly" 
                                                    end if
                                                %> 
                                            >
                                        </td>
                                           
                                    </tr>
                                    <tr>      
                                        <td align=right bgcolor=#C9E0F8 height="2">攬貨商：</td>
                                        <td align=left bgcolor=#C9E0F8 height="2">
                                            <select size="1" name="OwnerID" OnKeyDown="ChangeFocus(2)">             
                                                <option value="0">--請選擇--</option>
                                                <%   
                                                    dim rs1
                                                    set rs1 = nothing                                                
                                                    
                                                    '查詢攬貨商
                                                    sql = "select ID, Name from FreightOwner order by ID"
                                                    
                                                    set rs1 = conn.execute(sql)
                                               
                                                    while not rs1.eof
                                                        if rs1("ID") = szOwnerIDTmp then 
                                                            response.write "<option selected value=" & rs1("ID") & ">" & rs1("ID") & "-" & rs1("Name") &"</option>"            		 	
                                                        else
                                                            response.write "<option value=" & rs1("ID") & ">" & rs1("ID") & "-" & rs1("Name") & "</option>"            		 	
                                                        end if
                                                	    rs1.movenext
                                                    wend
                                                    
                                                    rs1.close
                                                    set rs1 = nothing
                                                %>
                                            </select>
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
                                <input type="button" style="background-color:#C9E0F8;border-style: outset; border-width: 2" value="儲存" OnKeyDown="AddByKeyPress()" OnMouseUp="Add()" onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">
                                <input type="button" style="background-color:#C9E0F8;border-style: outset; border-width: 2" value="查詢" OnKeyDown="SearchByKeyPress()" OnMouseUp="Search()" onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">
                                <input type="button" style="background-color:#C9E0F8;border-style: outset; border-width: 2" value="刪除" OnKeyDown="DeleteByKeyPress()" OnMouseUp="Delete()" onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">
                                <input type="button" style="background-color:#C9E0F8;border-style: outset; border-width: 2" value="重置" OnKeyDown="ResetByKeyPress()" OnMouseUp="Reset()" onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">
                                <input type="button" style="background-color:#C9E0F8;border-style: outset; border-width: 2" value="回上一頁" OnKeyDown="BackByKeyPress()" OnMouseUp="Back()" onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">
                                
                            </div>
                        </td>
                        <td width=1><img height=38 src="../image/box5.gif" width=20></td>
                    </tr>
                </tbody>
            </table> 
        </td>
    </tr>
</table>

<input type="hidden" name="Status" value=<%=szStatus%>>
<input type="hidden" name="FormIDToModify" value=<%=szFromIDTmp%>>
<input type="hidden" name="VesselLine" value=<%=szVesselLineTmp%>>


<br>

<%
    '查詢出底下的list
    set rs = nothing 
    sql = "select FormToOwner.*, FreightOwner.Name from FormToOwner, FreightOwner where FreightOwner.ID = FormToOwner.OwnerID"
    sql = sql + " and FormToOwner.FormID like '%" + szFormID + "%' and FormToOwner.VesselLine='" + szVesselLineTmp + "'"
    
    if szOwnerID <> "0" then
        sql = sql + " and FormToOwner.OwnerID like '%" + szOwnerID + "%'"
    end if
    
    sql = sql + " order by OwnerID, FormID"

    set rs = conn.execute(sql)   
%>

<A name=D></A>
<table cellspacing=0 cellpadding=0 width="100%" border="0" align="center" >
    <tr> 
        <td > 
        <table bgcolor=#C9E0F8 border=0 bordercolor=#000000 cellspacing=1 width=100% align="center">
            <tr bgcolor="#3366cc" align="center"> 
                <td width="20%" align="left"><font color="#FFFFFF">攬貨商</font></td>
                <td width="80%" align="left"><font color="#FFFFFF">單號</font></td>
            </tr>
        </table>
        </td>
    </tr>
    
    <tr> 
        <td > 
            <table width="100%" cellspacing=1 border=0>		
            <%
                dim fChangeOwner, nFormIDCount, i, nMaxColNum
                szOwnerID = ""
                nMaxColNum = 4
                fChangeOwner = 0
                nFormIDCount = 4
                
                dim fWriteFirstFormID, nCurFormIDTmp, nPreFormIDTmp, szPreFormID, nRangeCount, bGotLastFormID
                dim szFirstFormID, szFirstFormID_Num
                fWriteFirstFormID = 0
                nRangeCount = 0
                bGotLastFormID = 0                
                dim szPreFormIDTxtPart, szCurFormIDTxtPart, szFormIDNumPart, szCurFormIDNumPart, bGotNumberPart
                
                while not rs.eof
                    if szOwnerID <> rs("OwnerID") then
                        szOwnerID = rs("OwnerID")
                        fChangeOwner = 1
                    end if
                    
			'----show id-攬貨商
                    if fChangeOwner = 1 then                    
                        if fWriteFirstFormID and nRangeCount > 1 then
                            response.write " - " + "<a href=""FFfun03b.asp?Status=Modify&FormID=" +  szPreFormID + "&VesselLine=" + szVesselLineTmp + """>" + szPreFormID + "</a></td>"
                        end if
                        
                        for i = nFormIDCount to nMaxColNum - 1
                            response.write "<td width=""20%"" bgcolor=#C9E0F8 align=""left""></td>"
                        next
                        
                        response.write "<tr align=""center"">"
                        response.write "<td width=""20%"" bgcolor=#C9E0F8 align=""left""><a href=""FFfun03d.asp?Status=Modify&ID=" + szOwnerID + "&VesselLine=" + szVesselLineTmp + """>" + rs("OwnerID") + "-" + rs("Name") + "</a></td>"
                        fChangeOwner = 0
                        nFormIDCount = 0  
                        nPreFormIDTmp = 0
                        szPreFormID = ""
                        fWriteFirstFormID = 0
                        nRangeCount = 0
                    end if
                    
                    if fWriteFirstFormID = 0 then  
			'第一筆
                        response.write "<td width=""20%"" bgcolor=#C9E0F8 align=""left""><a href=""FFfun03b.asp?Status=Modify&FormID=" +  rs("FormID") + "&VesselLine=" + szVesselLineTmp + """>" + rs("FormID") + "</a>"
                        fWriteFirstFormID = 1
                        nPreFormIDTmp = rs("FormID")
                        nRangeCount = nRangeCount + 1
                    else
                        nCurFormIDTmp = rs("FormID")
'response.write "AddOne(nPreFormIDTmp)="&AddOne(nPreFormIDTmp)&"<BR>"
'response.write "nCurFormIDTmp="&nCurFormIDTmp&"<BR>"
			                  if nCurFormIDTmp<>AddOne(nPreFormIDTmp) then 	'--------不連續時
                           if fWriteFirstFormID and nRangeCount > 1 then
                             response.write " - " + "<a href=""FFfun03b.asp?Status=Modify&FormID=" +  szPreFormID + "&VesselLine=" + szVesselLineTmp + """>" + szPreFormID + "</a></td>"
                           else
                             response.write "</td>"
                           end if  
                           if nFormIDCount = nMaxColNum - 1 then
                             response.write "<tr align=""center"">"
                             response.write "<td width=""20%"" bgcolor=#C9E0F8 align=""left""></td>"
                             nFormIDCount = -1  
                           end if
                            
                           response.write "<td width=""20%"" bgcolor=#C9E0F8 align=""left""><a href=""FFfun03b.asp?Status=Modify&FormID=" +  rs("FormID") + "&VesselLine=" + szVesselLineTmp + """>" + rs("FormID") + "</a>"
                           nFormIDCount = nFormIDCount + 1
                           nPreFormIDTmp = rs("FormID")                                                            
                           nRangeCount = 1                                                                               
                        else			'-----------連號
                           nPreFormIDTmp = nCurFormIDTmp
                           szPreFormID = rs("FormID")
                           nRangeCount = nRangeCount + 1
                        end if
                    end if
                    
                    rs.movenext
                wend
                
                if fWriteFirstFormID and nRangeCount > 1 then
                    response.write " - " + "<a href=""FFfun03b.asp?Status=Modify&FormID=" +  szPreFormID + "&VesselLine=" + szVesselLineTmp + """>" + szPreFormID + "</a></td>"
                else
                    response.write "</td>"
                end if                
                
                for i = nFormIDCount to nMaxColNum - 1
                    response.write "<td width=""20%"" bgcolor=#C9E0F8 align=""left""></td>"
                next
            %>
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
 
 
</form>
</body>
</html>
