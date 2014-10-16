<!-- #include file = ../GlobalSet/conn.asp -->
<!--攬貨商資料列印-->
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title>攬貨商資料列印</title>
	<!-- #include file = ../head_bundle.html -->
	<link href="../GlobalSet/sg.css" rel="stylesheet" type="text/css" />
	<script type="text/javascript" src="FOfun02c.js"></script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="10" marginwidth="0" marginheight="0" onload="Print();">
<%
	'02-May2005: 權限檢查
	if Session.Contents("GroupType@TAIINS1") < 0 then
		response.redirect "../Login/Login.asp"
	end if

    '===============接收參數===============
    dim szID, szName   
       
    szSubStr     = request("ID")        '編號       
    
    
    '===========Parse 編號===================
    dim szStartIDTmp(100), szEndIDTmp(100), nCount, nPos
            
    nCount = 0
    
    nPos = Instr(szSubStr, ",")
    
    while (nPos <> 0)
        szStartIDTmp(nCount) = LTrim(Left(szSubStr, nPos - 1))
        szSubStr = LTrim(Mid (szSubStr, nPos + 1, Len(szSubStr))) 
        
        '找 "-"
        nPos = Instr(szStartIDTmp(nCount), "-")
        if nPos = 0 then
            szEndIDTmp(nCount) = szStartIDTmp(nCount)            
        else
            szEndIDTmp(nCount) = LTrim(Mid (szStartIDTmp(nCount), nPos + 1, Len(szStartIDTmp(nCount))))
            szStartIDTmp(nCount) = LTrim(Left(szStartIDTmp(nCount), nPos - 1))            
        end if
          
        nCount = nCount + 1
        nPos = Instr(szSubStr, ",")
        
    wend
    
    szStartIDTmp(nCount) = LTrim(szSubStr)
    '找 "-"
    nPos = Instr(szStartIDTmp(nCount), "-")
    if nPos = 0 then
        szEndIDTmp(nCount) = szStartIDTmp(nCount)            
    else
        szEndIDTmp(nCount) = LTrim(Mid (szStartIDTmp(nCount), nPos + 1, Len(szStartIDTmp(nCount))))
        szStartIDTmp(nCount) = LTrim(Left(szStartIDTmp(nCount), nPos - 1))            
    end if
        

    '===============查詢資料庫===============  
    dim sql, rs 
    
    'set rs = nothing 
    'sql = "select * from FreightOwner where ID like '%" + szID + "%' and Name like '%" + szName + "%' order by ID"
    
    'set rs = conn.execute(sql) 
    
%>

<!------------------------列表----------------------------------->
<form name="form" method="post" OnKeyDown="CheckHotKey()">
<table width="100%" height=""100%"">
    <tr> 
        <td > 
        <table border=0 cellspacing=1 height=""100%"" align="center">
            <tr>
                <td colspan=9 align="center"><font size=5>攬貨商資料列表</font></td>
            </tr>
            <tr>
                <td colspan=9 align="center"><hr></td>
            </tr>
            <tr valign="center"> 
                <td width="5%" align="left">編號</td>
                <td width="10%" align="left">公司名稱</td>
                <td width="10%" align="left">傳真電話1</td>
                <td width="10%" align="left">傳真電話2</td>
                <td width="10%" align="left">聯絡人1</td>
                <td width="10%" align="left">聯絡電話1</td>
                <td width="10%" align="left">聯絡人2</td>
                <td width="10%" align="left">聯絡電話2</td>
                <td width="25%" align="left">e-Mail</td>
            </tr>
            <tr>
                <td colspan=9 align="center"><hr></td>
            </tr>
       
        <%
            
            for i = 0 to nCount 
                
                set rs = nothing
                
                if nCount = 0 then
                    sql = "select * from FreightOwner order by ID"
                else
                    sql = "select * from FreightOwner where ID >= '" + szStartIDTmp(i) + "' and ID <= '" + szEndIDTmp(i) + "'" + " order by ID"
                end if
                set rs = conn.execute(sql)
                
                while not rs.eof 
        %>
                    <tr align="center"> 
                        <td align="left"><%= rs("ID")%></td>
                        <td align="left"><%= rs("Name")%></td>
                        <td align="left"><%= rs("FaxNo_1")%></td>
                        <td align="left"><%= rs("FaxNo_2")%></td>
                        <td align="left"><%= rs("Contact_1")%></td>
                        <td align="left"><%= rs("Phone_1")%></td>
                        <td align="left"><%= rs("Contact_2")%></td>
                        <td align="left"><%= rs("Phone_2")%></td>
                        <td align="left"><%= rs("MailAddr")%></td>
                    </tr>
        <%
                    rs.movenext
                wend
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