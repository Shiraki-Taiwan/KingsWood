<!-- #include file = ../GlobalSet/conn.asp -->
<!--公司資料列印-->
<html>
<head>
<title>公司資料列印</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<LINK href="../GlobalSet/sg.css" rel=stylesheet type=text/css>
</head>

<%
	'02-May2005: 權限檢查
	if Session.Contents("GroupType@TAIINS2") < 0 then
		response.redirect "../Login/Login.asp"
	end if


    '===============接收參數===============
    dim szID, szChineseName, szEnglishName, szFaxNo_1, szFaxNo_2, szPhone_1, szPhone_2, szAddress
    
    szID            = request("ID")               '編號
    szChineseName   = request("ChineseName")      '公司中文名稱
    szEnglishName   = request("EnglishName")      '公司英文名稱
    szFaxNo_1       = request("FaxNo_1")          '傳真電話1
    szFaxNo_2       = request("FaxNo_2")          '傳真電話2
    szPhone_1       = request("Phone_1")          '聯絡電話1
    szPhone_2       = request("Phone_2")          '聯絡電話2
    szAddress       = request("Address")          '地址
    

    '===============查詢資料庫===============  
    dim sql, rs 
    set rs = nothing 
  
    sql = "select * from CompanyInfo where ID like '%" + szID + "%' and ChineseName like '%" + szChineseName + "%'"
    sql = sql + " and EnglishName like '%" + szEnglishName + "%'"
    
    set rs = conn.execute(sql) 
    
%>

<!------------------------列表----------------------------------->

<table width="100%" height=""100%"">
    <tr> 
        <td > 
        <table border=0 cellspacing=1 height=""100%"" align="center">
            <tr>
                <td colspan=8 align="center"><font size=5>公司資料列表</font></td>
            </tr>
            <tr>
                <td colspan=8 align="center">
                <%
                    dim i, nLineLen
                    nLineLen = 112
                    
                    for i = 0 to nLineLen
                        response.write "="
                    next
                %>
                </td>
            </tr>
            <tr align="center"> 
                <td width="5%" align="left">編號</td>
                <td width="15%" align="left">公司中文名稱</td>
                <td width="15%" align="left">公司英文名稱</td>
                <td width="10%" align="left">傳真電話1</td>
                <td width="10%" align="left">傳真電話2</td>
                <td width="10%" align="left">聯絡電話1</td>
                <td width="10%" align="left">聯絡電話2</td>
                <td width="25%" align="left">地址</td>
            </tr>
            <tr>
                <td colspan=8 align="center">
                <%
                    for i = 0 to nLineLen
                        response.write "="
                    next
                %>
                </td>
            </tr>
       
            <%
                while not rs.eof
            %>
                <tr align="center"> 
                    <td width="5%" align="left"><%= rs("ID")%></td>
                    <td width="15%" align="left"><%= rs("ChineseName")%></td>
                    <td width="15%" align="left"><%= rs("EnglishName")%></td>
                    <td width="10%" align="left"><%= rs("FaxNo_1")%></td>
                    <td width="10%" align="left"><%= rs("FaxNo_2")%></td>
                    <td width="10%" align="left"><%= rs("Phone_1")%></td>
                    <td width="10%" align="left"><%= rs("Phone_2")%></td>
                    <td width="25%" align="left"><%= rs("Address")%></td>
                </tr>
            <%
                    rs.movenext
                wend
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