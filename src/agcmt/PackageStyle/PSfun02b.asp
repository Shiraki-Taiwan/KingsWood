<!-- #include file = ../GlobalSet/conn.asp -->
<!--包裝方式列印-->
<html>
<head>
<title>包裝方式列印</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<LINK href="../GlobalSet/sg.css" rel=stylesheet type=text/css>
</head>

<%
	'02-May2005: 權限檢查
	if Session.Contents("GroupType@AGCMT") < 0 then
		response.redirect "../Login/Login.asp"
	end if


    '===============接收參數===============
     dim szID, szName
    
    szID    = request("ID")     '編號   
    szName  = request("Name")   '名稱
    

    '===============查詢資料庫===============  
    dim sql, rs 
    
    set rs = nothing 
    sql = "select * from PackageStyle where ID like '%" + szID + "%' and Name like '%" + szName + "%' order by ID"
    
    set rs = conn.execute(sql) 
    
%>

<!------------------------列表----------------------------------->

<table width="100%" height=""100%"">
    <tr> 
        <td > 
        <table border=0 cellspacing=1 height=""100%"" align="center">
            <tr>
                <td colspan=2 align="center"><font size=5>包裝方式名稱列表</font></td>
            </tr>
            <tr>
                <td colspan=2 align="center">
                <%
                    dim i, nLineLen
                    nLineLen = 112
                    
                    for i = 0 to nLineLen
                        response.write "="
                    next
                %>
                </td>
            </tr>
            <tr valign="center"> 
                <td width="10%" align="left">編號</td>
                <td width="90%" align="left">名稱</td>
            </tr>
            <tr>
                <td colspan=2 align="center">
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
                    <td width="10%" align="left"><%= rs("ID")%></td>
                    <td width="90%" align="left"><%= rs("Name")%></td>
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