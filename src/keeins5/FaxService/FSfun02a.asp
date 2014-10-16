<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
	<title>傳真結果查詢</title>
	<!-- #include file = ../head_bundle.html -->
	<script type="text/javascript" src="FSfun02a.js"></script>
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="10" marginwidth="0" marginheight="0" onload="initialiseMenu(), CLocation();">
	<!-- #include file = ../title.htm -->
<%
    '查folder中有那些檔案
    Dim objFileSystem, objFile, Files, File, szLogFile(1000), nFileCounter, nTotalFileCounter
    
    nFileCounter = 0
    nTotalFileCounter = 0

    dim szFilePath
    szFilePath = Request.ServerVariables ("PATH_TRANSLATED")
    
    nPos = InStrRev (szFilePath, "\")
    
    szFilePath = Mid (szFilePath, 1, nPos)
    
    szFilePath = szFilePath + "FaxFiles\FailedList\"
    
    if szFilePath <> "" then
        Set objFileSystem = CreateObject ("Scripting.FileSystemObject")
        set objFile = objFileSystem.GetFolder(szFilePath)
        Set Files = objFile.Files
        
        For Each File in Files
            nTotalFileCounter = nTotalFileCounter + 1
            szLogFile(nTotalFileCounter-1) = File.name
        Next

    end if
    
    '27-Dec2004: 切頁
    dim nMaxItemCount, nPageNum, nTotalPage
    nMaxItemCount = 15
    
    nPageNum = request ("PageNum")
    if nPageNum = "" then
        nPageNum = 1
    else
        nPageNum = CLng(nPageNum)
    end if
    
    if (nTotalFileCounter mod nMaxItemCount) = 0 then
        nTotalPage = (nTotalFileCounter \ nMaxItemCount)
    else
        nTotalPage = (nTotalFileCounter \ nMaxItemCount) + 1
    end if
    
%>

<form name="form" method="post" action="FSfun02b.asp" OnKeyDown="CheckHotKey()">
<table cellspacing=0 cellpadding=0 width="100%" border="0" align="center" >
<tr> 
    <td > 
        <table cellspacing=0 cellpadding=0 width="100%" border=0>
            <tbody> 
                <tr bgcolor=#3366cc > 
                    <td width=1><img height=26 src="../image/coin2ltb.gif" width=20></td>
                    <td align=middle width="100%" ><div align="center"><font color="#FFFFFF" >傳真結果查詢</font></div></td>
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
                                
                                <tr><td colspan=2 bgcolor=#C9E0F8><br></td></tr>                               
        	                    <%
        	                        dim szTimeTmp, szDate, szTime, nPos
        	                        dim bBreak, j
        	                       
        	                        'j = nTotalFileCounter - (nPageNum - 1) * nMaxItemCount - 1
        	                        j = (nPageNum - 1) * nMaxItemCount
        	                         
        	                        bBreak = false
        	                        
        	                        while not bBreak
        	                        'for j = (nPageNum - 1) * nMaxItemCount to nPageNum * nMaxItemCount - 1 'nFileCounter-1
        	                            
        	                            if j < nTotalFileCounter then
            	                            nPos = InStrRev (szLogFile(j), "\")
                                            szTimeTmp = Mid (szLogFile(j), nPos+1)
                                            
                                            szDate = Mid (szTimeTmp, 1, 2) & "/" & Mid (szTimeTmp, 3, 2) & "/" & Mid (szTimeTmp, 5, 2)
                                            szTime = Mid (szTimeTmp, 8, 2) & ":" & Mid (szTimeTmp, 10, 2)
                    
            	                            response.write "<tr><td align=right width=""35%"" bgcolor=#C9E0F8 >傳真紀錄：</td>"
                                            response.write "<td bgcolor=#C9E0F8 width=""65%"" >"
                                            response.write "<input type=checkbox style=""border-color: #C9E0F8"" name=LogFileState_" & CStr(nFileCounter) & " value=1>" & szDate & "　" & szTime & "　" & "(" & szLogFile(j) & ")"
                                            response.write "</td></tr>"
                                            
                                            nFileCounter = nFileCounter + 1
                                            j = j + 1
                                        else
                                            bBreak = true
                                        end if
                                        
                                        if nFileCounter = 15 then
                                            bBreak = true
                                        end if
                                    wend
                                    'next
                                %>
                                
                                <tr><td colspan=2 bgcolor=#C9E0F8><br></td></tr>
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
                            <input type="button" name="Search" style="background-color:#C9E0F8;border-style: outset; border-width: 2" value="顯示細節" OnKeyDown="SearchByKeyPress()" OnMouseUp="OnSearch()" onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">
                            <input type="button" name="Delete" style="background-color:#C9E0F8;border-style: outset; border-width: 2" value="刪除紀錄" OnKeyDown="DeleteByKeyPress()" OnMouseUp="OnDelete()" onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">
                            <input type="button" name="Reset" style="background-color:#C9E0F8;border-style: outset; border-width: 2" value="重置" OnKeyDown="ResetByKeyPress()" OnMouseUp="OnReset()" onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)">
                    <%
                        '19-Dec2004: 加入分頁
                        if nTotalPage > 1 then
                            response.write "　" & nPageNum & "/" & nTotalPage
                            
                            if nPageNum > 2 then
                                response.write "　<a href=FSfun02a.asp?PageNum=1" + ">首頁</a>"
                            else
                                response.write "　首頁"
                            end if 
                            
                            if nPageNum > 1 then
                                response.write "　<a href=FSfun02a.asp?PageNum=" + CStr(nPageNum - 1) + ">上一頁</a>"
                            else
                                response.write "　上一頁"
                            end if
                            
                            if nPageNum < nTotalPage then
                                response.write "　<a href=FSfun02a.asp?PageNum=" + CStr(nPageNum+1) + ">下一頁</a>"
                            else
                                response.write "　下一頁"
                            end if                           
                            
                            if nPageNum + 1 < nTotalPage then
                                response.write "　<a href=FSfun02a.asp?PageNum=" + CStr(nTotalPage) + ">末頁</a>"
                            else
                                response.write "　末頁"
                            end if
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

<input type="hidden" name="Status" value=<%=szStatus%>>
<input type="hidden" name="FileCounter" value=<%=nFileCounter%>>
<%
    dim i
    j = (nPageNum - 1) * nMaxItemCount
    for i = 0 to nFileCounter-1
        response.write "<input type=hidden name=LogFile_" + CStr(i) + " value=""" + szFilePath + szLogFile(j) + """>"
        j = j + 1
    next
%>

</table>   
</form>
</body>
</html>
