<%
	SUB ShowText(conn, byref TextString)
      	dim rs2,i,sql2
        i=0
        sql2 = "select ID from StoreSum where VesselListID='" + szVesselID + "'"
        TextString = "function Repeat(a_nCurIndex){"      
        TextString = TextString & "var FormIDTmp;"
        TextString = TextString & "FormIDTmp = document.form[a_nCurIndex].value;"
        TextString = TextString & "FormIDTmp = FormatString_JS(FormIDTmp, 4);"
        
        set rs2 = conn.execute("select count(ID) as TotalCount from StoreSum")
        if not rs2.eof then
            i = rs2("TotalCount")
            TextString = TextString & "Show =new Array("& i &");"
            rs2.movenext
        end if 
        set rs2 = nothing
        i=0
        set rs2 = conn.execute(sql2)
        while not rs2.eof 
            TextString = TextString & "Show["&i&"]="""& rs2("ID") &""";"
            i = i + 1
            rs2.movenext
        wend
        
        rs2.close
        set rs2 = nothing
        
        TextString = TextString & "var LimitArray = " & i & ";" 
        TextString = TextString & "for( i = 0 ; i < LimitArray ;i++)"       
        TextString = TextString & "{"
        TextString = TextString & "    if(Show[i] == FormIDTmp)"
        TextString = TextString & "    {"
        TextString = TextString & "        alert(""此單號已核定過,請先取消核定再新增或修改!"");"
        TextString = TextString & "        document.form[a_nCurIndex].focus();"
        TextString = TextString & "    }"
        TextString = TextString & "}"
        TextString = TextString & "}"
    END SUB
    
    
%>