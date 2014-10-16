<!-- #include file = ../GlobalSet/conn.asp -->
<!-- #include file = ../GlobalSet/ShareFun.asp -->
<!--總表查詢儲存-->

<%
	'02-May2005: 權限檢查
	if Session.Contents("GroupType@KEEINS8") < 0 then
		response.redirect "../Login/Login.asp"
	end if

    '
    dim szReportType
    szReportType = request("ReportType")
    
    dim szVesselListID, szVesselLine, nFoundCount
    szVesselListID = request("VesselListID")
    szVesselLine = request("VesselLine")
    nFoundCount = request("FoundCount")

	    
	dim i, bIsSetToChecked, szRemark, szFormID
	for i=1 to nFoundCount
		
		szFormID = request("FormID_" & i)
		bIsSetToChecked = request("Chk_" & i)

		if bIsSetToChecked <> 1 then
			bIsSetToChecked = 0
		end if
		
		szRemark = request("Remark_" & i)
		
		sql = "select FormID from SumReportSearchRemark where FormID='" + szFormID + "' and VesselListID = '" + szVesselListID + "'"
    	
    	set rs = conn.execute(sql)
		
		if not rs.eof then
			sql = "update SumReportSearchRemark set IsSetToChecked = " + CStr(bIsSetToChecked) + ", Remark = '" + szRemark + "'"
			sql = sql + " where FormID='" + szFormID + "' and VesselListID = '" + szVesselListID + "'"
			conn.execute(sql)
		else			
			if bIsSetToChecked <> 0 OR szRemark <> "" then
				sql = "insert into SumReportSearchRemark(FormID, VesselListID, IsSetToChecked, Remark)"
				sql = sql + " values ('" + szFormID + "','" + szVesselListID + "'," + CStr(bIsSetToChecked) + ",'" + szRemark + "')"
				conn.execute(sql)
			end if
		end if
		
		rs.close
	next
	
    conn.close 
    set rs=nothing
    set conn=nothing
    
    response.redirect "FSfun03c.asp?ReportType=1&VesselLine=" + szVesselLine + "&VesselListID=" + szVesselListID
%>