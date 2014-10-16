<!-- #include file = ../GlobalSet/conn.asp -->
<!-- #include file = ../GlobalSet/ShareFun.asp -->
<!-- #include file = FSShare.asp -->
<%
	'02-May2005: 權限檢查
	if Session.Contents("GroupType@" & szSessionKey) < 0 then
		response.redirect "../Login/Login.asp"
	elseif Session.Contents("GroupType@" & szSessionKey) = "" then
		response.redirect "../Login/Login.asp"
	end if

    nGroupType      = Session.Contents("GroupType@" & szSessionKey)
	
	dim szOwner

	szOwner			= CStr(Session.Contents("Owner@" & szSessionKey))
	
	if nGroupType = 2 and szOwner = "" then
		response.redirect "../Login/Login.asp"
	end if

	Function IIf(bExp1, sVal1, sVal2) 
		If (bExp1) Then 
			IIf = sVal1 
		Else 
			IIf = sVal2 
		End If 
	End Function

    '===============接收參數===============
    dim szReportType
    szReportType    = request ("ReportType")   '表單種類
    
    dim szVesselListID, szVesselName, szVesselNo, szYear, szMonth, szDay, szDate, szVesselLine, szSO_ID
    szVesselListID = request ("VesselListID")
    szVesselName   = request ("VesselName")
    szVesselNo     = request ("VesselNo")
    szYear         = request ("Year")
    szMonth        = request ("Month")
    szDay          = request ("Day")
    szSO_ID        = request ("SO")

    szVesselLine   = request ("VesselLine")  
    
        
    '===============查詢資料庫===============  
    dim sql, rs     
    
    '查詢航次的資料
    dim szVesselList, szVesselDate, szVesselOwner
    set rs = nothing
    sql = "select * from VesselList where ID = '" + szVesselListID + "'"
    
    set rs = conn.execute(sql)
    
    if not rs.eof then
        szVesselName = rs("VesselName")
        szVesselList = rs("VesselNo")
        szVesselOwner= rs("Owner")
        szVesselDate = rs("Date")
    end if
%>
<!--倉單資料查詢-->
<html>
<head>
<title>倉單資料查詢</title>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <style type="text/css">
        table.grid_view {
            border-collapse: collapse;
        }

            table.grid_view td {
                border: 1px solid #000000;
            }

        table.grid_view2 {
            border-collapse: collapse;
        }

            table.grid_view2 td {
                padding: 0;
                margin: 0;
                border-collapse: collapse;
                font-family: Arial;
                font-size: 14.0pt;
                text-align: right;
            }
    </style>
	<link href="../GlobalSet/sg.css" rel="stylesheet" type="text/css" />
	<script src="//ajax.googleapis.com/ajax/libs/jquery/1.10.2/jquery.min.js"></script>
	<script type="text/javascript">
		(function ($) {
			$.url = {};
			$.extend($.url, { _params: {}, init: function () { var paramsRaw = ""; try { paramsRaw = (document.location.href.split("?", 2)[1] || "").split("#")[0].split("&") || []; for (var i = 0; i < paramsRaw.length; i++) { var single = paramsRaw[i].split("="); if (single[0]) this._params[single[0]] = unescape(single[1]); } } catch (e) { alert(e); } }, param: function (name) { return this._params[name] || ""; }, paramAll: function () { return this._params; } });
			$.url.init();
		})(jQuery);
	</script>
<%
	if nGroupType = 1 then
%>
	<script type="text/javascript" src="../GlobalSet/ShareFun.js"></script>
	<script type="text/javascript" src="FSfun03c.js"></script>
<%
    else
%>
	<script type="text/javascript" src="../GlobalSet/ShareFun2.js"></script>
	<script type="text/javascript" src="FSfun03c2.js"></script>
<%
	end if
%>
	<script src="../GlobalSet/jquery.PrintArea.js"></script>
	<script>
		var user_type=<%= nGroupType %>;
	</script>
	<script type="text/javascript">
		$(document).ready(function () {
		    $(".check_all").click(function () {
		        if ($.isFunction($(this).prop)) {
		            if ($(this).prop("checked")) $(".printNo").each(function () { $(this).prop("checked", true); });
		            else $(".printNo").each(function () { $(this).prop("checked", false); });
		        } else {
		            if ($(this).attr("checked")) $(".printNo").each(function () { $(this).attr("checked", true); });
		            else $(".printNo").each(function () { $(this).attr("checked", false); });
		        }
			});
			$(".print").click(function () {
				var searchIDs = $(".printNo:checked").map(function () { return $(this).val(); }).get();
				if (searchIDs.length > 0)
					window.open("PrintOfSelect.aspx?vlid=" + $.url.param("VesselListID") + "&line=" + $.url.param("VesselLine") + "&t=" + user_type + "&ids=" + searchIDs);
				else alert("未選取任何項目");
			});
			$(".print-local").click(function () {
			    var searchIDs = $(".printNo:checked").map(function () { return $(this).val(); }).get();
			    if (searchIDs.length > 0)
			        window.open("PrintOfSelect2.aspx?vlid=" + $.url.param("VesselListID") + "&line=" + $.url.param("VesselLine") + "&t=" + user_type + "&ids=" + searchIDs);
			    else alert("未選取任何項目");
			});
			$(".save").click(function () {
				var searchIDs = $(".printNo:checked").map(function () { return $(this).val(); }).get();
				if (searchIDs.length > 0)
					window.open("Save/SummaryHandler.ashx?f=" + $(this).attr("format") + "&vlid=" + $.url.param("VesselListID") + "&line=" + $.url.param("VesselLine") + "&t=" + user_type + "&ids=" + searchIDs);
				else alert("未選取任何項目");
			});
			$(".save-local").click(function () {
				var searchIDs = $(".printNo:checked").map(function () { return $(this).val(); }).get();
				if (searchIDs.length > 0)
					window.open("Save/SizeHandler.ashx?f=" + $(this).attr("format") + "&vlid=" + $.url.param("VesselListID") + "&line=" + $.url.param("VesselLine") + "&t=" + user_type + "&ids=" + searchIDs);
				else alert("未選取任何項目");
			});
			$(".printNo").click(function () {
			    if ($.isFunction($(this).prop)) {
			        if ($(this).prop("checked"))
			            $(".printNo[row-group=" + $(this).attr("row-group") + "]").prop("checked", true);
			    } else {
			        if ($(this).attr("checked"))
			            $(".printNo[row-group=" + $(this).attr("row-group") + "]").attr("checked", true);
			    }
			});
		});
	</script>
</head>
<body style="background-image: url(../image/Logo.gif); background-position: center; background-repeat: no-repeat; background-position-y: top;">
	<div style="margin: 0px auto; margin-top: 100px; text-align: center;">
		<span style="font-size: large;">上 林 公 證 有 限 公 司<br />KINGSWOOD SURVEY & MEASURER LTD</span>
		<br />
		<br />
		<span style="font-size: small;">版權所有 &copy 2013. All Rights Reserved.</span>
	</div>
<!------------------------列表----------------------------------->
<form name="form" method="post" onkeydown="CheckHotKey(); CheckPrivateHotKey();">
<table style="width: 100%;">
<%
    dim nNeededVolume
        
    dim nFoundCount

	dim r_szOwnerList, r_szOwnerTmp
	
	dim r_szInStr

	dim szOwners

    '===================詳細尺寸資料==============================
    if szReportType = 2 then
	
		r_szOwnerList		= ""
		
		if nGroupType = 2 and szOwner <> "" then
			szOwners		= Split(szOwner, ",")
			
			For Each szOwnerTmp2 In szOwners
				if szOwnerTmp2 <> "-" And szOwnerTmp2 <> "" then
					szOwnerTmp2 = Trim(szOwnerTmp2)
	
					r_szOwnerList = r_szOwnerList + "'" + szOwnerTmp2 + "',"
				end if
			Next
		end if

		if r_szOwnerList <> "" then
			r_szOwnerList	= Left(r_szOwnerList, Len(r_szOwnerList) - 1)
		end if
%>
	<tr>
		<td>
			<div style="width: 95%; margin: 0 auto;">
			<table border="0" cellspacing="1" align="center" bgcolor="#C9E0F8" style="background-color: #c9e0f8; width: 100%;">
				<tr>
					<td colspan="2" style="font-size: 14pt; color: #0000ff">
						<a href="#" class="save-local" format="txt" vesselName="<%=szVesselName%>" vesselList="<%=szVesselList%>" vesselDate="<%=szVesselDate%>">
							<span>另存新檔(txt)</span>
						</a>
						<a href="#" class="save-local" format="csv" vesselName="<%=szVesselName%>" vesselList="<%=szVesselList%>" vesselDate="<%=szVesselDate%>">
							<span>另存新檔(csv)</span>
						</a>
					</td>
					<td colspan="2" align="right" style="font-size: 14pt; color: #0000ff">
<%
	if nGroupType = 1 then
%>
						<a href="javascript:history.back();" style="margin-left: 24px;">
							<span>上一頁</span>
						</a>
						<a href="../FreightForm/FFfun02b.asp?Status=FirstLoad&VesselListID=<%=szVesselListID%>" style="margin-left: 24px;">
							<span>倉單查詢(F8)</span>
						</a>
						<a href="FSfun03c.asp?ReportType=1&VesselLine=<%=szVesselLine%>&VesselListID=<%=szVesselListID%>" style="margin-left: 24px;">
							<span>總表查詢(Ctrl+Z)</span>
						</a>
<%
	else
%>
						<a href="javascript:history.back();" style="margin-left: 24px;">
							<span>上一頁</span>
						</a>
<%
	end if
%>
						<a href="#" style="margin-left: 24px;" class="print-local" object="report01" vesselName="<%=szVesselName%>" vesselList="<%=szVesselList%>" vesselDate="<%=szVesselDate%>">
							<span>列印</span>
						</a>
					</td>
				</tr>
				<tr>
					<td colspan="4" align="center" style='font-size: 18.0pt'>尺寸資料查詢</td>
				</tr>
				<tr>
					<td colspan="4" align="center">
						<hr size="0" />
					</td>
				</tr>
				<tr valign="middle">
					<td width="1%"></td>
					<td width="20%" style="font-family: Arial; font-size: 14.0pt">船名：<%=szVesselName%></td>
					<td width="20%" style="font-family: Arial; font-size: 14.0pt">航次：<%=szVesselList%></td>
					<td width="69%" style="font-family: Arial; font-size: 14.0pt">結關日期：<%=szVesselDate%></td>
				</tr>
				<tr>
					<td colspan="4" align="center">
						<hr size="0" />
					</td>
				</tr>
			</table>
			</div>
			<div id="report01" style="width: 95%; margin: 0 auto;">
			<table style="width: 100%; background-color: #c9e0f8; border-width: 0;" class="grid_view2">
				<tr style="text-align: right; font-size: 14.0pt;">
					<td>
						<label>列印</label>
						<input type="checkbox" checked="checked" style="border-color: #C9E0F8" class="check_all" target="printNo" />
					</td>
					<td>頁次</td>
					<td>單號</td>
					<td>倉位</td>
					<td>板數</td>
					<td>堆量</td>
					<td>件數</td>
					<td>包裝</td>
					<td>長</td>
					<td>寬</td>
					<td>高</td>
					<td>體積</td>
					<td>單重</td>
					<td>總重</td>
				</tr>
<%
                nFoundCount = 0
                dim nPrevFormID, szBgColor, nSubTotalPiece, fSubTotalVolume

                nPrevFormID = ""
                nTotalPiece = 0
                nSubTotalPiece  = 0

                fVolume = 0
                fSubTotalVolume = 0
                
                for i = 0 to nCount                     
                    
                    '查倉單資料
                    set rs = nothing
                    
					if nGroupType = 1 Or nGroupType = 3 then
						sql		= ""
						sql		= sql + "SELECT FreightForm.*, 'X' as OwnerID "
						sql		= sql + "  FROM FreightForm "
						sql		= sql + " WHERE VesselID = '" + szVesselListID + "' "
					else
						sql		= ""
						sql		= sql + "SELECT DISTINCT FreightForm.*, FormToOwner.OwnerID "
						sql		= sql + "  FROM FreightForm, FormToOwner "
						sql		= sql + " WHERE FreightForm.VesselID = '" + szVesselListID + "' "
						sql		= sql + "   AND FormToOwner.FormID = FreightForm.ID "
					'	sql		= sql + "   AND FormToOwner.OwnerID = '" + szOwner + "' "
						sql		= sql + "   AND FormToOwner.VesselLine = '" + szVesselLine + "' "

						if r_szOwnerList = "" then
							sql	= sql + "   AND FormToOwner.OwnerID = '" + szOwner + "' "
						else
							sql	= sql + "   AND FormToOwner.OwnerID IN (" + r_szOwnerList + ") "
						end if
					end if

					if szSO_ID <> "" then
					   sql		= sql + "   AND FreightForm.ID = '" + szSO_ID + "' "
					end if
	
                    if szStartIDTmp(i) <> "" and szEndIDTmp(i) <> "" then
                        sql		= sql + " AND FreightForm.ID >= '" + szStartIDTmp(i) + "' AND FreightForm.ID <= '" + szEndIDTmp(i) + "'"
                    end if

                    '20-Jul2005: 以輸入順序顯示 (The priority is higher than PageNo)
                    sql = sql + " ORDER BY FreightForm.ID, FreightForm.SN, FreightForm.PageNo"
                    
                    set rs = conn.execute(sql)
	
                    while not rs.eof
                        If nPrevFormID = "" Then
                            nPrevFormID = rs("ID")
                            nSubTotalPiece  = 0
                            fSubTotalVolume = 0
                        ElseIf nPrevFormID <> rs("ID") Then
%>
				<tr style="font-family: Arial; text-align: right; font-size: 14.0pt;">
                    <td></td>
					<td>&nbsp;<!--頁次--></td>
					<td>&nbsp;<!--單號--></td>
					<td>&nbsp;<!--倉位--></td>
					<td>&nbsp;<!--板數--></td>
					<td>&nbsp;<!--堆量--></td>
					<td><%=nSubTotalPiece%></td>
					<td>&nbsp;<!--包裝--></td>
					<td>&nbsp;<!--長--></td>
					<td>&nbsp;<!--寬--></td>
					<td>&nbsp;<!--高--></td>
					<td><%=fSubTotalVolume%></td>
					<td>&nbsp;<!--單重--></td>
					<td>&nbsp;<!--總重--></td>
				</tr>
				<tr>
					<td colspan="14" style="border-left-width: 0; border-right-width: 0;">
						<table style="width: 100%;">
							<tr>
								<td style="border-bottom: 1px solid #000000; height: 5px;"></td>
							</tr>
							<!--<tr>
								<td style="border-top: 1px solid #000000; height: 5px;"></td>
							</tr>-->
						</table>
					</td>
				</tr>
<%
                            nPrevFormID = rs("ID")
                            nSubTotalPiece  = 0
                            fSubTotalVolume = 0
                        End If

                        nFoundCount = nFoundCount + 1
                        
                        fVolume = rs("Volume")
                        
                        Dim nBoard
                        If rs("Board") <> 0 Then
                            nBoard = rs("Board")
                            If IsNumeric(nBoard) Then
                                nBoard = CLng(nBoard)
                            Else
                                nBoard = 1  'just for aborting error
                            End If                            
                        End If
                        
                        dim bChangeBgColor
                        bChangeBgColor = 0
                        
                        if IsIrregularVolume(nBoard, fVolume) = 1 then
			            	bChangeBgColor = 1
			            end if
			                  
			            if IsIrregularLength(rs("IsPL"), rs("Length")) = 1 then
			            	bChangeBgColor = 1
			            end if
			                  
			            if IsIrregularHeight(rs("IsPL"), rs("Height")) = 1 then
			            	bChangeBgColor = 1
			            end if
			            
			            if IsIrregularWidth(rs("IsPL"), rs("Width")) = 1 then
			            	bChangeBgColor = 1
			            end if
                        
                        If bChangeBgColor = 1 Then
                            szBgColor = "#FFFFA0"
                        else
                            szBgColor = "#C9E0F8"   
                        end if
                   
						dim szPackageStyle, sql2, rs2
						szPackageStyle = ""
						if rs("PackageStyleID") <> "0" then
						    sql2 = "select * from PackageStyle where ID = '" + CStr(rs("PackageStyleID")) + "'"
						    set rs2 = conn.execute(sql2)
						    if not rs2.eof then
						        szPackageStyle = rs2("Name")                                
						    end if
						    
						    rs2.close
						    set rs2 = nothing
						end if

                        '20-Nov2004:如果單重為0, 則不要show
                        dim WeightTmp
                        if rs("Weight") = 0 then
                            WeightTmp = "　"
                        else
                            WeightTmp = MyFormatNumber(rs("Weight"), 1)
                        end if

                        '20-Nov2004:如果總重為0, 則不要show
                        dim TotalWeightTmp                                
                        if rs("TotalWeight") = 0 then
                            TotalWeightTmp = "　"
                        else
                            TotalWeightTmp = MyFormatNumber(rs("TotalWeight"), 1)
                        end if
%>
				<tr style="background-color: <%=szBgColor%>;">
                    <td>
						<input type="checkbox" checked="checked" name="printNo" value="<% = rs("SN") %>" row-group="<% = rs("ID") %>" class="printNo" style="border-color: #C9E0F8" />
                    </td>
					<td><%=rs("PageNo") %></td>
					<td><%=rs("ID") %></td>
					<td><%=rs("Storehouse") %></td>
					<% '20-Nov2004:如果板數為0, 則不要show %>
					<td><%=IIf(rs("Board") = 0, "　", rs("Board"))%></td>
					<td><%=IIf(rs("IsPL") = TRUE, "堆量", "　")%></td>
					<td><%=rs("Piece") %></td>
					<td><%=IIf(szPackageStyle = "", "　", szPackageStyle)%></td>
					<td><%=IIf(rs("Length") = "", "　", rs("Length"))%></td>
					<td><%=IIf(rs("Width") = "", "　", rs("Width"))%></td>
					<td><%=IIf(rs("Height") = "", "　", rs("Height"))%></td>
					<td><%=IIf(rs("Volume") = "", "　", FormatNumber(rs("Volume"), 2))%></td>
					<% '20-Nov2004:如果單重為0, 則不要show %>
					<td><%=WeightTmp%></td>
					<td style="text-align: right;"><%=TotalWeightTmp%></td>
				</tr>
<%
                        nTotalPiece = nTotalPiece + rs("Piece")
                        nSubTotalPiece = nSubTotalPiece + rs("Piece")

                        fTotalVolume = fTotalVolume + rs("Volume")
                        fSubTotalVolume = fSubTotalVolume + rs("Volume")

                        nNeededVolume = rs("NeededForestry")
                        
                        rs.movenext
                    wend
                    
                next
%>
				<tr>
                    <td></td>
					<td>&nbsp;<!--頁次--></td>
					<td>&nbsp;<!--單號--></td>
					<td>&nbsp;<!--倉位--></td>
					<td>&nbsp;<!--板數--></td>
					<td>&nbsp;<!--堆量--></td>
					<td><%=nSubTotalPiece%></td>
					<td>&nbsp;<!--包裝--></td>
					<td>&nbsp;<!--長--></td>
					<td>&nbsp;<!--寬--></td>
					<td>&nbsp;<!--高--></td>
					<td><%=fSubTotalVolume%></td>
					<td>&nbsp;<!--單重--></td>
					<td>&nbsp;<!--總重--></td>
				</tr>
				<tr>
					<td colspan="14">
						<table style="width: 100%;">
							<tr>
								<td style="border-bottom: 1px solid #000000; height: 5px;"></td>
							</tr>
							<tr>
								<td style="border-top: 1px solid #000000; height: 5px;"></td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
                    <td></td>
					<td>&nbsp;<!--頁次--></td>
					<td>&nbsp;<!--單號--></td>
					<td>&nbsp;<!--倉位--></td>
					<td>&nbsp;<!--板數--></td>
					<td>&nbsp;<!--堆量--></td>
					<td><%=nTotalPiece%></td>
					<td>&nbsp;<!--包裝--></td>
					<td>&nbsp;<!--長--></td>
					<td>&nbsp;<!--寬--></td>
					<td>&nbsp;<!--高--></td>
					<td><%=fTotalVolume%></td>
					<td>&nbsp;<!--單重--></td>
					<td>&nbsp;<!--總重--></td>
				</tr>
			</table>
			</div>
		</td>
	</tr>
	<tr style="text-align: center;">
		<td>
<%
        if nFoundCount = 0 then
            response.write "查詢不到任何資料"
        end if
%>
		</td>
	</tr>
	<tr>
		<td>
			<hr />
		</td>
	</tr>
<%
    '===================總表格式==============================
    elseif szReportType = 1 then
	
		dim nCurIndex
		dim szOwnerTmp
		dim ChkTmp, szRemarkTmp
        
        nFoundCount = 0
        
        dim nTotalPiece, nTotalVolume, nTotalWeight, j
        nNeededVolume = 0
        
		'查詢出底下的list
		set rs = nothing 
		sql = "select ID,Name from FreightOwner order by ID"
		set rs = conn.execute(sql) 
	
		Dim objDictionary
		Set objDictionary = CreateObject("Scripting.Dictionary")
		objDictionary.CompareMode = vbTextCompare 'makes the keys case insensitive'
	
		dim r_szIDTmp, r_szNameTmp, r_szSelected

		while not rs.eof
			r_szIDTmp			= rs("ID")
			r_szNameTmp			= rs("Name")

			objDictionary.Add r_szIDTmp, r_szNameTmp

			rs.movenext
		wend
	
		r_szOwnerList			= ""
		
		if nGroupType = 2 and szOwner <> "" then
			'r_szOwnerList	= szOwner
			
			szOwners		= Split(szOwner, ",")
			
			For Each szOwnerTmp2 In szOwners
				if szOwnerTmp2 <> "-" And szOwnerTmp2 <> "" then
					szOwnerTmp2 = Trim(szOwnerTmp2)
	
					r_szOwnerList = r_szOwnerList + "'" + szOwnerTmp2 + "',"
				end if
			Next
		end if

		if r_szOwnerList <> "" then
			r_szOwnerList	= Left(r_szOwnerList, Len(r_szOwnerList) - 1)
		end if
%>
	<tr>
		<td>
			<div style="width: 95%; margin: 0 auto;">
				<table border="0" cellspacing="1" align="center" bgcolor="#C9E0F8" style="background-color: #c9e0f8; width: 100%;">
					<tr>
						<td style="font-size: 14pt; color: #0000ff">
							<a href="#" class="save" format="txt" vesselName="<%=szVesselName%>" vesselList="<%=szVesselList%>" vesselDate="<%=szVesselDate%>">
								<span>另存新檔(txt)</span>
							</a>
							<a href="#" class="save" format="csv" vesselName="<%=szVesselName%>" vesselList="<%=szVesselList%>" vesselDate="<%=szVesselDate%>">
								<span>另存新檔(csv)</span>
							</a>
						</td>
						<td colspan="3" style="font-size: 14pt; color: #0000ff; text-align: right;">
<%
		if nGroupType = 1 then
%>
							<a href="javascript:history.back();" style="margin-left: 24px;">
								<span>上一頁</span>
							</a>
							<a href="#" onclick="OnSave()" style="margin-left: 24px;">
								<span>儲存(F2)</span>
							</a>
							<a href="../FreightForm/FFfun02b.asp?Status=FirstLoad&VesselListID=<%=szVesselListID%>" style="margin-left: 24px;">
								<span>倉單查詢(F8)</span>
							</a>
							<a href="FSfun03c.asp?ReportType=2&VesselLine=<%=szVesselLine%>&VesselListID=<%=szVesselListID%>" style="margin-left: 24px;">
								<span>尺寸資料查詢(Ctrl+X)</span>
							</a>
<%
		else
%>
							<a href="javascript:history.back();" style="margin-left: 24px;">
								<span>上一頁</span>
							</a>
							<a target="_blank" href="FSfun03c.asp?ReportType=2&VesselLine=<% = szVesselLine %>&VesselListID=<% = szVesselListID %>" style="margin-left: 24px;">
								<span>尺寸資料查詢(Ctrl+X)</span>
							</a>
<%
		end if
%>
							<!--<a target="_blank" href="FSfun01b.asp?Status=FirstLoad&VesselListID=<% = szVesselListID %>" style="margin-left: 24px;">
								<span>攬貨報告書列印/傳真</span>
							</a>-->
							<a href="#" style="margin-left: 24px;" class="print" object="report11" vesselName="<%=szVesselName%>" vesselList="<%=szVesselList%>" vesselDate="<%=szVesselDate%>">
								<span>列印(主要S/O)</span>
							</a>
						</td>
					</tr>
				</table>
				<table border="0" cellspacing="1" align="center" bgcolor="#C9E0F8" style="background-color: #c9e0f8; width: 100%;">
					<tr>
						<td colspan="4" style="font-size: 18.0pt; text-align: center;">總表查詢</td>
					</tr>
					<tr>
						<td colspan="4">
							<hr />
						</td>
					</tr>
					<tr style="font-family: Arial; font-size: 14.0pt">
						<td style="width: 1%;"></td>
						<td style="width: 20%;">船名：<%=szVesselName%></td>
						<td style="width: 20%;">航次：<%=szVesselList%></td>
						<td style="width: 69%;">結關日期：<%=szVesselDate%></td>
					</tr>
				</table>
			</div>
			<div id="report11" style="width: 95%; margin: 0 auto;">
			<table class="grid_view" align="center" style="background-color: #c9e0f8; width: 100%; border: 1px solid #000000;">
				<tr style="text-align: right; font-family: Arial; font-size: 14.0pt; border: 1px solid #000000;">
					<td>
						<label>列印</label>
						<input type="checkbox" checked="checked" style="border-color: #C9E0F8" class="check_all" target="printNo" />
					</td>
<%
		if nGroupType = 1 then
%>
					<td>選項</td>
<%
		end if
%>
					<td>攬貨商編號</td>
					<td>攬貨商</td>
					<td>核對</td>
					<td>單號S/O</td>
					<td>件數</td>
					<td>體積</td>
					<td>板數</td>
					<!--<td width="13%">單重</td>-->
					<td>總重</td>
<%
		if nGroupType = 1 then
%>
					<td>備註</td>
<%
		end if
%>
				</tr>
<%
		nTotalPiece = 0
		nTotalVolume = 0
		nTotalWeight = 0
		nTotalBoard = 0
                    
		dim i, rs1
        dim FoId, FoName
                    
		for i = 0 to nCount
			'查倉單資料
			set rs = nothing                            
                        
			if nGroupType = 1 then
				sql = "select ID, IsChecked, NeededForestry, sum(Piece) as TotalPiece, sum(Volume) as TotalVolume,"
				sql = sql + " sum(TotalWeight) as TotalWeightSum, sum(Board) as TotalBoard"
				sql = sql + " from FreightForm where VesselID = '" + szVesselListID + "' "
			
				if szStartIDTmp(i) <> "" and szEndIDTmp(i) <> "" then
				   sql = sql + " AND FreightForm.ID >= '" + szStartIDTmp(i) + "' and FreightForm.ID <= '" + szEndIDTmp(i) + "' "
				end if
			elseif nGroupType = 3 then
				sql		= ""
				sql		= sql + "SELECT FreightForm.ID, FreightForm.IsChecked, FreightForm.NeededForestry, "
				sql		= sql + "       SUM(FreightForm.Piece) as TotalPiece, "
				sql		= sql + "       SUM(FreightForm.Volume) as TotalVolume, "
				sql		= sql + "       SUM(FreightForm.TotalWeight) as TotalWeightSum, "
				sql		= sql + "       SUM(FreightForm.Board) as TotalBoard, "
				sql		= sql + "       iif(isnull(FreightOwner.ID), '', FreightOwner.ID) as FO_ID, "
				sql		= sql + "       iif(isnull(FreightOwner.Name), '', FreightOwner.Name) as FO_Name, "
				sql		= sql + "       iif(isnull(FormToOwner.VesselLine), '', FormToOwner.VesselLine) as VesselLine "
				sql		= sql + "  FROM ((FreightForm "
				sql		= sql + "  LEFT OUTER JOIN FormToOwner ON FormToOwner.FormID = FreightForm.ID) "
				sql		= sql + "  LEFT OUTER JOIN FreightOwner ON FormToOwner.OwnerID = FreightOwner.ID) "
				sql		= sql + " WHERE FreightForm.VesselID = '" + szVesselListID + "' "
				sql		= sql + "   AND (FormToOwner.VesselLine = '" + szVesselLine + "' OR FormToOwner.VesselLine IS NULL) "
			else
				sql		= ""
				sql		= sql + "SELECT FreightForm.ID, FreightForm.IsChecked, FreightForm.NeededForestry, "
				sql		= sql + "       SUM(FreightForm.Piece) as TotalPiece, "
				sql		= sql + "       SUM(FreightForm.Volume) as TotalVolume, "
				sql		= sql + "       SUM(FreightForm.TotalWeight) as TotalWeightSum, "
				sql		= sql + "       SUM(FreightForm.Board) as TotalBoard, "
				sql		= sql + "       iif(isnull(FreightOwner.ID), '', FreightOwner.ID) as FO_ID, "
				sql		= sql + "       iif(isnull(FreightOwner.Name), '', FreightOwner.Name) as FO_Name, "
				sql		= sql + "       iif(isnull(FormToOwner.VesselLine), '', FormToOwner.VesselLine) as VesselLine "
				sql		= sql + "  FROM ((FreightForm "
				sql		= sql + "  LEFT OUTER JOIN FormToOwner ON FormToOwner.FormID = FreightForm.ID) "
				sql		= sql + "  LEFT OUTER JOIN FreightOwner ON FormToOwner.OwnerID = FreightOwner.ID) "
				sql		= sql + " WHERE FreightForm.VesselID = '" + szVesselListID + "' "
				sql		= sql + "   AND (FormToOwner.VesselLine = '" + szVesselLine + "' OR FormToOwner.VesselLine IS NULL) "
	
				if r_szOwnerList = "" then
					if szOwner = "" then
						sql	= sql + "   AND FreightOwner.ID = '' "
					else
						sql	= sql + "   AND FreightOwner.ID = '" + szOwner + "' "
					end if
				else
					sql	= sql + "   AND FreightOwner.ID IN (" + r_szOwnerList + ") "
				end if
			end if
                        
			if nGroupType = 1 then
				sql = sql + " GROUP BY FreightForm.ID, FreightForm.IsChecked, FreightForm.NeededForestry "
				sql = sql + " ORDER BY FreightForm.ID"
			else
				sql = sql + " GROUP BY FreightForm.ID, FreightForm.IsChecked, FreightForm.NeededForestry, FreightOwner.ID, FreightOwner.Name, FormToOwner.VesselLine "
				sql = sql + " ORDER BY FreightOwner.ID DESC"
			end if
		
			set rs = conn.execute(sql) 
	
			nCurIndex = 0
	
			while not rs.eof
			    nNeededVolume = rs("NeededForestry")
			    nFoundCount = nFoundCount + 1
			    
				sql = "select * from SumReportSearchRemark where FormID='" + rs("ID") + "' and VesselListID = '" + szVesselListID + "'"
				set rs1 = conn.execute(sql)
				
				if not rs1.eof then
					if rs1("IsSetToChecked") = TRUE then
						ChkTmp = 1
					end if
					
					szRemarkTmp = rs1("Remark")
				else
					ChkTmp = 0
					szRemarkTmp = ""
				end if

			 	nCurIndex = nCurIndex + 1

				if nGroupType = 1 then
					sql = "select FreightOwner.ID, FreightOwner.Name from FreightOwner, FormToOwner"
					sql = sql + " where FormToOwner.FormID = '" + rs("ID") + "'"
					sql = sql + " and FormToOwner.OwnerID = FreightOwner.ID"
					sql = sql + " and FormToOwner.VesselLine = '" + szVesselLine + "'"
					
					set rs1 = conn.execute(sql)
					
					if not rs1.eof then
						szOwnerTmp			= rs1("ID")
					    FoId                = rs1("ID")
                        FoName              = rs1("Name")
					else
						szOwnerTmp			= ""
					    FoId                = ""
                        FoName              = ""
					end if 
				
					rs1.close
					set rs1 = nothing
				else
					szOwnerTmp				= rs("FO_ID")
					FoId                    = rs("FO_ID")
                    FoName                  = rs("FO_Name")
				end if
%>
				<tr style="text-align: right; font-family: Arial; font-size: 14.0pt; border: 1px solid #000000;">
					<td>
						<input type="checkbox" checked="checked" name="printNo" value="<% = rs("ID") %>" row-group="<% = FoId %>" class="printNo" style="border-color: #C9E0F8" />
					</td>
<%
				nCurIndex = nCurIndex + 1
                            	
				if nGroupType = 1 then
					if ChkTmp = 1 then
%>
					<td>
						<input type="hidden" name="FormID_<%=nFoundCount%>" value='<% = rs("ID") %>' />
						<input type="checkbox" checked="checked" name="Chk_<%=nFoundCount%>" value="1" style="border-color: #C9E0F8" />
					</td>
<%
					else
%>
					<td>
						<input type="hidden" name="FormID_<%=nFoundCount%>" value='<% = rs("ID") %>' />
						<input type="checkbox" name="Chk_<%=nFoundCount%>" value="1" style="border-color: #C9E0F8" />
					</td>
<%
					end if
				end if
%>
                    <td><% = FoId %></td>
                    <td><% = FoName %></td>
					<td>
<%
				if rs("IsChecked") = FALSE then
				    response.write "未核對"
				else
					response.write "　"
				end if
%>
					</td>
					<td>
						<a href="FSfun03c.asp?ReportType=2&VesselLine=<%=szVesselLine%>&VesselListID=<%=szVesselListID%>&SO=<%=rs("ID")%>"><%=rs("ID")%></a>
					</td>
					<td><%=rs("TotalPiece")%></td>
					<td>
<%
				if nNeededVolume <> 0 then
				    response.write FormatNumber(nNeededVolume, 2)
				else
				    response.write FormatNumber(rs("TotalVolume"), 2)
				end if
%>
					</td>
					<td>
<%
				if rs("TotalBoard") <> 0 then
				    response.write rs("TotalBoard")
				else
					response.write "　"
				end if
%>
					</td>
					<td>
<%
				if rs("TotalWeightSum") <> 0 then
				    response.write MyFormatNumber(rs("TotalWeightSum"), 1)
				else
					response.write "　"
				end if
%>
					</td>
<%
				nCurIndex = nCurIndex + 1
				
				if nGroupType = 1 then
%>
					<td>
						<select>
<%
				For Each Entry In objDictionary
					r_szSelected			= ""
				
					if Entry = szOwnerTmp then
						r_szSelected		= "selected"
					end if
%>
							<option value="<% = Entry %>" <% = r_szSelected %>>(<% = Entry %>)<% = objDictionary(Entry) %></option>
<%
				Next
%>
						</select>
						<input type="text" name="Remark_<%=nFoundCount%>" size="3" value="<%=szRemarkTmp%>"
							onkeydown="ChangeFocus(<% = nCurIndex + 2 %>)" onfocusin="SetFocusStyle(this, true, true)"
							onfocusout="SetFocusStyle(this, false, false)" />
					</td>
<%
				end if
%>
				</tr>
<%
				nTotalPiece = nTotalPiece + rs("TotalPiece")
	        	
				if nNeededVolume <> 0 then
					nTotalVolume = nTotalVolume + nNeededVolume
				else
					nTotalVolume = nTotalVolume + rs("TotalVolume")
				end if
	        	
				nTotalWeight = nTotalWeight + rs("TotalWeightSum")
	        	
				nTotalBoard = nTotalBoard + rs("TotalBoard")
	    
				response.flush

				rs.movenext
			wend
		next

		if nGroupType = 1 then
%>
				<tr>
					<td></td>
					<td colspan="8">
						<table style="width: 100%;">
							<tr>
								<td style="border-bottom: 1px solid #000000; height: 5px;"></td>
							</tr>
							<tr>
								<td style="border-top: 1px solid #000000; height: 5px;"></td>
							</tr>
						</table>
					</td>
					<td></td>
				</tr>
				<tr style="text-align: right; font-family: Arial; font-size: 14.0pt;">
					<td></td>
					<td>TOTAL：</td>
					<td></td>
					<td></td>
					<td></td>
					<td><%=nTotalPiece%></td>
					<td><%=FormatNumber(nTotalVolume, 2)%></td>
					<td><%=nTotalBoard%></td>
					<td><%=nTotalWeight%></td>
					<td></td>
				</tr>
<%
		else
%>
				<tr>
					<td></td>
					<td colspan="8">
						<table style="width: 100%;">
							<tr>
								<td style="border-bottom: 1px solid #000000; height: 5px;"></td>
							</tr>
							<tr>
								<td style="border-top: 1px solid #000000; height: 5px;"></td>
							</tr>
						</table>
					</td>
				</tr>
				<tr style="text-align: right; font-family: Arial; font-size: 14.0pt;">
					<td></td>
					<td>TOTAL：</td>
					<td></td>
					<td></td>
					<td></td>
					<td><%=nTotalPiece%></td>
					<td><%=FormatNumber(nTotalVolume, 2)%></td>
					<td><%=nTotalBoard%></td>
					<td><%=nTotalWeight%></td>
				</tr>
<%
		end if
%>
			</table>
			</div>
		</td>
	</tr>
	<tr style="text-align: center;">
		<td>
<%
        if nFoundCount = 0 then
            response.write "查詢不到任何資料"
        end if
%>
		</td>
	</tr>
	<tr>
		<td>
			<br />
			<hr />
			<br />
		</td>
	</tr>
<%
		if nGroupType <> 1 then
%>
	<tr>
		<td>
			<div style="width: 95%; margin: 0 auto;">
				<table border="0" cellspacing="1" width="100%" align="center" bgcolor="#C9E0F8">
					<tr>
						<td colspan="4" align="right" style="font-size: 14pt; color: #0000ff">
							<a href="#" style="margin-left: 24px;" class="print-local" object="report12" title="總表查詢" vesselName="<%=szVesselName%>" vesselList="<%=szVesselList%>" vesselDate="<%=szVesselDate%>">
								<span>列印(剩餘S/O)</span>
							</a>
						</td>
					</tr>
				</table>
				<table border="0" cellspacing="1" width="100%" align="center" bgcolor="#C9E0F8">
					<tr>
						<td colspan="4" style="font-size: 18.0pt; text-align: center;">總表查詢</td>
					</tr>
					<tr>
						<td colspan="4">
							<hr />
						</td>
					</tr>
					<tr style="font-family: Arial; font-size: 14.0pt">
						<td style="width: 1%;"></td>
						<td style="width: 20%;">船名：<%=szVesselName%></td>
						<td style="width: 20%;">航次：<%=szVesselList%></td>
						<td style="width: 69%;">結關日期：<%=szVesselDate%></td>
					</tr>
				</table>
			</div>
			<div id="report12" style="width: 95%; margin: 0 auto;">
			<table class="grid_view" align="center" style="background-color: #c9e0f8; width: 100%; border: 1px solid #000000;">
				<tr style="text-align: right;font-size: 14.0pt;">
					<td class="print-remove"></td>
					<td class="print-remove">攬貨商編號</td>
					<td>攬貨商</td>
					<td>核對</td>
					<td>單號S/O</td>
					<td>件數</td>
					<td>體積</td>
					<td class="print-remove">板數</td>
					<td>總重</td>
<%
		if nGroupType = 1 then
%>
					<td class="print-remove">備註</td>
<%
		end if
%>
				</tr>
<%
		nTotalPiece = 0
		nTotalVolume = 0
		nTotalWeight = 0
		nTotalBoard = 0
                    
		'dim i, rs1
                    
		for i = 0 to nCount
			'查倉單資料
			set rs = nothing

			if nGroupType = 1 then
				sql = "select ID, IsChecked, NeededForestry, sum(Piece) as TotalPiece, sum(Volume) as TotalVolume,"
				sql = sql + " sum(TotalWeight) as TotalWeightSum, sum(Board) as TotalBoard"
				sql = sql + " from FreightForm where VesselID = '" + szVesselListID + "' "
			
				if szStartIDTmp(i) <> "" and szEndIDTmp(i) <> "" then
				   sql = sql + " AND FreightForm.ID >= '" + szStartIDTmp(i) + "' and FreightForm.ID <= '" + szEndIDTmp(i) + "'"
				end if
			else
				sql		= ""
				sql		= sql + "SELECT FreightForm.ID, FreightForm.IsChecked, FreightForm.NeededForestry, "
				sql		= sql + "       SUM(FreightForm.Piece) as TotalPiece, "
				sql		= sql + "       SUM(FreightForm.Volume) as TotalVolume, "
				sql		= sql + "       SUM(FreightForm.TotalWeight) as TotalWeightSum, "
				sql		= sql + "       SUM(FreightForm.Board) as TotalBoard, "
				sql		= sql + "       iif(isnull(FreightOwner.ID), '', FreightOwner.ID) as FO_ID, "
				sql		= sql + "       iif(isnull(FreightOwner.Name), '', FreightOwner.Name) as FO_Name, "
				sql		= sql + "       iif(isnull(FormToOwner.VesselLine), '', FormToOwner.VesselLine) as VesselLine "
				sql		= sql + "  FROM ((FreightForm "
				sql		= sql + "  LEFT OUTER JOIN FormToOwner ON FormToOwner.FormID = FreightForm.ID) "
				sql		= sql + "  LEFT OUTER JOIN FreightOwner ON FormToOwner.OwnerID = FreightOwner.ID) "
				sql		= sql + " WHERE FreightForm.VesselID = '" + szVesselListID + "' "
				sql		= sql + "   AND (FormToOwner.VesselLine = '" + szVesselLine + "' OR FormToOwner.VesselLine IS NULL) "
				sql		= sql + "   AND FreightOwner.ID IS NULL "
			end if
                        
			if nGroupType = 1 then
				sql = sql + " GROUP BY FreightForm.ID, FreightForm.IsChecked, FreightForm.NeededForestry "
				sql = sql + " ORDER BY FreightForm.ID"
			else
				sql = sql + " GROUP BY FreightForm.ID, FreightForm.IsChecked, FreightForm.NeededForestry, FreightOwner.ID, FreightOwner.Name, FormToOwner.VesselLine "
				sql = sql + " ORDER BY FreightOwner.ID DESC"
			end if

			set rs = conn.execute(sql) 
	
			nCurIndex = 0
                        
			while not rs.eof
			    nNeededVolume = rs("NeededForestry")
			    nFoundCount = nFoundCount + 1
			    
			    sql = "select * from SumReportSearchRemark where FormID='" + rs("ID") + "' and VesselListID = '" + szVesselListID + "'"
				set rs1 = conn.execute(sql)
				
				if not rs1.eof then
					if rs1("IsSetToChecked") = TRUE then
						ChkTmp = 1
					end if
					
					szRemarkTmp = rs1("Remark")
				else
					ChkTmp = 0
					szRemarkTmp = ""
				end if

			 	nCurIndex = nCurIndex + 1
%>
				<tr style="text-align: right;font-size: 14.0pt;">
					<td style="border-bottom: 1px solid #000000; height: 5px;" class="print-remove">
<%
				nCurIndex = nCurIndex + 1
                            	
				'if ChkTmp = 1 then
%>
						<!--<input type="hidden" name="FormID_<%=nFoundCount%>" value='<% = rs("ID") %>' />
						<input type="checkbox" checked name="Chk_<%=nFoundCount%>" value="1" style="border-color: #C9E0F8" />-->
<%
				'else
%>
						<!--<input type="hidden" name="FormID_<%=nFoundCount%>" value='<% = rs("ID") %>' />
						<input type="checkbox" name="Chk_<%=nFoundCount%>" value="1" style="border-color: #C9E0F8" />-->
<%
				'end if
%>
					</td>
<%
				if nGroupType = 1 then
					sql = "select FreightOwner.ID, FreightOwner.Name from FreightOwner, FormToOwner"
					sql = sql + " where FormToOwner.FormID = '" + CStr(rs("ID")) + "'"
					sql = sql + " and FormToOwner.OwnerID = FreightOwner.ID"
					sql = sql + " and FormToOwner.VesselLine = '" + szVesselLine + "'"
					
					set rs1 = conn.execute(sql)
					
					if not rs1.eof then
					    response.write "<td class=""print-remove"">" + rs1("ID") + "</td>"
					    response.write "<td>" + rs1("Name") + "</td>"
					else
					    response.write "<td class=""print-remove"">　</td>"
					    response.write "<td>　</td>"
					end if 
					
					rs1.close
					set rs1 = nothing
				else
					response.write "<td class=""print-remove"">" + rs("FO_ID") + "</td>"
					response.write "<td>" + rs("FO_Name") + "</td>"
				end if
%>
					<td style="border-bottom: 1px solid #000000; height: 5px;">
<%
				if rs("IsChecked") = FALSE then
				    response.write "未核對"
				else
					response.write "　"
				end if
%>
					</td>
					<td style="border-bottom: 1px solid #000000; height: 5px;"><%=rs("ID")%></td>
					<td style="border-bottom: 1px solid #000000; height: 5px;"><%=rs("TotalPiece")%></td>
					<td style="border-bottom: 1px solid #000000; height: 5px;">
<%
				if nNeededVolume <> 0 then
				    response.write FormatNumber(nNeededVolume, 2)
				else
				    response.write FormatNumber(rs("TotalVolume"), 2)
				end if
%>
					</td>
					<td style="border-bottom: 1px solid #000000; height: 5px;" class="print-remove">
<%
				if rs("TotalBoard") <> 0 then
				    response.write rs("TotalBoard")
				else
					response.write "　"
				end if
%>
					</td>
					<td style="border-bottom: 1px solid #000000; height: 5px;">
<%
				if rs("TotalWeightSum") <> 0 then
				    response.write MyFormatNumber(rs("TotalWeightSum"), 1)
				else
					response.write "　"
				end if
%>
					</td>
<%
				nCurIndex = nCurIndex + 1
%>
				</tr>
<%
				nTotalPiece = nTotalPiece + rs("TotalPiece")
	        
				if nNeededVolume <> 0 then
					nTotalVolume = nTotalVolume + nNeededVolume
				else
					nTotalVolume = nTotalVolume + rs("TotalVolume")
				end if
	        
				nTotalWeight = nTotalWeight + rs("TotalWeightSum")
	        
				nTotalBoard = nTotalBoard + rs("TotalBoard")
	    
				rs.movenext
			wend
		next
%>
				<tr style="text-align: right; font-family: Arial; font-size: 14.0pt;">
					<td style="border-top: 1px solid #000000; height: 5px;" class="print-remove"></td>
					<td style="border-top: 1px solid #000000; height: 5px;">TOTAL：</td>
					<td style="border-top: 1px solid #000000; height: 5px;" class="print-remove"></td>
					<td style="border-top: 1px solid #000000; height: 5px;" class="print-remove"></td>
					<td style="border-top: 1px solid #000000; height: 5px;"></td>
					<td style="border-top: 1px solid #000000; height: 5px;"><%=nTotalPiece%></td>
					<td style="border-top: 1px solid #000000; height: 5px;"><%=FormatNumber(nTotalVolume, 2)%></td>
					<td style="border-top: 1px solid #000000; height: 5px;"><%=nTotalBoard%></td>
					<td style="border-top: 1px solid #000000; height: 5px;"><%=nTotalWeight%></td>
				</tr>
			</table>
			</div>
		</td>
	</tr>
	<tr style="text-align: center;">
		<td>
<%
        if nFoundCount = 0 then
            response.write "查詢不到任何資料"
        end if
%>
		</td>
	</tr>
	<tr>
		<td>
			<br />
			<hr />
			<br />
		</td>
	</tr>
<%
		end if
    end if

	if nGroupType = 1 then
%>
	<tr style="text-align: center;">
		<td>
<%
        if szReportType = 1 then
        	response.write "<input type=""hidden"" name=""hidden_tmp"">"  '只是為了set focus到正確位置而加的
        	response.write "<input type=""hidden"" name=""hidden_tmp"">"  '只是為了set focus到正確位置而加的
        	nCurIndex = nCurIndex + 1
        	response.write "<input name=""Save"" type=button style=""border-color:#C9E0F8; border-style: outset; border-width: 2;"" value=""儲存"" OnKeyDown=""ChangeFocus(" & nCurIndex+2 & "); SaveByKeyPress()"" OnMouseUp=""OnSave()"" onfocusin=""SetFocusStyle(this, true, true)"" onfocusout=""SetFocusStyle(this, false, true)"">"
        	response.write "　　"
        end if
%>
			<input name="Re_Search" type="button" style="border-color:#C9E0F8; border-style: outset; border-width: 2px" value="重新查詢" OnKeyDown="ReSearchByKeyPress()" OnMouseUp="OnReSearch()" onfocusin="SetFocusStyle(this, true, true)" onfocusout="SetFocusStyle(this, false, true)" />
		</td>
	</tr>
<%
	end if
%>  
</table>
<!--表單種類-->
<input type="hidden" name="ReportType" value="<% = szReportType %>" />
<!--航次-->
<input type="hidden" name="VesselListID" value="<% = szVesselListID %>" />
<!--航線-->
<input type="hidden" name="VesselLine" value="<% = szVesselLine %>" />
<!--筆數-->
<input type="hidden" name="FoundCount" value="<% = nFoundCount %>" />
<%
    rs.close
    conn.close
    
    set rs=nothing
    set conn=nothing
%>
</form>
</body>
</html>