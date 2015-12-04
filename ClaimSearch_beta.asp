<%@  language="VBSCRIPT" %>
<% Option Explicit %>
<%Response.Buffer = True %>
<%Response.Expires = 0 %>
<!--#include file="user.inc" -->
<!--#include file="adovbs_mod.inc" -->
<%
dim adoConn, adoRS, sSQL, sTemp
dim iPageNo, iRowCount, i, iPageCount
dim m_iSequence, m_iDistr, m_dFromDate, m_dThruDate
dim m_iSSN, m_sLastName, m_iDepNum, m_iTaxID, m_sProvName, m_iAssocNum
dim m_sFund, m_sPlan, m_sHWType, m_sStatus1, m_sStatus2, m_sStatus3 
dim m_dStatFrom, m_dStatThru, m_sRSN, m_dServiceFrom
dim m_dServiceThru, m_iBenPeriod, m_dBenFrom, m_dBenThru
dim m_iBillType, m_dAdmitFrom, m_dAdmitThru, m_dDisFrom, m_dDisThru, m_sAdmitType
dim m_dBillFrom, m_dBillThru, m_iConfine, m_iICD9_Proc, m_iDRG, m_iRevCode
dim m_sDiagCode, m_sCovCode, m_sPreCalCode, m_sPayCode, m_iPOSCode, m_cCharge
dim m_sProcCode, m_sModifier, m_sCheckAcct, m_iCheckNum, m_sTrackCode, m_bOverPayCheck
dim m_bOtherAdjustCheck, m_bCOBCheck, m_bCOBOverCheck, m_sPPO, m_bPPOCheck
dim bUseSQL, sNoMatches, m_bShowChargelines
dim m_iCOBOverride

m_iCOBOverride = 100021 'Per Russell, this code (in ErrorNumbers.h) indicates a Claim has a "COB Override"
						'associated with it.

function ConvertToClaimNum(LogDate,Sequence,Distr)
dim sClaimNum

	if month(LogDate) < 10 then
		sClaimNum = "0" & month(LogDate)
	else
		sClaimNum = month(LogDate)
	end if
	if day(LogDate) < 10 then
		sClaimNum = sClaimNum & "0" & day(LogDate)
	else
		sClaimNum = sClaimNum & day(LogDate)
	end if
	sClaimNum = sClaimNum & right(Year(LogDate),2)
	sClaimNum = sClaimNum & "-" & string(5-len(Sequence),"0") & Sequence
	sClaimNum = sClaimNum & "-" & Distr
	ConvertToClaimNum = sClaimNum
	
end function

m_bShowChargelines = false
if Request.Form("UserCriteria") = "Initial" then
	bUseSQL = false
else
	bUseSQL = true
	if Request.Form("ClaimType") <> "Claims" then
		m_bShowChargelines = true
	end if
	sSQL = "select distinct c.LogDate, c.Sequence, c.Distr,"
	if m_bShowChargelines then
		sSQL = sSQL & " g.LineNumber 'Line #',"
	else
		sSQL = sSQL & " g.LineNumber,"
	end if
	sSQL = sSQL & " c.SSN,"
	if m_bShowChargelines then
		sSQL = sSQL & " c.DepNumber 'Dep #',"
	else
		sSQL = sSQL & " c.DepNumber,"
	end if
	sSQL = sSQL & " n.FirstName + ' ' + n.MiddleName + ' ' + n.LastName Name,"
	sSQL = sSQL & " p.FullName Provider, c.TaxID,"
	if m_bShowChargelines then
		sSQL = sSQL & " c.Associate 'Assoc #',"
	else
		sSQL = sSQL & " c.Associate,"
	end if
	sSQL = sSQL & " g.ChargeAmount Charge"
	sSQL = sSQL & " from Claims c left join Charges g on (c.LogDate="
	sSQL = sSQL & " g.LogDate AND c.Sequence=g.Sequence AND c.Distr=g.Distr)"
	sSQL = sSQL & " left join ClaimDiagnosis d on (c.LogDate="
	sSQL = sSQL & " d.LogDate AND c.Sequence=d.Sequence AND c.Distr=d.Distr)"
	sSQL = sSQL & " left join ClaimOverrides o on (c.LogDate="
	sSQL = sSQL & " o.LogDate AND c.Sequence=o.Sequence AND c.Distr=o.Distr)"
	sSQL = sSQL & " left join HospClaim h on (c.LogDate="
	sSQL = sSQL & " h.LogDate AND c.Sequence=h.Sequence AND c.Distr=h.Distr) left join"
	sSQL = sSQL & " BillTypes t on (h.billtype = t.BillType) left join Names n on" 
	sSQL = sSQL & " (c.SSN=n.SSN and c.DepNumber=n.depnumber) left join Providers P on"
	sSQL = sSQL & " (c.TaxID=p.TaxID and c.Associate=p.Associate and p.RecordType = 0)"
	sSQL = sSQL & " left join ChargePreCal r on (g.LogDate=r.LogDate AND g.Sequence ="
	sSQL = sSQL & " r.Sequence AND g.Distr=r.Distr AND g.LineNumber=r.LineNumber)"
	sSQL = sSQL & " left join (select cbp.LogDate,cbp.Sequence,"
	sSQL = sSQL & " cbp.Distr, cbp.LineNumber, bp.FromDate, bp.ThruDate, bp.IncidentNumber"
	sSQL = sSQL & " from ChargeBenefitPeriod cbp left join BenefitPeriod bp on"
	sSQL = sSQL & " (cbp.SSN=bp.SSN AND cbp.DepNumber=bp.DepNumber AND cbp.IncidentNumber"
	sSQL = sSQL & " =bp.IncidentNumber)) bens on (g.LogDate=bens.LogDate AND g.Sequence ="
	sSQL = sSQL & " bens.Sequence AND g.Distr=bens.Distr AND g.LineNumber=bens.LineNumber)"
	sSQL = sSQL & " where "

	set adoConn = Server.CreateObject("ADODB.Connection")
	'
	' Setting CommandTimeout to 0 waits indefinitely for response; note that per 
	' ADO documentation Command object will not inherit the Connection setting (default
	' is 30 seconds).  This appeared to help with response issues on seasql02.
	'
	adoConn.CommandTimeout = 0
	set adoRS = Server.CreateObject("ADODB.Recordset")
	adoConn.Open Application("DataConn")

	if Trim(Request.Form("ActionType")) = "0" then
		iPageNo = 1
	else
		if Trim(Request.Form("ActionType")) = "" then
			iPageNo = 1
		else
			iPageNo = Cint(Request.Form("ActionType"))
			if iPageNo < 1 Then 
				iPageNo = 1
			end if
		end if
	end if
	if Request.Form("FromDate") <> "" then
		m_dFromDate = Request.Form("FromDate")
		sSQL = sSQL & " c.LogDate >= '" & m_dFromDate & "' AND"
	end if
	if Request.Form("ThruDate") <> "" then
		m_dThruDate = Request.Form("ThruDate")
		sSQL = sSQL & " c.LogDate <= '" & dateadd("s",-1,dateadd("d",1,m_dThruDate)) & "' AND"
	end if
	if Request.Form("Sequence") <> "" then
		m_iSequence = Request.Form("Sequence")
		sSQL = sSQL & " c.Sequence =" & m_iSequence & " AND"
	end if
	if Request.Form("Distr") <> "" then
		m_iDistr = Request.Form("Distr")
		sSQL = sSQL & " c.Distr =" & m_iDistr & " AND"
	end if
	if Request.Form("SSN") <> "" then
		m_iSSN = Request.Form("SSN")
		sSQL = sSQL & " c.SSN =" & m_iSSN & " AND"
	end if
	if Request.Form("LName") <> "" then
		m_sLastName = Request.Form("LName")
		sSQL = sSQL & " n.LastName Like '" & m_sLastName & "%' AND"
	end if
	if Request.Form("DepNum") <> "" then
		m_iDepNum = Request.Form("DepNum")
		sSQL = sSQL & " c.DepNumber =" & m_iDepNum & " AND"
	end if
	if Request.Form("TaxID") <> "" then
		m_iTaxID = Request.Form("TaxID")
		sSQL = sSQL & " c.TaxID =" & m_iTaxID & " AND"
	end if
	if Request.Form("ProvName") <> "" then
		m_sProvName = Request.Form("ProvName")
		sSQL = sSQL & " p.FullName Like '" & m_sProvName & "%' AND"
	end if
	if Request.Form("AssocNum") <> "" then
		m_iAssocNum = Request.Form("AssocNum")
		sSQL = sSQL & " c.Associate =" & m_iAssocNum & " AND"
	end if
	if Request.Form("Fund") <> "" then
		m_sFund = Request.Form("Fund")
		sSQL = sSQL & " c.Fund ='" & m_sFund & "' AND"
	end if
	if Request.Form("Plan") <> "" then
		m_sPlan = Request.Form("Plan")
		sSQL = sSQL & " c.PlanCode ='" & m_sPlan & "' AND"
	end if
	if Request.Form("HWType") <> "" then
		m_sHWType = Request.Form("HWType")
		sSQL = sSQL & " c.HWType ='" & m_sHWType & "' AND"
	end if
	if Request.Form("Status1") <> "" or Request.Form("Status2") <> "" or Request.Form("Status3") <> ""then
		sSQL = sSQL & "("
		if Request.Form("Status1") <> "" then
			m_sStatus1 = Request.Form("Status1")
			sSQL = sSQL & " c.Status ='" & m_sStatus1 & "'"
		end if
		if Request.Form("Status2") <> "" then
			m_sStatus2 = Request.Form("Status2")
			if right(sSQL,1) = "(" then
				sSQL = sSQL & " c.Status ='" & m_sStatus2 & "'"
			else
				sSQL = sSQL & " OR c.Status ='" & m_sStatus2 & "'"
			end if
		end if
		if Request.Form("Status3") <> "" then
			m_sStatus3 = Request.Form("Status3")
			if right(sSQL,1) = "(" then
				sSQL = sSQL & " c.Status ='" & m_sStatus3 & "'"
			else
				sSQL = sSQL & " OR c.Status ='" & m_sStatus3 & "'"
			end if
		end if
		sSQL = sSQL & ") AND"
	end if
	if Request.Form("StatFrom") <> "" then
		m_dStatFrom = Request.Form("StatFrom")
		sSQL = sSQL & " c.StatusDate >= '" & m_dStatFrom & "' AND"
	end if
	if Request.Form("StatThru") <> "" then
		m_dStatThru = Request.Form("StatThru")
		sSQL = sSQL & " c.StatusDate <= '" & dateadd("s",-1,dateadd("d",1,m_dStatThru)) & "' AND"
	end if
	if Request.Form("RSN") <> "" then
		m_sRSN = Request.Form("RSN")
		sSQL = sSQL & " c.StatusReason ='" & m_sRSN & "' AND"
	end if
	if Request.Form("ServiceFrom") <> "" then
		m_dServiceFrom = Request.Form("ServiceFrom")
		sSQL = sSQL & " g.FromDate >= '" & m_dServiceFrom & "' AND"
	end if
	if Request.Form("ServiceThru") <> "" then
		m_dServiceThru = Request.Form("ServiceThru")
		sSQL = sSQL & " g.ThruDate <= '" & dateadd("s",-1,dateadd("d",1,m_dServiceFrom)) & "' AND"
	end if
	if Request.Form("BenPeriod") <> "" then
		m_iBenPeriod = Request.Form("BenPeriod")
		sSQL = sSQL & " bens.IncidentNumber =" & m_iBenPeriod & " AND"
	end if
	if Request.Form("BenFrom") <> "" then
		m_dBenFrom = Request.Form("BenFrom")
		sSQL = sSQL & " bens.FromDate >= '" & m_dBenFrom & "' AND"
	end if
	if Request.Form("BenThru") <> "" then
		m_dBenThru = Request.Form("BenThru")
		sSQL = sSQL & " bens.ThruDate <= '" & dateadd("s",-1,dateadd("d",1,m_dBenThru)) & "' AND"
	end if
	if Request.Form("AdmitFrom") <> "" then
		m_dAdmitFrom = Request.Form("AdmitFrom")
		sSQL = sSQL & " h.AdmitDate >= '" & m_dAdmitFrom & "' AND"
	end if
	if Request.Form("AdmitThru") <> "" then
		m_dAdmitThru = Request.Form("AdmitThru")
		sSQL = sSQL & " h.AdmitDate <= '" & dateadd("s",-1,dateadd("d",1,m_dAdmitThru)) & "' AND"
	end if
	if Request.Form("DisFrom") <> "" then
		m_dDisFrom = Request.Form("DisFrom")
		sSQL = sSQL & " h.DischargeDate >= '" & m_dDisFrom & "' AND"
	end if
	if Request.Form("DisThru") <> "" then
		m_dDisThru = Request.Form("DisThru")
		sSQL = sSQL & " h.DischargeDate <= '" & dateadd("s",-1,dateadd("d",1,m_dDisThru)) & "' AND"
	end if
	if Request.Form("AdmitType") <> "" then
		m_sAdmitType = Request.Form("AdmitType")
		sSQL = sSQL & " h.AdmitType = '" & m_sAdmitType & "' AND"
	end if
	if Request.Form("BillType") <> "" then
		m_iBillType = Request.Form("BillType")
		sSQL = sSQL & " h.BillType =" & m_iBillType & " AND"
	end if
	if Request.Form("BillFrom") <> "" then
		m_dBillFrom = Request.Form("BillFrom")
		sSQL = sSQL & " c.BillDate >= '" & m_dBillFrom & "' AND"
	end if
	if Request.Form("BillThru") <> "" then
		m_dBillThru = Request.Form("BillThru")
		sSQL = sSQL & " c.BillDate <= '" & dateadd("s",-1,dateadd("d",1,m_dBillThru)) & "' AND"
	end if
	if Request.Form("Confine") <> "" then
		m_iConfine = Request.Form("Confine")
		sSQL = sSQL & " t.Confine = " & m_iConfine & " AND"
	end if
	if Request.Form("ICD9_Proc") <> "" then
		m_iICD9_Proc = Request.Form("ICD9_Proc")
		sSQL = sSQL & " (h.Procedure_1 ='" & m_iICD9_Proc & "' OR h.Procedure_2 ='" & m_iICD9_Proc & "' OR h.Procedure_3 ='" & m_iICD9_Proc & "')" & " AND"
	end if
	if Request.Form("DRG") <> "" then
		m_iDRG = Request.Form("DRG")
		sSQL = sSQL & " h.DRG = " & m_iDRG & " AND"
	end if
	if Request.Form("RevCode") <> "" then
		m_iRevCode = Request.Form("RevCode")
		sSQL = sSQL & " g.RevenueCode = " & m_iRevCode & " AND"
	end if
	if Request.Form("DiagCode") <> "" then
		m_sDiagCode = Request.Form("DiagCode")
		sSQL = sSQL & " d.DiagnosisCode = '" & m_sDiagCode & "' AND"
	end if
	if Request.Form("ProcCode") <> "" then
		m_sProcCode = Request.Form("ProcCode")
		sSQL = sSQL & " g.ProcedureCode = '" & m_sProcCode & "' AND"
	end if
	if Request.Form("Modifier") <> "" then
		m_sModifier = Request.Form("Modifier")
		sSQL = sSQL & " g.Modifier = '" & m_sModifier & "' AND"
	end if
	if Request.Form("TrackCode") <> "" then
		m_sTrackCode = Request.Form("TrackCode")
		sSQL = sSQL & " g.TrackCode ='" & m_sTrackCode & "' AND"
	end if
	if Request.Form("CovCode") <> "" then
		m_sCovCode = Request.Form("CovCode")
		sSQL = sSQL & " g.CovCode = '" & m_sCovCode & "' AND"
	end if
	if Request.Form("POSCode") <> "" then
		m_iPOSCode = Request.Form("POSCode")
		sSQL = sSQL & " g.PlaceOfService =" & m_iPOSCode & " AND"
	end if
	if Request.Form("PayCode") <> "" then
		m_sPayCode = Request.Form("PayCode")
		sSQL = sSQL & " c.PayCode = '" & m_sPayCode & "' AND"
	end if
	if Request.Form("Charge") <> "" then
		m_cCharge = Request.Form("Charge")
		sSQL = sSQL & " g.ChargeAmount =" & m_cCharge & " AND"
	end if
	if Request.Form("CheckAcct") <> "" then
		m_sCheckAcct = Request.Form("CheckAcct")
		sSQL = sSQL & " g.CheckAccount ='" & m_sCheckAcct & "' AND"
	end if
	if Request.Form("CheckNum") <> "" then
		m_iCheckNum = Request.Form("CheckNum")
		sSQL = sSQL & " g.CheckNumber =" & m_iCheckNum & " AND"
	end if
	if Request.Form("PPO") <> "" then
		m_sPPO = Request.Form("PPO")
		sSQL = sSQL & " c.PPOSponsor ='" & m_sPPO & "' AND"
	end if
	if Request.Form("PPOOverride") <> "" then
		m_bPPOCheck = true
		sSQL = sSQL & " c.PPOOverride = 1 AND"
	else
		m_bPPOCheck = false
	end if
	if Request.Form("OverPay") <> "" then
		m_bOverPayCheck = true
		sSQL = sSQL & " c.OverpaymentAdjustment <> 0 AND"
	else
		m_bOverPayCheck = false
	end if
	if Request.Form("NonOverPay") <> "" then
		m_bOtherAdjustCheck = true
		sSQL = sSQL & " c.OtherAdjustment <> 0 AND"
	else
		m_bOtherAdjustCheck = false
	end if
	if Request.Form("PreCalCode") <> "" then
		m_sPreCalCode = Request.Form("PreCalCode")
		sSQL = sSQL & " r.PreCalCode = '" & m_sPreCalCode & "' AND"
	end if
	if Request.Form("COB") <> "" then
		m_bCOBCheck = true
		sSQL = sSQL & " c.COBIndicator in ('Y','N','2','3') AND"
	else
		m_bCOBCheck = false
	end if
	if Request.Form("COBOver") <> "" then
		m_bCOBOverCheck = true
		sSQL = sSQL & " o.ErrorNumber=" & m_iCOBOverride & " AND"
	else
		m_bCOBOverCheck = false
	end if

	sSQL = left(sSQL,len(sSQL)-4)
	if m_bShowChargelines then
		sSQL = sSQL & " order by c.LogDate DESC, c.Sequence, c.Distr, g.LineNumber"
	else
		sTemp = "select distinct cg.LogDate, cg.Sequence, cg.Distr,"
		sTemp = sTemp & " sum(cg.Charge) 'Total Charges',"
		sTemp = sTemp & " cg.SSN, cg.DepNumber 'Dep #', cg.Name, cg.Provider, cg.TaxID, "
		sTemp = sTemp & "cg.Associate 'Assoc #' from ("
		sTemp = sTemp & sSQL
		sTemp = sTemp & ") cg"
		sTemp = sTemp & " group by cg.LogDate, cg.Sequence, cg.Distr, cg.SSN,"   
		sTemp = sTemp & " cg.DepNumber, cg.Name, cg.Provider, cg.TaxID, cg.Associate"
		sSQL = sTemp & " order by cg.LogDate DESC, cg.Sequence, cg.Distr"
	end if

end if
%>
<html>
<head>
    <title>Details</title>
</head>
<body topmargin="2" leftmargin="2" rightmargin="0" background="images/stoneback.jpg"
    language="VBScript" onload="UpdateScreen(3)">
    <link rel="STYLESHEET" href="styles/CritTable.css">
    <table width="100%" cols="3">
        <tr>
            <td align="CENTER" width="10%">
                <%
			if not Session("IsClerk") then
                %>
                <img src="images/log.gif" onclick="LogCall()">
                <%
			end if
                %>
            </td>
            <td align="CENTER">
                <font size="+2"><strong>Claim Search
                    <%
			if bUseSQL then
				if m_bShowChargelines then
					sTemp = " (Chargelines)"
				else
					sTemp = " (Claims)"
				end if
				Response.Write sTemp
			end if
                    %>
                </strong></font>
            </td>
            <td align="CENTER" width="10%">
                <img src="images/back.gif" onclick="history.go(-1)">
            </td>
        </tr>
    </table>
    <table cols="2" width="100%" cellpadding="0" cellspacing="0">
        <tr>
            <td width="60%">
                <input type="BUTTON" name="NewQuery" value="Claims" onclick="ValidSend(0)">
                <input type="BUTTON" name="NewQuery" value="Chargelines" onclick="ValidSend(-1)">
                <input type="BUTTON" name="ClearCrits" value="Clear Query" onclick="ValidSend(-2)">
            </td>
            <td align="RIGHT" name="MorePages" id="MorePages">
                <%
			if bUseSQL then	
				adoRS.Open sSql, adoConn, adOpenKeyset, adLockOptimistic
				if adoRS.EOF then
					adoRS.Close
					adoConn.Close
					set adoRS = nothing
					set adoConn = nothing
					sNoMatches = "No matches were found -- try a different search."
					bUseSQL = false
				else
					adoRS.PageSize = m_iPageSize ' Number of rows per page
					iPageCount = adoRS.PageCount
					adoRS.AbsolutePage = iPageNo
					if iPageCount > 1 then
						sTemp = "<FONT SIZE=+1>Page " & iPageNo & " of " & iPageCount & "</FONT>&nbsp;&nbsp;"
						if iPageNo > 1 Then
							sTemp = sTemp & "<INPUT ALIGN=RIGHT TYPE=BUTTON NAME=ScrollAction VALUE='Page " & iPageNo-1 & "' onClick='ValidSend(" & iPageNo-1 & ")'>"
						else
							sTemp = sTemp & "<INPUT STYLE='visibility:hidden;' ALIGN=RIGHT TYPE=BUTTON VALUE='Page " & iPageNo-1 & "' >"
						end if
						if iPageNo < iPageCount Then
							sTemp = sTemp & "<INPUT ALIGN=RIGHT TYPE=BUTTON NAME=ScrollAction VALUE='Page " & iPageNo+1 & "' onClick='ValidSend(" & iPageNo+1 & ")'>"
						else
							sTemp = sTemp & "<INPUT STYLE='visibility:hidden;' ALIGN=RIGHT TYPE=BUTTON VALUE='Page " & iPageNo+1 & "' >"
						end if
						Response.Write sTemp
					end if
				end if
			end if
                %>
            </td>
        </tr>
    </table>
    <div style="height=33%; width: 100%; overflow=auto">
        <form id="Criteria" method="POST" action="ClaimSearch.asp" target="Details">
        <input type="HIDDEN" id="ActionType" name="ActionType" value>
        <input type="HIDDEN" id="ClaimType" name="ClaimType" value>
        <table class="CriteriaTable" cols="8" cellpadding="0" cellspacing="0">
            <tr>
                <td align="RIGHT">
                    Log Date From:
                </td>
                <td>
                    <input type="TEXT" id="FromDate" name="FromDate" size="10" maxlength="10" value="<%= m_dFromDate%>">
                </td>
                <td align="RIGHT">
                    Log Date Thru:
                </td>
                <td>
                    <input type="TEXT" id="ThruDate" name="ThruDate" size="10" maxlength="10" value="<%= m_dThruDate%>">
                </td>
                <td align="RIGHT">
                    Sequence:
                </td>
                <td>
                    <input type="TEXT" id="Sequence" name="Sequence" size="6" maxlength="10" value="<%= m_iSequence%>">
                </td>
                <td align="RIGHT">
                    Distr:
                </td>
                <td>
                    <input type="TEXT" id="Distr" name="Distr" size="3" maxlength="5" value="<%= m_iDistr%>">
                </td>
            </tr>
            <tr>
                <td align="RIGHT">
                    SSN:
                </td>
                <td>
                    <input type="TEXT" id="SSN" name="SSN" size="10" maxlength="10" value="<%= m_iSSN%>">
                </td>
                <td align="RIGHT">
                    Last Name:
                </td>
                <td>
                    <input type="TEXT" id="LName" name="LName" size="15" maxlength="20" value="<%= m_sLastName%>">
                </td>
                <td align="RIGHT">
                    Dep. #:
                </td>
                <td>
                    <input type="TEXT" id="DepNum" name="DepNum" size="3" maxlength="5" value="<%= m_iDepNum%>">
                </td>
                <td colspan="2">
                </td>
            </tr>
            <tr>
                <td align="RIGHT">
                    Provider Tax ID:
                </td>
                <td>
                    <input type="TEXT" id="TaxID" name="TaxID" size="10" maxlength="10" value="<%= m_iTaxID%>">
                </td>
                <td align="RIGHT">
                    Provider Name:
                </td>
                <td>
                    <input type="TEXT" id="ProvName" name="ProvName" size="15" maxlength="40" value="<%= m_sProvName%>">
                </td>
                <td align="RIGHT">
                    Assoc. #:
                </td>
                <td>
                    <input type="TEXT" id="AssocNum" name="AssocNum" size="3" maxlength="10" value="<%= m_iAssocNum%>">
                </td>
                <td colspan="2">
                </td>
            </tr>
            <tr>
                <td align="RIGHT">
                    Fund:
                </td>
                <td>
                    <input type="TEXT" id="Fund" name="Fund" size="5" maxlength="3" value="<%= m_sFund%>">
                </td>
                <td align="RIGHT">
                    Status:
                </td>
                <td>
                    <input type="TEXT" id="Status1" name="Status1" size="2" maxlength="1" value="<%= m_sStatus1%>">
                    <input type="TEXT" id="Status2" name="Status2" size="2" maxlength="1" value="<%= m_sStatus2%>">
                    <input type="TEXT" id="Status3" name="Status3" size="2" maxlength="1" value="<%= m_sStatus3%>">
                </td>
                <td align="RIGHT">
                    Status Date From:
                </td>
                <td>
                    <input type="TEXT" id="StatFrom" name="StatFrom" size="10" maxlength="10" value="<%= m_dStatFrom%>">
                </td>
                <td align="RIGHT">
                    Status Date Thru:
                </td>
                <td>
                    <input type="TEXT" id="StatThru" name="StatThru" size="10" maxlength="10" value="<%= m_dStatThru%>">
                </td>
            </tr>
            <tr>
                <td align="RIGHT">
                    Plan:
                </td>
                <td>
                    <input type="TEXT" id="Plan" name="Plan" size="8" maxlength="6" value="<%= m_sPlan%>">
                </td>
                <td align="RIGHT">
                    Pend/Deny Reason:
                </td>
                <td>
                    <input type="TEXT" id="RSN" name="RSN" size="5" maxlength="4" value="<%= m_sRSN%>">
                </td>
                <td align="RIGHT">
                    Claim Service Date From:
                </td>
                <td>
                    <input type="TEXT" id="ServiceFrom" name="ServiceFrom" size="10" maxlength="10" value="<%= m_dServiceFrom%>">
                </td>
                <td align="RIGHT">
                    Claim Service Date Thru:
                </td>
                <td>
                    <input type="TEXT" id="ServiceThru" name="ServiceThru" size="10" maxlength="10" value="<%= m_dServiceThru%>">
                </td>
            </tr>
            <tr>
                <td align="RIGHT">
                    HWType:
                </td>
                <td>
                    <input type="TEXT" id="HWType" name="HWType" size="3" maxlength="1" value="<%= m_sHWType%>">
                </td>
                <td align="RIGHT">
                    Benefit Period:
                </td>
                <td>
                    <input type="TEXT" id="BenPeriod" name="BenPeriod" size="5" maxlength="5" value="<%= m_iBenPeriod%>">
                </td>
                <td align="RIGHT">
                    Benefit Period Date From:
                </td>
                <td>
                    <input type="TEXT" id="BenFrom" name="BenFrom" size="10" maxlength="10" value="<%= m_dBenFrom%>">
                </td>
                <td align="RIGHT">
                    Benefit Period Date Thru:
                </td>
                <td>
                    <input type="TEXT" id="BenThru" name="BenThru" size="10" maxlength="10" value="<%= m_dBenThru%>">
                </td>
            </tr>
            <tr>
                <td align="RIGHT">
                    Admit Date From:
                </td>
                <td>
                    <input type="TEXT" id="AdmitFrom" name="AdmitFrom" size="10" maxlength="10" value="<%= m_dAdmitFrom%>">
                </td>
                <td align="RIGHT">
                    Admit Date Thru:
                </td>
                <td>
                    <input type="TEXT" id="AdmitThru" name="AdmitThru" size="10" maxlength="10" value="<%= m_dAdmitThru%>">
                </td>
                <td align="RIGHT">
                    Discharge Date From:
                </td>
                <td>
                    <input type="TEXT" id="DisFrom" name="DisFrom" size="10" maxlength="10" value="<%= m_dDisFrom%>">
                </td>
                <td align="RIGHT">
                    Discharge Date Thru:
                </td>
                <td>
                    <input type="TEXT" id="DisThru" name="DisThru" size="10" maxlength="10" value="<%= m_dDisThru%>">
                </td>
            </tr>
            <tr>
                <td align="RIGHT">
                    Admit Type:
                </td>
                <td>
                    <input type="TEXT" id="AdmitType" name="AdmitType" size="3" maxlength="1" value="<%= m_sAdmitType%>">
                </td>
                <td align="RIGHT">
                    Bill Type:
                </td>
                <td>
                    <input type="TEXT" id="BillType" name="BillType" size="5" maxlength="5" value="<%= m_iBillType%>">
                </td>
                <td align="RIGHT">
                    Bill Date From:
                </td>
                <td>
                    <input type="TEXT" id="BillFrom" name="BillFrom" size="10" maxlength="10" value="<%= m_dBillFrom%>">
                </td>
                <td align="RIGHT">
                    Bill Date Thru:
                </td>
                <td>
                    <input type="TEXT" id="BillThru" name="BillThru" size="10" maxlength="10" value="<%= m_dBillThru%>">
                </td>
            </tr>
            <tr>
                <td align="RIGHT">
                    Confine Type:
                </td>
                <td>
                    <input type="TEXT" id="Confine" name="Confine" size="5" maxlength="5" value="<%= m_iConfine%>">
                </td>
                <td align="RIGHT">
                    ICD9 Proc. Code:
                </td>
                <td>
                    <input type="TEXT" id="ICD9_Proc" name="ICD9_Proc" size="6" maxlength="5" value="<%= m_iICD9_Proc%>">
                </td>
                <td align="RIGHT">
                    DRG:
                </td>
                <td>
                    <input type="TEXT" id="DRG" name="DRG" size="5" maxlength="5" value="<%= m_iDRG%>">
                </td>
                <td align="RIGHT">
                    Revenue Code:
                </td>
                <td>
                    <input type="TEXT" id="RevCode" name="RevCode" size="5" maxlength="5" value="<%= m_iRevCode%>">
                </td>
            </tr>
            <tr>
                <td align="RIGHT">
                    Diagnosis:
                </td>
                <td>
                    <input type="TEXT" id="DiagCode" name="DiagCode" size="8" maxlength="6" value="<%= m_sDiagCode%>">
                </td>
                <td align="RIGHT">
                    Procedure Code:
                </td>
                <td>
                    <input type="TEXT" id="ProcCode" name="ProcCode" size="7" maxlength="5" value="<%= m_sProcCode%>">
                </td>
                <td align="RIGHT">
                    Modifier:
                </td>
                <td>
                    <input type="TEXT" id="Modifier" name="Modifier" size="4" maxlength="2" value="<%= m_sModifier%>">
                </td>
                <td align="RIGHT">
                    Track Code:
                </td>
                <td>
                    <input type="TEXT" id="TrackCode" name="TrackCode" size="5" maxlength="3" value="<%= m_sTrackCode%>">
                </td>
            </tr>
            <tr>
                <td align="RIGHT">
                    Coverage Code:
                </td>
                <td>
                    <input type="TEXT" id="CovCode" name="CovCode" size="2" maxlength="2" value="<%= m_sCovCode%>">
                </td>
                <td align="RIGHT">
                    Pre Cal Code:
                </td>
                <td>
                    <input type="TEXT" id="PreCalCode" name="PreCalCode" size="4" maxlength="2" value="<%= m_sPreCalCode%>">
                </td>
                <td align="RIGHT">
                    Place of Service Code:
                </td>
                <td>
                    <input type="TEXT" id="POSCode" name="POSCode" size="5" maxlength="5" value="<%= m_iPOSCode%>">
                </td>
                <td align="RIGHT">
                    Pay To Code:
                </td>
                <td>
                    <input type="TEXT" id="PayCode" name="PayCode" size="3" maxlength="1" value="<%= m_sPayCode%>">
                </td>
            </tr>
            <tr>
                <td align="RIGHT">
                    Charge:
                </td>
                <td>
                    <input type="TEXT" id="Charge" name="Charge" size="8" maxlength="11" value="<%= m_cCharge%>">
                </td>
                <td align="RIGHT">
                    Check Account:
                </td>
                <td>
                    <input type="TEXT" id="CheckAcct" name="CheckAcct" size="8" maxlength="6" value="<%= m_sCheckAcct%>">
                </td>
                <td align="RIGHT">
                    Check Number:
                </td>
                <td>
                    <input type="TEXT" id="CheckNum" name="CheckNum" size="7" maxlength="10" value="<%= m_iCheckNum%>">
                </td>
                <td colspan="2">
                </td>
            </tr>
            <tr>
            </tr>
            <tr>
                <td align="RIGHT">
                    PPO Code:
                </td>
                <td>
                    <input type="TEXT" id="PPO" name="PPO" size="5" maxlength="3" value="<%= m_sPPO%>">
                </td>
                <td align="RIGHT">
                    PPO Override:
                </td>
                <td>
                    <input type="CHECKBOX" id="PPOOverride" name="PPOOverride" <%
							if m_bPPOCheck then
								Response.Write " CHECKED "
							end if
%>>
                </td>
                <td align="RIGHT">
                    COB:
                </td>
                <td>
                    <input type="CHECKBOX" id="COB" name="COB" <%
							if m_bCOBCheck then
								Response.Write " CHECKED "
							end if
%>>
                </td>
                <td align="RIGHT">
                    COB Override:
                </td>
                <td>
                    <input type="CHECKBOX" id="COBOver" name="COBOver" <%
							if m_bCOBOverCheck then
								Response.Write " CHECKED "
							end if
%>>
                </td>
            </tr>
            <tr>
                <td align="RIGHT">
                    Overpayment Adjustments:
                </td>
                <td>
                    <input type="CHECKBOX" id="OverPay" name="OverPay" <%
							if m_bOverPayCheck then
								Response.Write " CHECKED "
							end if
%>>
                </td>
                <td align="RIGHT">
                    Non-Overpayment Adjustments:
                </td>
                <td>
                    <input type="CHECKBOX" id="NonOverPay" name="NonOverPay" <%
							if m_bOtherAdjustCheck then
								Response.Write " CHECKED "
							end if
%>>
                </td>
                <td colspan="4">
                </td>
            </tr>
        </table>
        </form>
    </div>
    <br>
    <%
if bUseSQL then
    %>
    <div style="height: 45%; width: 100%; overflow: auto">
        <table width="100%" border="1">
            <tr>
                <td align="CENTER">
                    <font color="BLUE"><b>Claim #</b></font>
                </td>
                <%
			For i = 3 to adoRS.Fields.Count - 1 
                %>
                <td align="CENTER" nowrap>
                    <font color="BLUE"><b>
                        <%=adoRS(i).Name %></b></font>
                </td>
                <%
			Next
                %>
            </tr>
            <% 
			iRowCount = adoRS.PageSize
			if m_bShowChargelines then
				sTemp = "ClaimDetails.asp"
			else
				sTemp = "ClaimDetails.asp"
			end if
			Do While Not adoRS.EOF and iRowCount > 0
            %>
            <tr>
                <td nowrap>
                    <a href="<% = sTemp %>?LogDate=<% = adoRS("LogDate") %>&amp;Sequence=<%= adoRS("Sequence") %>&amp;Distr=<%=  adoRS("Distr") %>"
                        onclick="WorkingStatus()">
                        <% 
						Response.Write ConvertToClaimNum(adoRS("LogDate"),adoRS("Sequence"),adoRS("Distr"))
                        %>
                    </a>
                </td>
                <%
					For i = 3 to adoRS.Fields.Count - 1 
						if adoRS(i).type = adCurrency then
							if isnull(adoRS(i)) then
								Response.Write "<TD></TD>"
							else
								Response.Write "<td ALIGN=RIGHT NOWRAP>" & formatcurrency(adoRS(i)) & "</td>"
							end if
						else 
							Response.Write "<td NOWRAP>" & adoRS(i) & "</td>"
						end if
					Next
                %>
            </tr>
            <%
				iRowCount = iRowCount - 1
				adoRS.MoveNext
			Loop
			adoRS.Close
			adoConn.Close
			set adoRS = nothing
			set adoConn = nothing
            %>
        </table>
    </div>
    <%
else
	if sNoMatches <> "" then
		Response.Write sNoMatches
	end if
end if
    %>
    <div style="display: none">
        <form id="ClearCrits" method="POST" action="ClaimSearch.asp" target="Details">
        <input type="HIDDEN" id="UserCriteria" name="UserCriteria" value>
        </form>
    </div>
</body>
<!--#include file="VBFuncs.inc" -->
<script language="VBSCRIPT">

	dim iCurrDiv
	iCurrDiv=0
	
	sub ValidSend(iIndex)
	dim bValidData
	
		if iIndex = -2 then
			WorkingStatus
			ClearCrits.UserCriteria.value="Initial"
			ClearCrits.submit
			exit sub
		end if

		Criteria.FromDate.Value = trim(Criteria.FromDate.Value)
		Criteria.ThruDate.Value = trim(Criteria.ThruDate.Value)
		Criteria.Sequence.Value = trim(Criteria.Sequence.Value)
		Criteria.Distr.Value = trim(Criteria.Distr.Value)
		Criteria.SSN.Value = trim(Criteria.SSN.Value)
		Criteria.LName.Value = trim(Criteria.LName.Value)
		Criteria.DepNum.Value = trim(Criteria.DepNum.Value)
		Criteria.TaxID.Value = trim(Criteria.TaxID.Value)
		Criteria.ProvName.Value = trim(Criteria.ProvName.Value)
		Criteria.AssocNum.Value = trim(Criteria.AssocNum.Value)
		Criteria.Fund.Value = trim(Criteria.Fund.Value)
		Criteria.Status1.Value = trim(Criteria.Status1.Value)
		Criteria.Status2.Value = trim(Criteria.Status2.Value)
		Criteria.Status3.Value = trim(Criteria.Status3.Value)
		Criteria.StatFrom.Value = trim(Criteria.StatFrom.Value)
		Criteria.StatThru.Value = trim(Criteria.StatThru.Value)
		Criteria.Plan.Value = trim(Criteria.Plan.Value)
		Criteria.RSN.Value = trim(Criteria.RSN.Value)
		Criteria.ServiceFrom.Value = trim(Criteria.ServiceFrom.Value)
		Criteria.ServiceThru.Value = trim(Criteria.ServiceThru.Value)
		Criteria.HWType.Value = trim(Criteria.HWType.Value)
		Criteria.BenPeriod.Value = trim(Criteria.BenPeriod.Value)
		Criteria.BenFrom.Value = trim(Criteria.BenFrom.Value)
		Criteria.BenThru.Value = trim(Criteria.BenThru.Value)
		Criteria.AdmitFrom.Value = trim(Criteria.AdmitFrom.Value)
		Criteria.AdmitThru.Value = trim(Criteria.AdmitThru.Value)
		Criteria.DisFrom.Value = trim(Criteria.DisFrom.Value)
		Criteria.DisThru.Value = trim(Criteria.DisThru.Value)
		Criteria.AdmitType.Value = trim(Criteria.AdmitType.Value)
		Criteria.BillType.Value = trim(Criteria.BillType.Value)
		Criteria.BillFrom.Value = trim(Criteria.BillFrom.Value)
		Criteria.BillThru.Value = trim(Criteria.BillThru.Value)
		Criteria.Confine.Value = trim(Criteria.Confine.Value)
		Criteria.ICD9_Proc.Value = trim(Criteria.ICD9_Proc.Value)
		Criteria.DRG.Value = trim(Criteria.DRG.Value)
		Criteria.RevCode.Value = trim(Criteria.RevCode.Value)
		Criteria.DiagCode.Value = trim(Criteria.DiagCode.Value)
		Criteria.ProcCode.Value = trim(Criteria.ProcCode.Value)
		Criteria.Modifier.Value = trim(Criteria.Modifier.Value)
		Criteria.TrackCode.Value = trim(Criteria.TrackCode.Value)
		Criteria.CovCode.Value = trim(Criteria.CovCode.Value)
		Criteria.PreCalCode.Value = trim(Criteria.PreCalCode.Value)
		Criteria.POSCode.Value = trim(Criteria.POSCode.Value)
		Criteria.PayCode.Value = trim(Criteria.PayCode.Value)
		Criteria.Charge.Value = trim(Criteria.Charge.Value)
		Criteria.CheckAcct.Value = trim(Criteria.CheckAcct.Value)
		Criteria.CheckNum.Value = trim(Criteria.CheckNum.Value)
		Criteria.PPO.Value = trim(Criteria.PPO.Value)
'
' Primary
'
		bValidData = false
		if Criteria.FromDate.Value <> "" then
			if not isDate(Criteria.FromDate.Value) then
				if len(Criteria.FromDate.Value) = 6 then
					sTemp = mid(Criteria.FromDate.Value,1,2) & "/" & mid(Criteria.FromDate.Value,3,2) & "/" & mid(Criteria.FromDate.Value,5,2) 
					if isDate(sTemp) then
						Criteria.FromDate.Value = sTemp
					else
						msgbox "Please enter a valid date for the Log Date From field.",,"Invalid Data"
						exit sub
					end if
				else
					msgbox "Please enter a valid date for the Log Date From field.",,"Invalid Data"
					exit sub
				end if
			end if
			bValidData = true
		end if
		if Criteria.ThruDate.Value <> "" then
			if not isDate(Criteria.ThruDate.Value) then
				if len(Criteria.ThruDate.Value) = 6 then
					sTemp = mid(Criteria.ThruDate.Value,1,2) & "/" & mid(Criteria.ThruDate.Value,3,2) & "/" & mid(Criteria.ThruDate.Value,5,2) 
					if isDate(sTemp) then
						Criteria.ThruDate.Value = sTemp
					else
						msgbox "Please enter a valid date for the Log Date Thru field.",,"Invalid Data"
						exit sub
					end if
				else
					msgbox "Please enter a valid date for the Log Date Thru field.",,"Invalid Data"
					exit sub
				end if
			end if
			bValidData = true
		end if
		if Criteria.Sequence.Value <> ""  then
			if not isNumeric(Criteria.Sequence.Value) then
				msgbox "Please enter a valid Sequence number.",,"Invalid Data"
				exit sub
			else
				if Criteria.Sequence.Value > <%= Application("IntMax")%> or _
						Criteria.Sequence.Value < <%= Application("IntMin")%> then
					msgbox "Please enter a valid Sequence number.",,"Invalid Data"
					exit sub
				end if
			end if
			bValidData = true
		end if
		if Criteria.Distr.Value <> ""  then
			if not isNumeric(Criteria.Distr.Value) then
				msgbox "Please enter a valid Distr Number.",,"Invalid Data"
				exit sub
			else
				if Criteria.Distr.Value > <%= Application("SmlIntMax")%> or _
						Criteria.Distr.Value < <%= Application("SmlIntMin")%> then
					msgbox "Please enter a valid Distr Number.",,"Invalid Data"
					exit sub
				end if
			end if
			bValidData = true
		end if
		if Criteria.SSN.Value <> ""  then
			if not isNumeric(Criteria.SSN.Value) then
				msgbox "Please enter a valid SSN.",,"Invalid Data"
				exit sub
			else
				if Criteria.SSN.Value > <%= Application("IntMax")%> or _
						Criteria.SSN.Value < <%= Application("IntMin")%> then
					msgbox "Please enter a valid SSN.",,"Invalid Data"
					exit sub
				end if
			end if
			bValidData = true
		end if
		if Criteria.LName.Value <> "" then
			if ContainsInvalids(Criteria.LName.Value) then
				msgbox "Please remove invalid characters from the Last Name field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.DepNum.Value <> ""  then
			if not isNumeric(Criteria.DepNum.Value) then
				msgbox "Please enter a valid Dependent Number.",,"Invalid Data"
				exit sub
			else
				if Criteria.DepNum.Value > <%= Application("SmlIntMax")%> or _
						Criteria.DepNum.Value < <%= Application("SmlIntMin")%> then
					msgbox "Please enter a valid Dependent Number.",,"Invalid Data"
					exit sub
				end if
			end if
			bValidData = true
		end if
		if Criteria.TaxID.Value <> ""  then
			if not isNumeric(Criteria.TaxID.Value) then
				msgbox "Please enter a valid Tax ID.",,"Invalid Data"
				exit sub
			else
				if Criteria.TaxID.Value > <%= Application("IntMax")%> or _
						Criteria.TaxID.Value < <%= Application("IntMin")%> then
					msgbox "Please enter a valid Tax ID.",,"Invalid Data"
					exit sub
				end if
			end if
			bValidData = true
		end if
		if Criteria.ProvName.Value <> "" then
			if ContainsInvalids(Criteria.ProvName.Value) then
				msgbox "Please remove invalid characters from the Provider Name field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.AssocNum.Value <> ""  then
			if not isNumeric(Criteria.AssocNum.Value) then
				msgbox "Please enter a valid Associate Number.",,"Invalid Data"
				exit sub
			else
				if Criteria.AssocNum.Value > <%= Application("IntMax")%> or _
						Criteria.AssocNum.Value < <%= Application("IntMin")%> then
					msgbox "Please enter a valid Associate Number.",,"Invalid Data"
					exit sub
				end if
			end if
			bValidData = true
		end if
'
' Secondary
'
		if Criteria.Fund.Value <> "" then
			if ContainsInvalids(Criteria.Fund.Value) then
				msgbox "Please remove invalid characters from the Fund field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.Status1.Value <> "" then
			if ContainsInvalids(Criteria.Status1.Value) then
				msgbox "Please remove invalid characters from the first Status field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.Status2.Value <> "" then
			if ContainsInvalids(Criteria.Status2.Value) then
				msgbox "Please remove invalid characters from the second Status field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.Status3.Value <> "" then
			if ContainsInvalids(Criteria.Status3.Value) then
				msgbox "Please remove invalid characters from the third Status field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.StatFrom.Value <> "" then
			if not isDate(Criteria.StatFrom.Value) then
				msgbox "Please enter a valid date for the Status Date From field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.StatThru.Value <> "" then
			if not isDate(Criteria.StatThru.Value) then
				msgbox "Please enter a valid date for the Status Date Thru field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.Plan.Value <> "" then
			if ContainsInvalids(Criteria.Plan.Value) then
				msgbox "Please remove invalid characters from the Plan field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.RSN.Value <> "" then
			if ContainsInvalids(Criteria.RSN.Value) then
				msgbox "Please remove invalid characters from the Pend/Deny Reason field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.ServiceFrom.Value <> "" then
			if not isDate(Criteria.ServiceFrom.Value) then
				msgbox "Please enter a valid date for the Service Date From field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.ServiceThru.Value <> "" then
			if not isDate(Criteria.ServiceThru.Value) then
				msgbox "Please enter a valid date for the Service Date Thru field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.HWType.Value <> "" then
			if ContainsInvalids(Criteria.HWType.Value) then
				msgbox "Please remove invalid characters from the HW Type field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.BenPeriod.Value <> ""  then
			if not isNumeric(Criteria.BenPeriod.Value) then
				msgbox "Please enter a valid Benefit Period.",,"Invalid Data"
				exit sub
			else
				if Criteria.BenPeriod.Value > <%= Application("SmlIntMax")%> or _
						Criteria.BenPeriod.Value < <%= Application("SmlIntMin")%> then
					msgbox "Please enter a valid Benefit Period.",,"Invalid Data"
					exit sub
				end if
			end if
			bValidData = true
		end if
		if Criteria.BenFrom.Value <> "" then
			if not isDate(Criteria.BenFrom.Value) then
				msgbox "Please enter a valid date for the Benefit Date From field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.BenThru.Value <> "" then
			if not isDate(Criteria.BenThru.Value) then
				msgbox "Please enter a valid date for the Benefit Date Thru field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
'
' Hospital
'
		if Criteria.AdmitFrom.Value <> "" then
			if not isDate(Criteria.AdmitFrom.Value) then
				msgbox "Please enter a valid date for the Admit Date From field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.AdmitThru.Value <> "" then
			if not isDate(Criteria.AdmitThru.Value) then
				msgbox "Please enter a valid date for the Admit Date Thru field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.DisFrom.Value <> "" then
			if not isDate(Criteria.DisFrom.Value) then
				msgbox "Please enter a valid date for the Discharge Date From field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.DisThru.Value <> "" then
			if not isDate(Criteria.DisThru.Value) then
				msgbox "Please enter a valid date for the Discharge Date Thru field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.AdmitType.Value <> "" then
			if ContainsInvalids(Criteria.AdmitType.Value) then
				msgbox "Please remove invalid characters from the Admit Type field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.BillType.Value <> ""  then
			if not isNumeric(Criteria.BillType.Value) then
				msgbox "Please enter a valid Bill Type.",,"Invalid Data"
				exit sub
			else
				if Criteria.BillType.Value > <%= Application("SmlIntMax")%> or _
						Criteria.BillType.Value < <%= Application("SmlIntMin")%> then
					msgbox "Please enter a valid Bill Type.",,"Invalid Data"
					exit sub
				end if
			end if
			bValidData = true
		end if
		if Criteria.BillFrom.Value <> "" then
			if not isDate(Criteria.BillFrom.Value) then
				msgbox "Please enter a valid date for the Bill Date From field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.BillThru.Value <> "" then
			if not isDate(Criteria.BillThru.Value) then
				msgbox "Please enter a valid date for the Bill Date Thru field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.Confine.Value <> ""  then
			if not isNumeric(Criteria.Confine.Value) then
				msgbox "Please enter a valid Confine value.",,"Invalid Data"
				exit sub
			else
				if Criteria.Confine.Value > <%= Application("SmlIntMax")%> or _
						Criteria.Confine.Value < <%= Application("SmlIntMin")%> then
					msgbox "Please enter a valid Confine value.",,"Invalid Data"
					exit sub
				end if
			end if
			bValidData = true
		end if
		if Criteria.ICD9_Proc.Value <> "" then
			if ContainsInvalids(Criteria.ICD9_Proc.Value) then
				msgbox "Please remove invalid characters from the ICD9 Procedure field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.DRG.Value <> ""  then
			if not isNumeric(Criteria.DRG.Value) then
				msgbox "Please enter a valid DRG value.",,"Invalid Data"
				exit sub
			else
				if Criteria.DRG.Value > <%= Application("SmlIntMax")%> or _
						Criteria.DRG.Value < <%= Application("SmlIntMin")%> then
					msgbox "Please enter a valid DRG value.",,"Invalid Data"
					exit sub
				end if
			end if
			bValidData = true
		end if
		if Criteria.RevCode.Value <> ""  then
			if not isNumeric(Criteria.RevCode.Value) then
				msgbox "Please enter a valid Revenue Code.",,"Invalid Data"
				exit sub
			else
				if Criteria.RevCode.Value > <%= Application("SmlIntMax")%> or _
						Criteria.RevCode.Value < <%= Application("SmlIntMin")%> then
					msgbox "Please enter a valid Revenue Code.",,"Invalid Data"
					exit sub
				end if
			end if
			bValidData = true
		end if
'
' Other
'
		if Criteria.DiagCode.Value <> "" then
			if ContainsInvalids(Criteria.DiagCode.Value) then
				msgbox "Please remove invalid characters from the Diagnosis field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.ProcCode.Value <> "" then
			if ContainsInvalids(Criteria.ProcCode.Value) then
				msgbox "Please remove invalid characters from the Procedure Code field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.Modifier.Value <> "" then
			if ContainsInvalids(Criteria.Modifier.Value) then
				msgbox "Please remove invalid characters from the Modifier field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.TrackCode.Value <> "" then
			if ContainsInvalids(Criteria.TrackCode.Value) then
				msgbox "Please remove invalid characters from the Track Code field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.CovCode.Value <> "" then
			if ContainsInvalids(Criteria.CovCode.Value) then
				msgbox "Please remove invalid characters from the Coverage Code field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.PreCalCode.Value <> "" then
			if ContainsInvalids(Criteria.PreCalCode.Value) then
				msgbox "Please remove invalid characters from the PreCal Code field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.POSCode.Value <> ""  then
			if not isNumeric(Criteria.POSCode.Value) then
				msgbox "Please enter a valid Place Of Service.",,"Invalid Data"
				exit sub
			else
				if Criteria.POSCode.Value > <%= Application("SmlIntMax")%> or _
						Criteria.POSCode.Value < <%= Application("SmlIntMin")%> then
					msgbox "Please enter a valid Place Of Service.",,"Invalid Data"
					exit sub
				end if
			end if
			bValidData = true
		end if
		if trim(Criteria.PayCode.Value) <> "" then
			if ContainsInvalids(Criteria.PayCode.Value) then
				msgbox "Please remove invalid characters from the Pay Code field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.Charge.Value <> ""  then
			if not isNumeric(Criteria.Charge.Value) then
				msgbox "Please enter a valid Charge Amount.",,"Invalid Data"
				exit sub
			else
				if Criteria.Charge.Value > <%= Application("SmlMnyMax")%> or _
						Criteria.Charge.Value < <%= Application("SmlMnyMin")%> then
					msgbox "Please enter a valid Charge Amount.",,"Invalid Data"
					exit sub
				end if
			end if
			bValidData = true
		end if
		if Criteria.CheckAcct.Value <> "" then
			if ContainsInvalids(Criteria.CheckAcct.Value) then
				msgbox "Please remove invalid characters from the Check Account field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.CheckNum.Value <> ""  then
			if not isNumeric(Criteria.CheckNum.Value) then
				msgbox "Please enter a valid Check Number.",,"Invalid Data"
				exit sub
			else
				if Criteria.CheckNum.Value > <%= Application("IntMax")%> or _
						Criteria.CheckNum.Value < <%= Application("IntMin")%> then
					msgbox "Please enter a valid Check Number.",,"Invalid Data"
					exit sub
				end if
			end if
			bValidData = true
		end if
		if Criteria.PPO.Value <> "" then
			if ContainsInvalids(Criteria.PPO.Value) then
				msgbox "Please remove invalid characters from the PPO Sponsor field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		
		if not bValidData then
			msgbox "You haven't entered anything to search for.",,"No Criteria"
			exit sub
		end if
		
		WorkingStatus

		select case iIndex
			case 0
				Criteria.ActionType.value=""
				Criteria.ClaimType.value="Claims"
				Criteria.submit
			case -1
				Criteria.ActionType.value=""
				Criteria.ClaimType.value="Chargelines"
				Criteria.submit
			case else
				if "<%= m_dFromDate%>" <> Criteria.FromDate.Value or _
						"<%= m_dThruDate%>" <> Criteria.ThruDate.Value or _
						"<%= m_iSequence%>" <> Criteria.Sequence.Value or _
						"<%= m_iDistr%>" <> Criteria.Distr.Value or _
						"<%= m_iSSN%>" <> Criteria.SSN.Value or _
						"<%= m_sLastName%>" <> Criteria.LName.Value or _
						"<%= m_iDepNum%>" <> Criteria.DepNum.Value or _
						"<%= m_iTaxID%>" <> Criteria.TaxID.Value or _
						"<%= m_sProvName%>" <> Criteria.ProvName.Value or _
						"<%= m_iAssocNum%>" <> Criteria.AssocNum.Value or _
						"<%= m_sFund%>" <> Criteria.Fund.Value or _
						"<%= m_sStatus1%>" <> Criteria.Status1.Value or _
						"<%= m_sStatus2%>" <> Criteria.Status2.Value or _
						"<%= m_sStatus3%>" <> Criteria.Status3.Value or _
						"<%= m_dStatFrom%>" <> Criteria.StatFrom.Value or _
						"<%= m_dStatThru%>" <> Criteria.StatThru.Value or _
						"<%= m_sPlan%>" <> Criteria.Plan.Value or _
						"<%= m_sRSN%>" <> Criteria.RSN.Value or _
						"<%= m_dServiceFrom%>" <> Criteria.ServiceFrom.Value or _
						"<%= m_dServiceThru%>" <> Criteria.ServiceThru.Value or _
						"<%= m_sHWType%>" <> Criteria.HWType.Value or _
						"<%= m_iBenPeriod%>" <> Criteria.BenPeriod.Value or _
						"<%= m_dBenFrom%>" <> Criteria.BenFrom.Value or _
						"<%= m_dBenThru%>" <> Criteria.BenThru.Value or _
						"<%= m_dAdmitFrom%>" <> Criteria.AdmitFrom.Value or _
						"<%= m_dAdmitThru%>" <> Criteria.AdmitThru.Value or _
						"<%= m_dDisFrom%>" <> Criteria.DisFrom.Value or _
						"<%= m_dDisThru%>" <> Criteria.DisThru.Value or _
						"<%= m_sAdmitType%>" <> Criteria.AdmitType.Value or _
						"<%= m_iBillType%>" <> Criteria.BillType.Value or _
						"<%= m_dBillFrom%>" <> Criteria.BillFrom.Value or _
						"<%= m_dBillThru%>" <> Criteria.BillThru.Value or _
						"<%= m_iConfine%>" <> Criteria.Confine.Value or _
						"<%= m_iICD9_Proc%>" <> Criteria.ICD9_Proc.Value or _
						"<%= m_iDRG%>" <> Criteria.DRG.Value or _
						"<%= m_iRevCode%>" <> Criteria.RevCode.Value or _
						"<%= m_sDiagCode%>" <> Criteria.DiagCode.Value or _
						"<%= m_sProcCode%>" <> Criteria.ProcCode.Value or _
						"<%= m_sModifier%>" <> Criteria.Modifier.Value or _
						"<%= m_sTrackCode%>" <> Criteria.TrackCode.Value or _
						"<%= m_sCovCode%>" <> Criteria.CovCode.Value or _
						"<%= m_sPreCalCode%>" <> Criteria.PreCalCode.Value or _
						"<%= m_iPOSCode%>" <> Criteria.POSCode.Value or _
						"<%= m_sPayCode%>" <> Criteria.PayCode.Value or _
						"<%= m_cCharge%>" <> Criteria.Charge.Value or _
						"<%= m_sCheckAcct%>" <> Criteria.CheckAcct.Value or _
						"<%= m_iCheckNum%>" <> Criteria.CheckNum.Value or _
						"<%= m_sPPO%>" <> Criteria.PPO.Value or _
						"<%= m_bPPOCheck%>" <> Criteria.PPOOverride.Checked or _
						"<%= m_bCOBCheck%>" <> Criteria.COB.Checked or _
						"<%= m_bOverPayCheck%>" <> Criteria.OverPay.Checked or _
						"<%= m_bOtherAdjustCheck%>" <> Criteria.NonOverPay.Checked or _
						"<%= m_bCOBOverCheck%>" <> Criteria.COBOver.Checked then
					Criteria.ActionType.value=""
				else
					Criteria.ActionType.value=iIndex
				end if
				if <%= m_bShowChargelines%> then
					Criteria.ClaimType.value="Chargelines"
				else
					Criteria.ClaimType.value="Claims"
				end if
				Criteria.submit
		end select
	end sub
	
	sub SwitchCritTabs()
	
		if iCurrDiv <> window.event.srcElement.myTag then
			if iCurrDiv<> -1 then
				document.all.item("CritDiv",iCurrDiv).style.display="none"
				document.all.item("Search",iCurrDiv).background=document.all.item("Search",iCurrDiv).myDis
				document.all.item("Search",iCurrDiv).style.fontsize="8pt"
			end if
			iCurrDiv=window.event.srcElement.myTag
			document.all.item("CritDiv",iCurrDiv).style.display="block"
			document.all.item("Search",iCurrDiv).background=document.all.item("Search",iCurrDiv).myNorm
			document.all.item("Search",iCurrDiv).style.fontsize="10pt"
		end if

	end sub
</script>
</html>
