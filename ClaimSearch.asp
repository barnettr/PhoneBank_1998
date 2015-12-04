<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<% Response.Buffer = True %>
<% Response.Expires = 0 %>
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
dim bUseSQL, sNoMatches, m_bShowChargelines, adoCmd, m_iOVP
dim m_iCOBOverride, ClaimNum, Con2, qryGetName, PendDeny, Status, OVP, CheckCOB, adoParam1, adoParam2
dim adoParam3, adoParam4, adoParam5, adoParam6, adoParam7, adoParam8, adoParam9, adoParam10
dim adoParam11, adoParam12, adoParam13, adoParam14, adoParam15, adoParam16, adoParam17, adoParam18
dim adoParam19, adoParam20, adoParam21, adoParam22, adoParam23, adoParam24, adoParam25, adoParam26
dim Manager, Supervisor, Auditor, locktype
dim M, D, T, C, P, V


Set adoConn = Server.CreateObject("ADODB.Connection")
adoConn.ConnectionTimeout = 300
adoConn.CommandTimeout = 300
adoConn.Open Application("DataConn")
'adoConn.Open = "driver={SQL Server};server=seasql01;database=claimstest;uid=;pwd=;Trusted_Connection=Yes;DSN="

Set adoCmd = Server.CreateObject("ADODB.Command")
adoCmd.CommandTimeout = 300
adoCmd.ActiveConnection = adoConn

adoCmd.CommandType = &H0001
adoCmd.CommandText = "select * from UserInformation where LogonID='" & Session("User") & "'"

Set adoRS = Server.CreateObject("ADODB.Recordset")
adoRS.Open adoCmd,,3,3

If adoRS.state = 1 Then
	If adoRS.EOF Then
		norecords = 1
	End If
End If

if not adoRS.EOF then
  Manager = adoRS("IsManager")
  Supervisor = adoRS("IsSupervisor")
  Auditor = adoRS("IsAuditor")
end if
adoRS.Close

                        '****************************************************************************************************
m_iCOBOverride = 100021 'Per Russell, this code (in ErrorNumbers.h) indicates a Claim has a "COB Override" associated with it.
                        '****************************************************************************************************

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
	sClaimNum = sClaimNum & "_" & string(5-len(Sequence),"0") & Sequence
	sClaimNum = sClaimNum & "_" & Distr
	ConvertToClaimNum = sClaimNum
	
end function



m_bShowChargelines = false
if Request.Querystring("UserCriteria") = "Initial" then
	bUseSQL = false
else
	bUseSQL = true
	if Request.Querystring("ClaimType") <> "Claims" then
		m_bShowChargelines = true
	end if

  set adoCmd = Server.CreateObject("ADODB.Command")
  adoCmd.CommandTimeout = 300
  set adoRS = Server.CreateObject("ADODB.Recordset")
  

	if Trim(Request.Querystring("ActionType")) = "0" then
		iPageNo = 1
	else
		if Trim(Request.Querystring("ActionType")) = "" then
			iPageNo = 1
		else
			iPageNo = Cint(Request.Querystring("ActionType"))
			if iPageNo < 1 Then 
				iPageNo = 1
			end if
		end if
	end if
  	if Request.Querystring("FromDate") <> "" then
		m_dFromDate = Request.Querystring("FromDate")
	end if
	if Request.Querystring("ThruDate") <> "" then
		m_dThruDate = Request.Querystring("ThruDate")
	end if
	if Request.Querystring("Sequence") <> "" then
		m_iSequence = Request.Querystring("Sequence")
	end if
	if Request.Querystring("Distr") <> "" then
		m_iDistr = Request.Querystring("Distr")
	end if  
	if Request.Querystring("SSN") <> "" then
		m_iSSN = Request.Querystring("SSN")
	end if
	if Request.Querystring("LName") <> "" then
		m_sLastName = Request.Querystring("LName")
	end if
	if Request.Querystring("DepNum") <> "" then
		m_iDepNum = Request.Querystring("DepNum")
	end if
	if Request.Querystring("TaxID") <> "" then
		m_iTaxID = Request.Querystring("TaxID")
	end if
	if Request.Querystring("ProvName") <> "" then
		m_sProvName = Request.Querystring("ProvName")
	end if
	if Request.Querystring("AssocNum") <> "" then
		m_iAssocNum = Request.Querystring("AssocNum")
	end if
	if Request.Querystring("Fund") <> "" then
		m_sFund = Request.Querystring("Fund")
	end if
	if Request.Querystring("Plan") <> "" then
		m_sPlan = Request.Querystring("Plan")
	end if
	if Request.Querystring("HWType") <> "" then
		m_sHWType = Request.Querystring("HWType")
	end if
	if Request.Querystring("Status1") <> "" or Request.Querystring("Status2") <> "" or Request.Querystring("Status3") <> ""then
		if Request.Querystring("Status1") <> "" then
			m_sStatus1 = Request.Querystring("Status1")
		end if
		if Request.Querystring("Status2") <> "" then
			m_sStatus2 = Request.Querystring("Status2")
			if right(sSQL,1) = "(" then
			else
			end if
		end if
		if Request.Querystring("Status3") <> "" then
			m_sStatus3 = Request.Querystring("Status3")
			if right(sSQL,1) = "(" then
			else
			end if
		end if
	end if
	if Request.Querystring("StatFrom") <> "" then
		m_dStatFrom = Request.Querystring("StatFrom")
	end if
	if Request.Querystring("StatThru") <> "" then
		m_dStatThru = Request.Querystring("StatThru")
	end if
	if Request.Querystring("RSN") <> "" then
		m_sRSN = Request.Querystring("RSN")
	end if
	if Request.Querystring("ServiceFrom") <> "" then
		m_dServiceFrom = Request.Querystring("ServiceFrom")
	end if
	if Request.Querystring("ServiceThru") <> "" then
		m_dServiceThru = Request.Querystring("ServiceThru")
	end if
	if Request.Querystring("BenPeriod") <> "" then
		m_iBenPeriod = Request.Querystring("BenPeriod")
	end if
	if Request.Querystring("BenFrom") <> "" then
		m_dBenFrom = Request.Querystring("BenFrom")
	end if
	if Request.Querystring("BenThru") <> "" then
		m_dBenThru = Request.Querystring("BenThru")
	end if
	if Request.Querystring("AdmitFrom") <> "" then
		m_dAdmitFrom = Request.Querystring("AdmitFrom")
	end if
	if Request.Querystring("AdmitThru") <> "" then
		m_dAdmitThru = Request.Querystring("AdmitThru")
	end if
	if Request.Querystring("DisFrom") <> "" then
		m_dDisFrom = Request.Querystring("DisFrom")
	end if
	if Request.Querystring("DisThru") <> "" then
		m_dDisThru = Request.Querystring("DisThru")
	end if
	if Request.Querystring("AdmitType") <> "" then
		m_sAdmitType = Request.Querystring("AdmitType")
	end if
	if Request.Querystring("BillType") <> "" then
		m_iBillType = Request.Querystring("BillType")
	end if
	if Request.Querystring("BillFrom") <> "" then
		m_dBillFrom = Request.Querystring("BillFrom")
	end if
	if Request.Querystring("BillThru") <> "" then
		m_dBillThru = Request.Querystring("BillThru")
	end if
	if Request.Querystring("Confine") <> "" then
		m_iConfine = Request.Querystring("Confine")
	end if
	if Request.Querystring("ICD9_Proc") <> "" then
		m_iICD9_Proc = Request.Querystring("ICD9_Proc")
	end if
	if Request.Querystring("DRG") <> "" then
		m_iDRG = Request.Querystring("DRG")
	end if
	if Request.Querystring("RevCode") <> "" then
		m_iRevCode = Request.Querystring("RevCode")
	end if
	if Request.Querystring("DiagCode") <> "" then
		m_sDiagCode = Request.Querystring("DiagCode")
	end if
	if Request.Querystring("ProcCode") <> "" then
		m_sProcCode = Request.Querystring("ProcCode")
	end if
	if Request.Querystring("Modifier") <> "" then
		m_sModifier = Request.Querystring("Modifier")
	end if
	if Request.Querystring("TrackCode") <> "" then
		m_sTrackCode = Request.Querystring("TrackCode")
	end if
	if Request.Querystring("CovCode") <> "" then
		m_sCovCode = Request.Querystring("CovCode")
	end if
	if Request.Querystring("POSCode") <> "" then
		m_iPOSCode = Request.Querystring("POSCode")
	end if
	if Request.Querystring("PayCode") <> "" then
		m_sPayCode = Request.Querystring("PayCode")
	end if
	if Request.Querystring("Charge") <> "" then
		m_cCharge = Request.Querystring("Charge")
	end if
	if Request.Querystring("CheckAcct") <> "" then
		m_sCheckAcct = Request.Querystring("CheckAcct")
	end if
	if Request.Querystring("CheckNum") <> "" then
		m_iCheckNum = Request.Querystring("CheckNum")
	end if
	if Request.Querystring("PPO") <> "" then
		m_sPPO = Request.Querystring("PPO")
	end if
	if Request.Querystring("PPOOverride") <> "" then
		m_bPPOCheck = true
	else
		m_bPPOCheck = false
	end if
	if Request.Querystring("OverPay") <> "" then
		m_bOverPayCheck = true
	else
		m_bOverPayCheck = false
	end if
	if Request.Querystring("NonOverPay") <> "" then
		m_bOtherAdjustCheck = true
	else
		m_bOtherAdjustCheck = false
	end if
	if Request.Querystring("PreCalCode") <> "" then
		m_sPreCalCode = Request.Querystring("PreCalCode")
	end if
	if Request.Querystring("COB") <> "" then
		m_bCOBCheck = request.querystring("COB")
	end if
	if Request.Querystring("COBOver") <> "" then
		m_bCOBOverCheck = true
	else
		m_bCOBOverCheck = false
	end if

	if m_bShowChargelines then
    
    adoCmd.CommandText = "FindCharge"
	  adoCmd.CommandType = adCmdStoredProc
	  Set adoCmd.ActiveConnection = adoConn
    Set adoParam1 = adoCmd.CreateParameter("@P1_SSN", adInteger, adParamInput)
	  adoCmd.Parameters.Append adoParam1
    Set adoParam2 = adoCmd.CreateParameter("@P2_DepNumber", adInteger, adParamInput)
    adoCmd.Parameters.Append adoParam2
    Set adoParam3 = adoCmd.CreateParameter("@P3_HWType", adChar, adParamInput, 1)
    adoCmd.Parameters.Append adoParam3
    Set adoParam4 = adoCmd.CreateParameter("@P4_From", adDBTimeStamp, adParamInput)
    adoCmd.Parameters.Append adoParam4
    Set adoParam5 = adoCmd.CreateParameter("@P5_Thru", adDBTimeStamp, adParamInput)
    adoCmd.Parameters.Append adoParam5
    Set adoParam6 = adoCmd.CreateParameter("@P6_Fund", adChar, adParamInput, 3)
    adoCmd.Parameters.Append adoParam6
    Set adoParam7 = adoCmd.CreateParameter("@P7_PlanCode", adChar, adParamInput, 6)
    adoCmd.Parameters.Append adoParam7
    Set adoParam8 = adoCmd.CreateParameter("@P8_Status", adChar, adParamInput, 1)
    adoCmd.Parameters.Append adoParam8
    Set adoParam9 = adoCmd.CreateParameter("@P9_CovCode", adChar, adParamInput, 2)
    adoCmd.Parameters.Append adoParam9
    Set adoParam10 = adoCmd.CreateParameter("@P10_TrackCode", adChar, adParamInput, 3)
    adoCmd.Parameters.Append adoParam10
    Set adoParam11 = adoCmd.CreateParameter("@P11_Procedure", adChar, adParamInput, 5)
    adoCmd.Parameters.Append adoParam11
    Set adoParam12 = adoCmd.CreateParameter("@P12_Modifier", adChar, adParamInput, 2)
    adoCmd.Parameters.Append adoParam12
    Set adoParam13 = adoCmd.CreateParameter("@P13_Charge", adCurrency, adParamInput)
    adoCmd.Parameters.Append adoParam13
    Set adoParam14 = adoCmd.CreateParameter("@P14_PreCalCode", adChar, adParamInput, 2)
    adoCmd.Parameters.Append adoParam14
    Set adoParam15 = adoCmd.CreateParameter("@P15_Diagnosis", adChar, adParamInput, 6)
    adoCmd.Parameters.Append adoParam15
    Set adoParam16 = adoCmd.CreateParameter("@P16_TaxID", adInteger, adParamInput)
    adoCmd.Parameters.Append adoParam16
    Set adoParam17 = adoCmd.CreateParameter("@P17_Associate", adInteger, adParamInput)
    adoCmd.Parameters.Append adoParam17
    Set adoParam18 = adoCmd.CreateParameter("@P18_PendDeny", adChar, adParamInput, 4)
    adoCmd.Parameters.Append adoParam18
    Set adoParam19 = adoCmd.CreateParameter("@P19_AdmitType", adChar, adParamInput, 1)
    adoCmd.Parameters.Append adoParam19
    Set adoParam20 = adoCmd.CreateParameter("@P20_BillType", adInteger, adParamInput)
    adoCmd.Parameters.Append adoParam20
    Set adoParam21 = adoCmd.CreateParameter("@P21_OVP", adInteger, adParamInput)
    adoCmd.Parameters.Append adoParam21
    Set adoParam22 = adoCmd.CreateParameter("@P22_COB", adInteger, adParamInput)
    adoCmd.Parameters.Append adoParam22
    Set adoParam23 = adoCmd.CreateParameter("@P23_CheckAccount", adChar, adParamInput, 6)
    adoCmd.Parameters.Append adoParam23
    Set adoParam24 = adoCmd.CreateParameter("@P24_CheckNumber", adInteger, adParamInput)
    adoCmd.Parameters.Append adoParam24
    Set adoParam25 = adoCmd.CreateParameter("@P25_StatusFrom", adDBTimeStamp, adParamInput)
    adoCmd.Parameters.Append adoParam25
    Set adoParam26 = adoCmd.CreateParameter("@P26_StatusThru", adDBTimeStamp, adParamInput)
    adoCmd.Parameters.Append adoParam26
    adoCmd("@P1_SSN") = m_iSSN
    adoCmd("@P2_DepNumber") = m_iDepNum
    adoCmd("@P3_HWType") = m_sHWType
    adoCmd("@P4_From") = m_dServiceFrom
    adoCmd("@P5_Thru") = m_dServiceThru
    adoCmd("@P6_Fund") = m_sFund
    adoCmd("@P7_PlanCode") = m_sPlan
    adoCmd("@P8_Status") = m_sStatus1
    adoCmd("@P9_CovCode") = m_sCovCode
    adoCmd("@P10_TrackCode") = m_sTrackCode
    adoCmd("@P11_Procedure") = m_sProcCode
    adoCmd("@P12_Modifier") = m_sModifier
    adoCmd("@P13_Charge") = m_cCharge
    adoCmd("@P14_PreCalCode") = m_sPreCalCode
    adoCmd("@P15_Diagnosis") = m_sDiagCode
    adoCmd("@P16_TaxID") = m_iTaxID
    adoCmd("@P17_Associate") = m_iAssocNum
    adoCmd("@P18_PendDeny") = m_sRSN
    adoCmd("@P19_AdmitType") = m_sAdmitType
    adoCmd("@P20_BillType") = m_iBillType
    adoCmd("@P21_OVP") = m_iOVP
    adoCmd("@P22_COB") = m_bCOBCheck
    adoCmd("@P23_CheckAccount") = m_sCheckAcct
    adoCmd("@P24_CheckNumber") = m_iCheckNum
    adoCmd("@P25_StatusFrom") = m_dStatFrom
    adoCmd("@P26_StatusThru") = m_dStatThru
    
	else
		adoCmd.CommandText = "PB_FindClaims"
	  adoCmd.CommandType = adCmdStoredProc
	  Set adoCmd.ActiveConnection = adoConn
    Set adoParam1 = adoCmd.CreateParameter("@P1_SSN", adInteger, adParamInput)
	  adoCmd.Parameters.Append adoParam1
    Set adoParam2 = adoCmd.CreateParameter("@P2_DepNumber", adInteger, adParamInput)
    adoCmd.Parameters.Append adoParam2
    Set adoParam3 = adoCmd.CreateParameter("@P3_HWType", adChar, adParamInput, 1)
    adoCmd.Parameters.Append adoParam3
    adoCmd("@P1_SSN") = m_iSSN
    adoCmd("@P2_DepNumber") = m_iDepNum
    adoCmd("@P3_HWType") = m_sHWType
    
	end if 
end if
%>
<html>
<head>
<title>Details</title>
</head>
<body TOPMARGIN="2" LEFTMARGIN="2" RIGHTMARGIN="0" LANGUAGE="VBScript" onload="UpdateScreen(3)">
<link REL="STYLESHEET" HREF="styles/CritTable.css">
<table WIDTH="100%" COLS="3">
	<tr>
		<td ALIGN="CENTER" WIDTH="10%">
<%
			if not Session("IsClerk") then
%>
				<img SRC="images/log.gif" onClick="LogCall()">
<%
			end if
%>
		</td>
		<td ALIGN="CENTER">
			<font SIZE="+2" face="verdana, arial, helvetica"><strong>Claim Search
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
		<td ALIGN="CENTER" WIDTH="20%">
			<img SRC="images/bluebar2.gif" onClick="history.go(-1)">&nbsp;&nbsp;<img src="images/bluebar3.gif" onClick="history.go(+1)" border="0">
		</td>
	</tr>
</table>
<table COLS="2" WIDTH="100%" CELLPADDING="0" CELLSPACING="0" border="0">
	<tr>
		<td WIDTH="60%">
			<input TYPE="BUTTON" NAME="NewQuery" VALUE="Claims" onClick="ValidSend(0)">
			<input TYPE="BUTTON" NAME="NewQuery" VALUE="Chargeline" onClick="ValidSend(-1)">
      <input type="button" name="Phone" value="Log Call" onClick="LogCall()">
			<!--- <input TYPE="BUTTON" NAME="ClearCrits" VALUE="Clear Query" onClick="ValidSend(-2)"> --->
		</td>
		<td ALIGN="RIGHT" NAME="MorePages" ID="MorePages">
<%
			if bUseSQL then	
        adoRS.Open adoCmd
				if adoRS.EOF then
					adoRS.Close
					adoConn.Close
					set adoRS = nothing
					set adoConn = nothing
					sNoMatches = "<font size=2 color='red' face='verdana, arial, helvetica'><b>No matches were found -- try a different search.</b></font>"
					bUseSQL = false
				else
          adoRS.PageSize = m_iPageSize3 ' Number of rows per page
					iPageCount = adoRS.PageCount
					'adoRS.AbsolutePage = iPageNo
					if iPageCount > 1 then
						sTemp = "<FONT SIZE=2 face='verdana, arial, helvetica'><b>Page " & iPageNo & " of " & iPageCount & "</b></FONT>&nbsp;&nbsp;"
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
<div STYLE="height=33%; width:100%;  overflow=auto">
	<form ID="Criteria" METHOD="GET" ACTION="ClaimSearch.asp" TARGET="Details">
	  <input TYPE="HIDDEN" ID="ActionType" NAME="ActionType" VALUE>
		<input TYPE="HIDDEN" ID="ClaimType" NAME="ClaimType" VALUE>
    
		<table CLASS="CriteriaTable" COLS="8" CELLPADDING="0" CELLSPACING="0">
      <tr>
				<td ALIGN="RIGHT" class="White">SSN:</td>
				<td><input TYPE="TEXT" ID="SSN" NAME="SSN" SIZE="10" MAXLENGTH=10 VALUE="<%= m_iSSN%>"></td>
				<td ALIGN="RIGHT" class="White">Dep. #:</td>
				<td><input TYPE="TEXT" ID="DepNum" NAME="DepNum" SIZE="3" MAXLENGTH=5 VALUE="<%= m_iDepNum%>"></td>
				<td ALIGN="RIGHT" class="White">HWType:</td>
				<td><input TYPE="TEXT" ID="HWType" NAME="HWType" SIZE="3" MAXLENGTH=1 VALUE="<%= m_sHWType%>"></td>
				<td ALIGN="RIGHT" class="White"></td>
				<td></td>
			</tr>
			<tr>
				<td ALIGN="RIGHT" class="White">Claim Service Date From:</td>
				<td><input TYPE="TEXT" ID="ServiceFrom" NAME="ServiceFrom" SIZE="10" MAXLENGTH=10 VALUE="<%= m_dServiceFrom%>"></td>
				<td ALIGN="RIGHT" class="White">Claim Service Date Thru:</td>
				<td><input TYPE="TEXT" ID="ServiceThru" NAME="ServiceThru" SIZE="10" MAXLENGTH=10 VALUE="<%= m_dServiceThru%>"></td>
				<td ALIGN="RIGHT" class="White">Status:</td>
				<td>
          <input TYPE="TEXT" ID="Status1" NAME="Status1" SIZE="2" MAXLENGTH=1 VALUE="<%= m_sStatus1%>">
					<input TYPE="TEXT" ID="Status2" NAME="Status2" SIZE="2" MAXLENGTH=1 VALUE="<%= m_sStatus2%>">
					<input TYPE="TEXT" ID="Status3" NAME="Status3" SIZE="2" MAXLENGTH=1 VALUE="<%= m_sStatus3%>">
        </td>
				<td align="right" class="White">Pend/Deny Reason:</td>
        <td><input TYPE="TEXT" ID="RSN" NAME="RSN" SIZE="5" MAXLENGTH=4 VALUE="<%= m_sRSN%>"></td>
			</tr>
			<tr>
				<td ALIGN="RIGHT" class="White">Provider Tax ID:</td>
				<td><input TYPE="TEXT" ID="TaxID" NAME="TaxID" SIZE="10" MAXLENGTH=10 VALUE="<%= m_iTaxID%>"></td>
				<td ALIGN="RIGHT" class="White">Provider Name:</td>
				<td><input TYPE="TEXT" ID="ProvName" NAME="ProvName" SIZE="15" MAXLENGTH=40 VALUE="<%= m_sProvName%>"></td>
				<td ALIGN="RIGHT" class="White">Assoc. #:</td>
				<td><input TYPE="TEXT" ID="AssocNum" NAME="AssocNum" SIZE="3" MAXLENGTH=10 VALUE="<%= m_iAssocNum%>"></td>
				<td align="right" class="White">Coverage Code:</td>
        <td><input TYPE="TEXT" ID="CovCode" NAME="CovCode" SIZE="2" MAXLENGTH=2 VALUE="<%= m_sCovCode%>"></td>
			</tr>
			<tr>
				<td ALIGN="RIGHT" class="White">Track Code:</td>
				<td><input TYPE="TEXT" ID="TrackCode" NAME="TrackCode" SIZE="5" MAXLENGTH=3 VALUE="<%= m_sTrackCode%>"></td>
				<td ALIGN="RIGHT" class="White">Diagnosis:</td>
				<td><input TYPE="TEXT" ID="DiagCode" NAME="DiagCode" SIZE="8" MAXLENGTH=6 VALUE="<%= m_sDiagCode%>"></td>
				<td ALIGN="RIGHT" class="White">Procedure Code:</td>
				<td><input TYPE="TEXT" ID="ProcCode" NAME="ProcCode" SIZE="7" MAXLENGTH=5 VALUE="<%= m_sProcCode%>"></td>
				<td ALIGN="RIGHT" class="White"></td>
				<td></td>
			</tr>
			<tr>
				<td ALIGN="RIGHT" class="White">Charge:</td>
				<td><input TYPE="TEXT" ID="Charge" NAME="Charge" SIZE="8" MAXLENGTH=11 VALUE="<%= m_cCharge%>"></td>
				<td ALIGN="RIGHT" class="White">Pre Cal Code:</td>
				<td><input TYPE="TEXT" ID="PreCalCode" NAME="PreCalCode" SIZE="4" MAXLENGTH=2 VALUE="<%= m_sPreCalCode%>"></td>
				<td ALIGN="RIGHT" class="White">Overpayment Adjustments:</td>
				<td><input TYPE="CHECKBOX" ID="OverPay" NAME="OverPay" <% if m_bOverPayCheck then Response.Write " CHECKED " end if %>></td>
				<td ALIGN="RIGHT" class="White">Non-Overpayment Adjustments:</td>
				<td><input TYPE="CHECKBOX" ID="NonOverPay" NAME="NonOverPay" <% if m_bOtherAdjustCheck then Response.Write " CHECKED " end if %>></td>
			</tr>
			<tr>
				<td ALIGN="RIGHT" class="White">COB(1=Yes 0=No):</td>
				<td><input type="text" id="COB" name="COB" size="4" maxlength="4" value="<%= m_bCOBCheck %>"></td>
				<td ALIGN="RIGHT" class="White">COB Override:</td>
				<td><input TYPE="CHECKBOX" ID="COBOver" NAME="COBOver" <% if m_bCOBOverCheck then Response.Write " CHECKED " end if %>></td>
				<td ALIGN="RIGHT" class="White">Check Account:</td>
				<td><input TYPE="TEXT" ID="CheckAcct" NAME="CheckAcct" SIZE="8" MAXLENGTH=6 VALUE="<%= m_sCheckAcct%>"></td>
				<td ALIGN="RIGHT" class="White">Check Number:</td>
				<td><input TYPE="TEXT" ID="CheckNum" NAME="CheckNum" SIZE="7" MAXLENGTH=10 VALUE="<%= m_iCheckNum%>"></td>
			</tr>
			<tr>
				<td ALIGN="RIGHT" class="White">Status Date From:</td>
				<td><input TYPE="TEXT" ID="StatFrom" NAME="StatFrom" SIZE="10" MAXLENGTH=10 VALUE="<%= m_dStatFrom%>"></td>
				<td ALIGN="RIGHT" class="White">Status Date Thru:</td>
				<td><input TYPE="TEXT" ID="StatThru" NAME="StatThru" SIZE="10" MAXLENGTH=10 VALUE="<%= m_dStatThru%>"></td>
				<td ALIGN="RIGHT" class="White"></td>
				<td></td>
				<td ALIGN="RIGHT" class="White"></td>
				<td></td>
			</tr>
      <tr>
				<td ALIGN="RIGHT" class="White">Log Date From:</td>
				<td><input TYPE="TEXT" ID="FromDate" NAME="FromDate" SIZE="10" MAXLENGTH=10 VALUE="<%= m_dFromDate%>"></td>
				<td ALIGN="RIGHT" class="White">Log Date Thru:</td>
				<td><input TYPE="TEXT" ID="ThruDate" NAME="ThruDate" SIZE="10" MAXLENGTH=10 VALUE="<%= m_dThruDate%>"></td>
				<td ALIGN="RIGHT" class="White">Sequence:</td>
				<td><input TYPE="TEXT" ID="Sequence" NAME="Sequence" SIZE="6" MAXLENGTH=10 VALUE="<%= m_iSequence%>"></td>
				<td ALIGN="RIGHT" class="White">Distr:</td>
				<td><input TYPE="TEXT" ID="Distr" NAME="Distr" SIZE="3" MAXLENGTH=5 VALUE="<%= m_iDistr%>"></td>
			</tr>
      <tr>
        <td colspan="8">&nbsp;</td>
      </tr>
		</table>
	</form>
</div>
<br>
<%
    
    
    if bUseSQL then
      if adoRS("locktype") <> "" then
        If Manager = "True" or Supervisor = "True" or Auditor = "True" Then 
          if m_bShowChargelines then
          iRowCount = adoRS.PageSize

%>
<table width="100%" border="0">
  <tr>
    <td><font size="2" face="verdana, arial, helvetica"><b>This is the Claim Pick List for: <%= adoRS("Dependent Name") %>/<%= m_iSSN %></b></font><% if adoRS("locktype") <> "" then %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="images/lock1.gif" width="32" height="32" alt="LOCKED FILE!!"><font size="2" face="verdana, arial, helvetica" color="red"><b>&nbsp;&nbsp;&nbsp;LOCKED "<%= adoRS("locktype") %>"  RECORDS!!!</b></font><% end if %></td>
    <td align="right"><font size="2" face="verdana, arial, helvetica"><b>Maximum Row Return is <%= iRowCount %> Records.</b></font></td>
  </tr>
</table>

<div style="height:68%; width:100%;">
		<table WIDTH="100%" BORDER="1" bordercolor="white" bgcolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			<tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
				<td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>Claim  Number</b></font></td>
        <td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>Tracking</b></font></td>
        <td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>Adj Hist</b></font></td>
        <td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>FON</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Status<br>Date</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>ChkNum</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Provider</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>PlaCo</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Status</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>FromDate</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>ThruDate</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Line #</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Dep #</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>TaxID</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Assoc#</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Charge</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Paid<br>Amount</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>CC</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>HWT</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>ProcCode</b></font></td>
      </tr>
<% 
	        iRowCount = adoRS.PageSize		
			    Do While Not adoRS.EOF and iRowCount > 0 
      
%> 
				<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
					<td NOWRAP bgcolor="#F0F0F0"><a HREF="aspexec.asp?LogDate=<%= adoRS("LogDate")%>&amp;Sequence=<%= adoRS("Sequence")%>&amp;Distr=<%=  adoRS("Distr")%>&ClaimNum=<%= ConvertToClaimNum(adoRS("LogDate"),adoRS("Sequence"),adoRS("Distr")) %>" onClick="WorkingStatus()"><%= ConvertToClaimNum(adoRS("LogDate"),adoRS("Sequence"),adoRS("Distr")) %></a></td>
          <td nowrap bgcolor="#F0F0F0"><a href="Tracking.asp?SSN=<%= m_iSSN %>&DepNum=<%= m_iDepNum %>&HWType=<%= adoRS("HWType") %>&Fund=<%= adoRS("Fund") %>&Years=<%= Year(Date) %>" onClick="WorkingStatus()">Tracking</a></td>
          <td nowrap bgcolor="#F0F0F0"><a href="AdjustHistory.asp?LogDate=<%= adoRS("LogDate") %>&Sequence=<%= adoRS("Sequence") %>&Distr=<%=adoRS("Distr") %>" onClick="WorkingStatus()">Adj Hist</a></td>
          <td nowrap bgcolor="#F0F0F0"><a href="PhoneSearch.asp?SSN=<%= m_iSSN %>&ClaimNum=<%= ConvertToClaimNum(adoRS("LogDate"),adoRS("Sequence"),adoRS("Distr")) %>&LogType=Claim">FON</td>
          <td nowrap bgcolor="#F0F0F0"><% if adoRS("StatusDate") <> "" then %><%= FORMATDATETIME(adoRS("StatusDate"),vbShortDate)  %><% else %>&nbsp;<%= adoRS("StatusDate") %><% end if %></td>
          <td nowrap bgcolor="#F0F0F0"><a href="checkdetails.asp?CheckNum=<%= adoRS("CheckNumber") %>&CheckAcct=<%= adoRS("CheckAccount") %>" onClick="WorkingStatus()"><%= adoRS("CheckNumber") %></a></td>
          <td nowrap bgcolor="#F0F0F0"><a href="providersearch.asp?TaxID=<%= adoRS("TaxID") %>&FullName=<%= adoRS("ProviderName") %>" onClick="WorkingStatus()"><%= adoRS("ProviderName") %></a></td>
          <td align="center" nowrap bgcolor="#F0F0F0"><%= adoRS("PlanCode") %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("Status") %></td>
          <td nowrap bgcolor="#F0F0F0"><% if adoRS("FromDate") <> "" then %><%= FORMATDATETIME(adoRS("FromDate"),vbShortDate) %><% else %>&nbsp;<%= adoRS("FromDate") %><% end if %></td>
          <td nowrap bgcolor="#F0F0F0"><% if adoRS("ThruDate") <> "" then %><%= FORMATDATETIME(adoRS("ThruDate"),vbShortDate) %><% else %>&nbsp;<%= adoRS("ThruDate") %><% end if %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("LineNumber") %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("DepNumber") %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("TaxID") %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("Associate") %></td>
          <td nowrap bgcolor="#F0F0F0"><% if adoRS("ChargeAmount") <> "" then %><%= formatcurrency(adoRS("ChargeAmount")) %><% else %>&nbsp;<%= adoRS("ChargeAmount") %><% end if %></td>
          <td nowrap bgcolor="#F0F0F0"><% if adoRS("PaidAmount") <> "" then %><%= formatcurrency(adoRS("PaidAmount")) %><% else %>&nbsp;<%= adoRS("PaidAmount") %><% end if %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("CovCode") %></td>
          <td align="center" nowrap bgcolor="#F0F0F0"><%= adoRS("HWType") %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("ProcedureCode") %></td>
        </tr>
<% 
			    iRowCount = iRowCount - 1
			    adoRS.MoveNext
			    Loop
          else 'else for m_bShowChargelines
%>
      </table>
      <table width="100%" border="0">
        <tr>
          <td><font size="2" face="verdana, arial, helvetica"><b><% if adoRS("locktype") <> "" then %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="images/lock1.gif" width="32" height="32" alt="LOCKED FILE!!"><font size="2" face="verdana, arial, helvetica" color="red"><b>&nbsp;&nbsp;&nbsp;LOCKED "<%= adoRS("locktype") %>"  RECORDS!!!</b></font><% end if %></td>
        </tr>
      </table>
      
      <table WIDTH="100%" BORDER="1" bordercolor="white" bgcolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			  <tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
          <td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>Claim  Number</b></font></td>
          <td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>Adj Hist</b></font></td>
          <td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>PhoneCall</b></font></td>
          <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>StatusDate</b></font></td>
          <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Provider</b></font></td>
          <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Tax ID</b></font></td>
          <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Assoc #</b></font></td>
          <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Plan Code</b></font></td>
          <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Fund</b></font></td>
          <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Status</b></font></td>
          <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>HWType</b></font></td>
        </tr>
				
<%
			    iRowCount = adoRS.PageSize
			    Do While Not adoRS.EOF and iRowCount > 0
%>
				<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
          <td NOWRAP bgcolor="#F0F0F0"><a HREF="aspexec.asp?LogDate=<%= adoRS("LogDate")%>&amp;Sequence=<%= adoRS("Sequence")%>&amp;Distr=<%=  adoRS("Distr")%>&ClaimNum=<%= ConvertToClaimNum(adoRS("LogDate"),adoRS("Sequence"),adoRS("Distr")) %>" onClick="WorkingStatus()"><%= ConvertToClaimNum(adoRS("LogDate"),adoRS("Sequence"),adoRS("Distr")) %></a></td>
          <td nowrap bgcolor="#F0F0F0"><a href="AdjustHistory.asp?LogDate=<%= adoRS("LogDate") %>&Sequence=<%= adoRS("Sequence") %>&Distr=<%=adoRS("Distr") %>" onClick="WorkingStatus()">Adj Hist</a></td>
          <td nowrap bgcolor="#F0F0F0"><a href="PhoneSearch.asp?SSN=<%= m_iSSN %>&ClaimNum=<%= ConvertToClaimNum(adoRS("LogDate"),adoRS("Sequence"),adoRS("Distr")) %>&LogType=Claim">Phone Calls</td>
          <td nowrap bgcolor="#F0F0F0"><%= formatdatetime(adoRS("StatusDate"),vbShortDate) %></td>
          <td nowrap bgcolor="#F0F0F0"><a href="providersearch.asp?TaxID=<%= adoRS("TaxID") %>&FullName=<%= adoRS("ProviderName") %>" onClick="WorkingStatus()"><%= adoRS("ProviderName") %></a></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("TaxID") %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("Associate") %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("PlanCode") %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("Fund") %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("Status") %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("HWType") %></td>
        </tr>
<%
			
			    iRowCount = iRowCount - 1
			    adoRS.MoveNext
			    Loop
          adoRS.Close
			    adoConn.Close
			    set adoRS = nothing
			    set adoConn = nothing
          end if
%>
  </table>
</div>
<%
        else 'else Admin, Manager, Auditor not true
          if m_bShowChargelines then
          iRowCount = adoRS.PageSize
%>        
<table width="100%" border="0">
  <tr>
    <td><font size="2" face="verdana, arial, helvetica"><b>This is the Claim Pick List for: <%= adoRS("Dependent Name") %>/<%= m_iSSN %></b></font><% if adoRS("locktype") <> "" then %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="images/lock1.gif" width="32" height="32" alt="LOCKED FILE!!"><font size="2" face="verdana, arial, helvetica" color="red"><b>&nbsp;&nbsp;&nbsp;LOCKED "<%= adoRS("locktype") %>"  RECORDS!!!</b></font><% end if %></td>
    <td align="right"><font size="2" face="verdana, arial, helvetica"><b>Maximum Row Return is <%= iRowCount %> Records.</b></font></td>
  </tr>
</table>

<div style="height:68%; width:100%;">
		<table WIDTH="100%" BORDER="1" bordercolor="white" bgcolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			<tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
				<td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>Claim  Number</b></font></td>
        <td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>Tracking</b></font></td>
        <td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>Adj Hist</b></font></td>
        <td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>FON</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Status<br>Date</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>ChkNum</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Provider</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>PlaCo</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Status</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>FromDate</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>ThruDate</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Line #</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Dep #</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>TaxID</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Assoc#</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Charge</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Paid<br>Amount</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>CC</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>HWT</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>ProcCode</b></font></td>
      </tr>
<% 
	        iRowCount = adoRS.PageSize		
			    Do While Not adoRS.EOF and iRowCount > 0
            if instr(adoRS("locktype"),adoRS("HWType")) = 0 then 
      
%> 
				<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
					<td NOWRAP bgcolor="#F0F0F0"><a HREF="aspexec.asp?LogDate=<%= adoRS("LogDate")%>&amp;Sequence=<%= adoRS("Sequence")%>&amp;Distr=<%=  adoRS("Distr")%>&ClaimNum=<%= ConvertToClaimNum(adoRS("LogDate"),adoRS("Sequence"),adoRS("Distr")) %>" onClick="WorkingStatus()"><%= ConvertToClaimNum(adoRS("LogDate"),adoRS("Sequence"),adoRS("Distr")) %></a></td>
          <td nowrap bgcolor="#F0F0F0"><a href="Tracking.asp?SSN=<%= m_iSSN %>&DepNum=<%= m_iDepNum %>&HWType=<%= adoRS("HWType") %>&Fund=<%= adoRS("Fund") %>&Years=<%= Year(Date) %>" onClick="WorkingStatus()">Tracking</a></td>
          <td nowrap bgcolor="#F0F0F0"><a href="AdjustHistory.asp?LogDate=<%= adoRS("LogDate") %>&Sequence=<%= adoRS("Sequence") %>&Distr=<%=adoRS("Distr") %>" onClick="WorkingStatus()">Adj Hist</a></td>
          <td nowrap bgcolor="#F0F0F0"><a href="PhoneSearch.asp?SSN=<%= m_iSSN %>&ClaimNum=<%= ConvertToClaimNum(adoRS("LogDate"),adoRS("Sequence"),adoRS("Distr")) %>&LogType=Claim">FON</td>
          <td nowrap bgcolor="#F0F0F0"><% if adoRS("StatusDate") <> "" then %><%= FORMATDATETIME(adoRS("StatusDate"),vbShortDate)  %><% else %>&nbsp;<%= adoRS("StatusDate") %><% end if %></td>
          <td nowrap bgcolor="#F0F0F0"><a href="checkdetails.asp?CheckNum=<%= adoRS("CheckNumber") %>&CheckAcct=<%= adoRS("CheckAccount") %>" onClick="WorkingStatus()"><%= adoRS("CheckNumber") %></a></td>
          <td nowrap bgcolor="#F0F0F0"><a href="providersearch.asp?TaxID=<%= adoRS("TaxID") %>&FullName=<%= adoRS("ProviderName") %>" onClick="WorkingStatus()"><%= adoRS("ProviderName") %></a></td>
          <td align="center" nowrap bgcolor="#F0F0F0"><%= adoRS("PlanCode") %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("Status") %></td>
          <td nowrap bgcolor="#F0F0F0"><% if adoRS("FromDate") <> "" then %><%= FORMATDATETIME(adoRS("FromDate"),vbShortDate) %><% else %>&nbsp;<%= adoRS("FromDate") %><% end if %></td>
          <td nowrap bgcolor="#F0F0F0"><% if adoRS("ThruDate") <> "" then %><%= FORMATDATETIME(adoRS("ThruDate"),vbShortDate) %><% else %>&nbsp;<%= adoRS("ThruDate") %><% end if %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("LineNumber") %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("DepNumber") %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("TaxID") %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("Associate") %></td>
          <td nowrap bgcolor="#F0F0F0"><% if adoRS("ChargeAmount") <> "" then %><%= formatcurrency(adoRS("ChargeAmount")) %><% else %>&nbsp;<%= adoRS("ChargeAmount") %><% end if %></td>
          <td nowrap bgcolor="#F0F0F0"><% if adoRS("PaidAmount") <> "" then %><%= formatcurrency(adoRS("PaidAmount")) %><% else %>&nbsp;<%= adoRS("PaidAmount") %><% end if %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("CovCode") %></td>
          <td align="center" nowrap bgcolor="#F0F0F0"><%= adoRS("HWType") %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("ProcedureCode") %></td>
        </tr>
<% 
			      iRowCount = iRowCount - 1
			      end if
            adoRS.MoveNext
            Loop
          else 'else for m_bShowChargelines
%>
      </table>
      <table width="100%" border="0">
        <tr>
          <td><font size="2" face="verdana, arial, helvetica"><b><% if adoRS("locktype") <> "" then %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img SRC="images/lock1.gif" width="32" height="32" alt="LOCKED FILE!!"><font size="2" face="verdana, arial, helvetica" color="red"><b>&nbsp;&nbsp;&nbsp;LOCKED "<%= adoRS("locktype") %>"  RECORDS!!!</b></font><% end if %></td>
        </tr>
      </table>
      
      <table WIDTH="100%" BORDER="1" bordercolor="white" bgcolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			  <tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
          <td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>Claim  Number</b></font></td>
          <td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>Adj Hist</b></font></td>
          <td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>PhoneCall</b></font></td>
          <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>StatusDate</b></font></td>
          <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Provider</b></font></td>
          <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Tax ID</b></font></td>
          <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Assoc #</b></font></td>
          <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Plan Code</b></font></td>
          <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Fund</b></font></td>
          <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Status</b></font></td>
          <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>HWType</b></font></td>
        </tr>
				
<%
			    iRowCount = adoRS.PageSize
			    Do While Not adoRS.EOF and iRowCount > 0
            if instr(adoRS("locktype"),adoRS("HWType")) = 0 then 
%>
				<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
          <td NOWRAP bgcolor="#F0F0F0"><a HREF="aspexec.asp?LogDate=<%= adoRS("LogDate")%>&amp;Sequence=<%= adoRS("Sequence")%>&amp;Distr=<%=  adoRS("Distr")%>&ClaimNum=<%= ConvertToClaimNum(adoRS("LogDate"),adoRS("Sequence"),adoRS("Distr")) %>" onClick="WorkingStatus()"><%= ConvertToClaimNum(adoRS("LogDate"),adoRS("Sequence"),adoRS("Distr")) %></a></td>
          <td nowrap bgcolor="#F0F0F0"><a href="AdjustHistory.asp?LogDate=<%= adoRS("LogDate") %>&Sequence=<%= adoRS("Sequence") %>&Distr=<%=adoRS("Distr") %>" onClick="WorkingStatus()">Adj Hist</a></td>
          <td nowrap bgcolor="#F0F0F0"><a href="PhoneSearch.asp?SSN=<%= m_iSSN %>&ClaimNum=<%= ConvertToClaimNum(adoRS("LogDate"),adoRS("Sequence"),adoRS("Distr")) %>&LogType=Claim">Phone Calls</td>
          <td nowrap bgcolor="#F0F0F0"><%= formatdatetime(adoRS("StatusDate"),vbShortDate) %></td>
          <td nowrap bgcolor="#F0F0F0"><a href="providersearch.asp?TaxID=<%= adoRS("TaxID") %>&FullName=<%= adoRS("ProviderName") %>" onClick="WorkingStatus()"><%= adoRS("ProviderName") %></a></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("TaxID") %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("Associate") %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("PlanCode") %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("Fund") %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("Status") %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("HWType") %></td>
        </tr>
<%
			
			      iRowCount = iRowCount - 1
            end if
			      adoRS.MoveNext
			      Loop
            adoRS.Close
			      adoConn.Close
			      set adoRS = nothing
			      set adoConn = nothing
          end if
%>
  </table>
</div>
<%
        end if
      else 'else adoRS("locktype") is empty
        if m_bShowChargelines then
        iRowCount = adoRS.PageSize

%>
<table width="100%" border="0">
  <tr>
    <td><font size="2" face="verdana, arial, helvetica"><b>This is the Claim Pick List for: <%= adoRS("Dependent Name") %>/<%= m_iSSN %></b></font></td>
    <td align="right"><font size="2" face="verdana, arial, helvetica"><b>Maximum Row Return is <%= iRowCount %> Records.</b></font></td>
  </tr>
</table>

<div style="height:68%; width:100%;">
		<table WIDTH="100%" BORDER="1" bordercolor="white" bgcolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			<tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
				<td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>Claim  Number</b></font></td>
        <td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>Tracking</b></font></td>
        <td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>Adj Hist</b></font></td>
        <td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>FON</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Status<br>Date</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>ChkNum</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Provider</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>PlaCo</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Status</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>FromDate</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>ThruDate</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Line #</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Dep #</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>TaxID</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Assoc#</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Charge</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Paid<br>Amount</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>CC</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>HWT</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>ProcCode</b></font></td>
      </tr>
<% 
	      iRowCount = adoRS.PageSize		
			  Do While Not adoRS.EOF and iRowCount > 0
      
%> 
				<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
					<td NOWRAP bgcolor="#F0F0F0"><a HREF="aspexec.asp?LogDate=<%= adoRS("LogDate")%>&amp;Sequence=<%= adoRS("Sequence")%>&amp;Distr=<%=  adoRS("Distr")%>&ClaimNum=<%= ConvertToClaimNum(adoRS("LogDate"),adoRS("Sequence"),adoRS("Distr")) %>" onClick="WorkingStatus()"><%= ConvertToClaimNum(adoRS("LogDate"),adoRS("Sequence"),adoRS("Distr")) %></a></td>
          <td nowrap bgcolor="#F0F0F0"><a href="Tracking.asp?SSN=<%= m_iSSN %>&DepNum=<%= m_iDepNum %>&HWType=<%= adoRS("HWType") %>&Fund=<%= adoRS("Fund") %>&Years=<%= Year(Date) %>" onClick="WorkingStatus()">Tracking</a></td>
          <td nowrap bgcolor="#F0F0F0"><a href="AdjustHistory.asp?LogDate=<%= adoRS("LogDate") %>&Sequence=<%= adoRS("Sequence") %>&Distr=<%=adoRS("Distr") %>" onClick="WorkingStatus()">Adj Hist</a></td>
          <td nowrap bgcolor="#F0F0F0"><a href="PhoneSearch.asp?SSN=<%= m_iSSN %>&ClaimNum=<%= ConvertToClaimNum(adoRS("LogDate"),adoRS("Sequence"),adoRS("Distr")) %>&LogType=Claim">FON</td>
          <td nowrap bgcolor="#F0F0F0"><% if adoRS("StatusDate") <> "" then %><%= FORMATDATETIME(adoRS("StatusDate"),vbShortDate)  %><% else %>&nbsp;<%= adoRS("StatusDate") %><% end if %></td>
          <td nowrap bgcolor="#F0F0F0"><a href="checkdetails.asp?CheckNum=<%= adoRS("CheckNumber") %>&CheckAcct=<%= adoRS("CheckAccount") %>" onClick="WorkingStatus()"><%= adoRS("CheckNumber") %></a></td>
          <td nowrap bgcolor="#F0F0F0"><a href="providersearch.asp?TaxID=<%= adoRS("TaxID") %>&FullName=<%= adoRS("ProviderName") %>" onClick="WorkingStatus()"><%= adoRS("ProviderName") %></a></td>
          <td align="center" nowrap bgcolor="#F0F0F0"><%= adoRS("PlanCode") %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("Status") %></td>
          <td nowrap bgcolor="#F0F0F0"><% if adoRS("FromDate") <> "" then %><%= FORMATDATETIME(adoRS("FromDate"),vbShortDate) %><% else %>&nbsp;<%= adoRS("FromDate") %><% end if %></td>
          <td nowrap bgcolor="#F0F0F0"><% if adoRS("ThruDate") <> "" then %><%= FORMATDATETIME(adoRS("ThruDate"),vbShortDate) %><% else %>&nbsp;<%= adoRS("ThruDate") %><% end if %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("LineNumber") %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("DepNumber") %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("TaxID") %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("Associate") %></td>
          <td nowrap bgcolor="#F0F0F0"><% if adoRS("ChargeAmount") <> "" then %><%= formatcurrency(adoRS("ChargeAmount")) %><% else %>&nbsp;<%= adoRS("ChargeAmount") %><% end if %></td>
          <td nowrap bgcolor="#F0F0F0"><% if adoRS("PaidAmount") <> "" then %><%= formatcurrency(adoRS("PaidAmount")) %><% else %>&nbsp;<%= adoRS("PaidAmount") %><% end if %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("CovCode") %></td>
          <td align="center" nowrap bgcolor="#F0F0F0"><%= adoRS("HWType") %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("ProcedureCode") %></td>
        </tr>
<% 
			  iRowCount = iRowCount - 1
			  adoRS.MoveNext
			  Loop
      else 'else for m_bShowChargeLines
%>
      </table>
      <table WIDTH="100%" BORDER="1" bordercolor="white" bgcolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			  <tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
          <td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>Claim  Number</b></font></td>
          <td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>Adj Hist</b></font></td>
          <td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>PhoneCall</b></font></td>
          <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>StatusDate</b></font></td>
          <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Provider</b></font></td>
          <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Tax ID</b></font></td>
          <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Assoc #</b></font></td>
          <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Plan Code</b></font></td>
          <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Fund</b></font></td>
          <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Status</b></font></td>
          <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>HWType</b></font></td>
        </tr>
				
<%
			  iRowCount = adoRS.PageSize
			  Do While Not adoRS.EOF and iRowCount > 0
%>
				<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
          <td NOWRAP bgcolor="#F0F0F0"><a HREF="aspexec.asp?LogDate=<%= adoRS("LogDate")%>&amp;Sequence=<%= adoRS("Sequence")%>&amp;Distr=<%=  adoRS("Distr")%>&ClaimNum=<%= ConvertToClaimNum(adoRS("LogDate"),adoRS("Sequence"),adoRS("Distr")) %>" onClick="WorkingStatus()"><%= ConvertToClaimNum(adoRS("LogDate"),adoRS("Sequence"),adoRS("Distr")) %></a></td>
          <td nowrap bgcolor="#F0F0F0"><a href="AdjustHistory.asp?LogDate=<%= adoRS("LogDate") %>&Sequence=<%= adoRS("Sequence") %>&Distr=<%=adoRS("Distr") %>" onClick="WorkingStatus()">Adj Hist</a></td>
          <td nowrap bgcolor="#F0F0F0"><a href="PhoneSearch.asp?SSN=<%= m_iSSN %>&ClaimNum=<%= ConvertToClaimNum(adoRS("LogDate"),adoRS("Sequence"),adoRS("Distr")) %>&LogType=Claim">Phone Calls</td>
          <td nowrap bgcolor="#F0F0F0"><%= formatdatetime(adoRS("StatusDate"),vbShortDate) %></td>
          <td nowrap bgcolor="#F0F0F0"><a href="providersearch.asp?TaxID=<%= adoRS("TaxID") %>&FullName=<%= adoRS("ProviderName") %>" onClick="WorkingStatus()"><%= adoRS("ProviderName") %></a></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("TaxID") %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("Associate") %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("PlanCode") %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("Fund") %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("Status") %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("HWType") %></td>
        </tr>
<%
			
			  iRowCount = iRowCount - 1
			  adoRS.MoveNext
			  Loop
        adoRS.Close
			  adoConn.Close
			  set adoRS = nothing
			  set adoConn = nothing
      end if
    end if
%>
  </table>
</div>
<%
  else 'else for bUseSQL
	  if sNoMatches <> "" then
		  Response.Write sNoMatches
	  end if
  end if
%>
<div STYLE="display:none">
	<form ID="ClearCrits" METHOD="GET" ACTION="ClaimSearch.asp" TARGET="Details">
		<input TYPE="HIDDEN" ID="UserCriteria" NAME="UserCriteria" VALUE>
	</form>
</div>
</body>
<!--#include file="VBFuncs.inc" -->
<script LANGUAGE="VBSCRIPT">
  
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
		
		Criteria.DepNum.Value = trim(Criteria.DepNum.Value)
		Criteria.TaxID.Value = trim(Criteria.TaxID.Value)
		Criteria.ProvName.Value = trim(Criteria.ProvName.Value)
		Criteria.AssocNum.Value = trim(Criteria.AssocNum.Value)
		'Criteria.Fund.Value = trim(Criteria.Fund.Value)
		Criteria.Status1.Value = trim(Criteria.Status1.Value)
		Criteria.Status2.Value = trim(Criteria.Status2.Value)
		Criteria.Status3.Value = trim(Criteria.Status3.Value)
		Criteria.StatFrom.Value = trim(Criteria.StatFrom.Value)
		Criteria.StatThru.Value = trim(Criteria.StatThru.Value)
		'Criteria.Plan.Value = trim(Criteria.Plan.Value)
		Criteria.RSN.Value = trim(Criteria.RSN.Value)
		Criteria.ServiceFrom.Value = trim(Criteria.ServiceFrom.Value)
		Criteria.ServiceThru.Value = trim(Criteria.ServiceThru.Value)
		Criteria.HWType.Value = trim(Criteria.HWType.Value)
		'Criteria.BenPeriod.Value = trim(Criteria.BenPeriod.Value)
		'Criteria.BenFrom.Value = trim(Criteria.BenFrom.Value)
		'Criteria.BenThru.Value = trim(Criteria.BenThru.Value)
		'Criteria.AdmitFrom.Value = trim(Criteria.AdmitFrom.Value)
		'Criteria.AdmitThru.Value = trim(Criteria.AdmitThru.Value)
		'Criteria.DisFrom.Value = trim(Criteria.DisFrom.Value)
		'Criteria.DisThru.Value = trim(Criteria.DisThru.Value)
		'Criteria.AdmitType.Value = trim(Criteria.AdmitType.Value)
		'Criteria.BillType.Value = trim(Criteria.BillType.Value)
		'Criteria.BillFrom.Value = trim(Criteria.BillFrom.Value)
		'Criteria.BillThru.Value = trim(Criteria.BillThru.Value)
		'Criteria.Confine.Value = trim(Criteria.Confine.Value)
		'Criteria.ICD9_Proc.Value = trim(Criteria.ICD9_Proc.Value)
		'Criteria.DRG.Value = trim(Criteria.DRG.Value)
		'Criteria.RevCode.Value = trim(Criteria.RevCode.Value)
		'Criteria.DiagCode.Value = trim(Criteria.DiagCode.Value)
		Criteria.ProcCode.Value = trim(Criteria.ProcCode.Value)
		'Criteria.Modifier.Value = trim(Criteria.Modifier.Value)
		Criteria.TrackCode.Value = trim(Criteria.TrackCode.Value)
		Criteria.CovCode.Value = trim(Criteria.CovCode.Value)
		Criteria.PreCalCode.Value = trim(Criteria.PreCalCode.Value)
		'Criteria.POSCode.Value = trim(Criteria.POSCode.Value)
		'Criteria.PayCode.Value = trim(Criteria.PayCode.Value)
		Criteria.Charge.Value = trim(Criteria.Charge.Value)
		Criteria.CheckAcct.Value = trim(Criteria.CheckAcct.Value)
		Criteria.CheckNum.Value = trim(Criteria.CheckNum.Value)
		'Criteria.PPO.Value = trim(Criteria.PPO.Value)
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
		'if Criteria.LName.Value <> "" then
			'if ContainsInvalids(Criteria.LName.Value) then
				'msgbox "Please remove invalid characters from the Last Name field.",,"Invalid Data"
				'exit sub
			'end if
			'bValidData = true
		'end if
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
		'if Criteria.Fund.Value <> "" then
			'if ContainsInvalids(Criteria.Fund.Value) then
				'msgbox "Please remove invalid characters from the Fund field.",,"Invalid Data"
				'exit sub
			'end if
			'bValidData = true
		'end if
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
		'if Criteria.Plan.Value <> "" then
			'if ContainsInvalids(Criteria.Plan.Value) then
				'msgbox "Please remove invalid characters from the Plan field.",,"Invalid Data"
				'exit sub
			'end if
			'bValidData = true
		'end if
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
		'if Criteria.BenPeriod.Value <> ""  then
			'if not isNumeric(Criteria.BenPeriod.Value) then
				'msgbox "Please enter a valid Benefit Period.",,"Invalid Data"
				'exit sub
			'else
				'if Criteria.BenPeriod.Value > <%= Application("SmlIntMax")%> or _
						'Criteria.BenPeriod.Value < <%= Application("SmlIntMin")%> then
					'msgbox "Please enter a valid Benefit Period.",,"Invalid Data"
					'exit sub
				'end if
			'end if
			'bValidData = true
		'end if
		'if Criteria.BenFrom.Value <> "" then
			'if not isDate(Criteria.BenFrom.Value) then
				'msgbox "Please enter a valid date for the Benefit Date From field.",,"Invalid Data"
				'exit sub
			'end if
			'bValidData = true
		'end if
		'if Criteria.BenThru.Value <> "" then
			'if not isDate(Criteria.BenThru.Value) then
				'msgbox "Please enter a valid date for the Benefit Date Thru field.",,"Invalid Data"
				'exit sub
			'end if
			'bValidData = true
		'end if
'
' Hospital
'
		'if Criteria.AdmitFrom.Value <> "" then
			'if not isDate(Criteria.AdmitFrom.Value) then
				'msgbox "Please enter a valid date for the Admit Date From field.",,"Invalid Data"
				'exit sub
			'end if
			'bValidData = true
		'end if
		'if Criteria.AdmitThru.Value <> "" then
			'if not isDate(Criteria.AdmitThru.Value) then
				'msgbox "Please enter a valid date for the Admit Date Thru field.",,"Invalid Data"
				'exit sub
			'end if
			'bValidData = true
		'end if
		'if Criteria.DisFrom.Value <> "" then
			'if not isDate(Criteria.DisFrom.Value) then
				'msgbox "Please enter a valid date for the Discharge Date From field.",,"Invalid Data"
				'exit sub
			'end if
			'bValidData = true
		'end if
		'if Criteria.DisThru.Value <> "" then
			'if not isDate(Criteria.DisThru.Value) then
				'msgbox "Please enter a valid date for the Discharge Date Thru field.",,"Invalid Data"
				'exit sub
			'end if
			'bValidData = true
		'end if
		'if Criteria.AdmitType.Value <> "" then
			'if ContainsInvalids(Criteria.AdmitType.Value) then
				'msgbox "Please remove invalid characters from the Admit Type field.",,"Invalid Data"
				'exit sub
			'end if
			'bValidData = true
		'end if
		'if Criteria.BillType.Value <> ""  then
			'if not isNumeric(Criteria.BillType.Value) then
				'msgbox "Please enter a valid Bill Type.",,"Invalid Data"
				'exit sub
			'else
				'if Criteria.BillType.Value > <%= Application("SmlIntMax")%> or _
						'Criteria.BillType.Value < <%= Application("SmlIntMin")%> then
					'msgbox "Please enter a valid Bill Type.",,"Invalid Data"
					'exit sub
				'end if
			'end if
			'bValidData = true
		'end if
		'if Criteria.BillFrom.Value <> "" then
			'if not isDate(Criteria.BillFrom.Value) then
				'msgbox "Please enter a valid date for the Bill Date From field.",,"Invalid Data"
				'exit sub
			'end if
			'bValidData = true
		'end if
		'if Criteria.BillThru.Value <> "" then
			'if not isDate(Criteria.BillThru.Value) then
				'msgbox "Please enter a valid date for the Bill Date Thru field.",,"Invalid Data"
				'exit sub
			'end if
			'bValidData = true
		'end if
		'if Criteria.Confine.Value <> ""  then
			'if not isNumeric(Criteria.Confine.Value) then
				'msgbox "Please enter a valid Confine value.",,"Invalid Data"
				'exit sub
			'else
				'if Criteria.Confine.Value > <%= Application("SmlIntMax")%> or _
						'Criteria.Confine.Value < <%= Application("SmlIntMin")%> then
					'msgbox "Please enter a valid Confine value.",,"Invalid Data"
					'exit sub
				'end if
			'end if
			'bValidData = true
		'end if
		'if Criteria.ICD9_Proc.Value <> "" then
			'if ContainsInvalids(Criteria.ICD9_Proc.Value) then
				'msgbox "Please remove invalid characters from the ICD9 Procedure field.",,"Invalid Data"
				'exit sub
			'end if
			'bValidData = true
		'end if
		'if Criteria.DRG.Value <> ""  then
			'if not isNumeric(Criteria.DRG.Value) then
				'msgbox "Please enter a valid DRG value.",,"Invalid Data"
				'exit sub
			'else
				'if Criteria.DRG.Value > <%= Application("SmlIntMax")%> or _
						'Criteria.DRG.Value < <%= Application("SmlIntMin")%> then
					'msgbox "Please enter a valid DRG value.",,"Invalid Data"
					'exit sub
				'end if
			'end if
			'bValidData = true
		'end if
		'if Criteria.RevCode.Value <> ""  then
			'if not isNumeric(Criteria.RevCode.Value) then
				'msgbox "Please enter a valid Revenue Code.",,"Invalid Data"
				'exit sub
			'else
				'if Criteria.RevCode.Value > <%= Application("SmlIntMax")%> or _
						'Criteria.RevCode.Value < <%= Application("SmlIntMin")%> then
					'msgbox "Please enter a valid Revenue Code.",,"Invalid Data"
					'exit sub
				'end if
			'end if
			'bValidData = true
		'end if
'
' Other
'
		'if Criteria.DiagCode.Value <> "" then
			'if ContainsInvalids(Criteria.DiagCode.Value) then
				'msgbox "Please remove invalid characters from the Diagnosis field.",,"Invalid Data"
				'exit sub
			'end if
			'bValidData = true
		'end if
		if Criteria.ProcCode.Value <> "" then
			if ContainsInvalids(Criteria.ProcCode.Value) then
				msgbox "Please remove invalid characters from the Procedure Code field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		'if Criteria.Modifier.Value <> "" then
			'if ContainsInvalids(Criteria.Modifier.Value) then
				'msgbox "Please remove invalid characters from the Modifier field.",,"Invalid Data"
				'exit sub
			'end if
			'bValidData = true
		'end if
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
		'if Criteria.POSCode.Value <> ""  then
			'if not isNumeric(Criteria.POSCode.Value) then
				'msgbox "Please enter a valid Place Of Service.",,"Invalid Data"
				'exit sub
			'else
				'if Criteria.POSCode.Value > <%= Application("SmlIntMax")%> or _
						'Criteria.POSCode.Value < <%= Application("SmlIntMin")%> then
					'msgbox "Please enter a valid Place Of Service.",,"Invalid Data"
					'exit sub
				'end if
			'end if
			'bValidData = true
		'end if
		'if trim(Criteria.PayCode.Value) <> "" then
			'if ContainsInvalids(Criteria.PayCode.Value) then
				'msgbox "Please remove invalid characters from the Pay Code field.",,"Invalid Data"
				'exit sub
			'end if
			'bValidData = true
		'end if
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
		'if Criteria.PPO.Value <> "" then
			'if ContainsInvalids(Criteria.PPO.Value) then
				'msgbox "Please remove invalid characters from the PPO Sponsor field.",,"Invalid Data"
				'exit sub
			'end if
			'bValidData = true
		'end if
		
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
				if  "<%= m_dFromDate%>" <> Criteria.FromDate.Value or _
						"<%= m_dThruDate%>" <> Criteria.ThruDate.Value or _
						"<%= m_iSequence%>" <> Criteria.Sequence.Value or _
						"<%= m_iDistr%>" <> Criteria.Distr.Value or _
						"<%= m_iSSN%>" <> Criteria.SSN.Value or _
						"<%= m_iDepNum%>" <> Criteria.DepNum.Value or _
						"<%= m_iTaxID%>" <> Criteria.TaxID.Value or _
						"<%= m_sProvName%>" <> Criteria.ProvName.Value or _
						"<%= m_iAssocNum%>" <> Criteria.AssocNum.Value or _
						"<%= m_sStatus1%>" <> Criteria.Status1.Value or _
						"<%= m_sStatus2%>" <> Criteria.Status2.Value or _
						"<%= m_sStatus3%>" <> Criteria.Status3.Value or _
						"<%= m_dStatFrom%>" <> Criteria.StatFrom.Value or _
						"<%= m_dStatThru%>" <> Criteria.StatThru.Value or _
						"<%= m_sRSN%>" <> Criteria.RSN.Value or _
						"<%= m_dServiceFrom%>" <> Criteria.ServiceFrom.Value or _
						"<%= m_dServiceThru%>" <> Criteria.ServiceThru.Value or _
						"<%= m_sHWType%>" <> Criteria.HWType.Value or _
						"<%= m_sDiagCode%>" <> Criteria.DiagCode.Value or _
						"<%= m_sProcCode%>" <> Criteria.ProcCode.Value or _
						"<%= m_sTrackCode%>" <> Criteria.TrackCode.Value or _
            "<%= m_sCovCode%>" <> Criteria.CovCode.Value or _
						"<%= m_sPreCalCode%>" <> Criteria.PreCalCode.Value or _
						"<%= m_cCharge%>" <> Criteria.Charge.Value or _
						"<%= m_sCheckAcct%>" <> Criteria.CheckAcct.Value or _
						"<%= m_iCheckNum%>" <> Criteria.CheckNum.Value or _
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
