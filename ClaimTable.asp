<% 
'****************************************************************************************************
'* This is the origional table with all the fields from the ClaimSearch.asp page. If any of these
'* fields needs to be added back for search reasons it should be a fairly simple process.
'*
'* You will also find alot of code that is commented out on the ClaimSearch.asp page. For any field that
'* you add back into the table the corresponding code that error-checks that particular field will have
'* to be uncommented or you will get run-time errors.
'*
'* You also need to check NeedToSave.asp where the origional conditional statement at the end of ClaimSearch.asp
'* resides. These could not be commented out because of the ASP tags. They had to be removed and if you add
'* any fields back to the table the corresponding field must be added back into this statement or the page foreward
'* by number will break.
'*
'* Rob Barnett 
'* September 20, 1999
'*
'*
'*********************************************************************************************************
 %>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Claim Table (Old)</title>
</head>

<body>
 
 	<table CLASS="CriteriaTable" COLS="8" CELLPADDING="0" CELLSPACING="0">
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
				<td ALIGN="RIGHT" class="White">SSN:</td>
				<td><input TYPE="TEXT" ID="SSN" NAME="SSN" SIZE="10" MAXLENGTH=10 VALUE="<%= m_iSSN%>"></td>
				<td ALIGN="RIGHT" class="White">Last Name:</td>
				<td><input TYPE="TEXT" ID="LName" NAME="LName" SIZE="15" MAXLENGTH=20 VALUE="<%= m_sLastName%>"></td>
				<td ALIGN="RIGHT" class="White">Dep. #:</td>
				<td><input TYPE="TEXT" ID="DepNum" NAME="DepNum" SIZE="3" MAXLENGTH=5 VALUE="<%= m_iDepNum%>"></td>
				<td COLSPAN="2">&nbsp;</td>
			</tr>
			<tr>
				<td ALIGN="RIGHT" class="White">Provider Tax ID:</td>
				<td><input TYPE="TEXT" ID="TaxID" NAME="TaxID" SIZE="10" MAXLENGTH=10 VALUE="<%= m_iTaxID%>"></td>
				<td ALIGN="RIGHT" class="White">Provider Name:</td>
				<td><input TYPE="TEXT" ID="ProvName" NAME="ProvName" SIZE="15" MAXLENGTH=40 VALUE="<%= m_sProvName%>"></td>
				<td ALIGN="RIGHT" class="White">Assoc. #:</td>
				<td><input TYPE="TEXT" ID="AssocNum" NAME="AssocNum" SIZE="3" MAXLENGTH=10 VALUE="<%= m_iAssocNum%>"></td>
				<td COLSPAN="2">&nbsp;</td>
			</tr>
			<tr>
				<td ALIGN="RIGHT" class="White">Fund:</td>
				<td><input TYPE="TEXT" ID="Fund" NAME="Fund" SIZE="5" MAXLENGTH=3 VALUE="<%= m_sFund%>"></td>
				<td ALIGN="RIGHT" class="White">Status:</td>
				<td>
					<input TYPE="TEXT" ID="Status1" NAME="Status1" SIZE="2" MAXLENGTH=1 VALUE="<%= m_sStatus1%>">
					<input TYPE="TEXT" ID="Status2" NAME="Status2" SIZE="2" MAXLENGTH=1 VALUE="<%= m_sStatus2%>">
					<input TYPE="TEXT" ID="Status3" NAME="Status3" SIZE="2" MAXLENGTH=1 VALUE="<%= m_sStatus3%>">
				</td>
				<td ALIGN="RIGHT" class="White">Status Date From:</td>
				<td><input TYPE="TEXT" ID="StatFrom" NAME="StatFrom" SIZE="10" MAXLENGTH=10 VALUE="<%= m_dStatFrom%>"></td>
				<td ALIGN="RIGHT" class="White">Status Date Thru:</td>
				<td><input TYPE="TEXT" ID="StatThru" NAME="StatThru" SIZE="10" MAXLENGTH=10 VALUE="<%= m_dStatThru%>"></td>
			</tr>
			<tr>
				<td ALIGN="RIGHT" class="White">Plan:</td>
				<td><input TYPE="TEXT" ID="Plan" NAME="Plan" SIZE="8" MAXLENGTH=6 VALUE="<%= m_sPlan%>"></td>
				<td ALIGN="RIGHT" class="White">Pend/Deny Reason:</td>
				<td><input TYPE="TEXT" ID="RSN" NAME="RSN" SIZE="5" MAXLENGTH=4 VALUE="<%= m_sRSN%>"></td>
				<td ALIGN="RIGHT" class="White">Claim Service Date From:</td>
				<td><input TYPE="TEXT" ID="ServiceFrom" NAME="ServiceFrom" SIZE="10" MAXLENGTH=10 VALUE="<%= m_dServiceFrom%>"></td>
				<td ALIGN="RIGHT" class="White">Claim Service Date Thru:</td>
				<td><input TYPE="TEXT" ID="ServiceThru" NAME="ServiceThru" SIZE="10" MAXLENGTH=10 VALUE="<%= m_dServiceThru%>"></td>
			</tr>
			<tr>
				<td ALIGN="RIGHT" class="White">HWType:</td>
				<td><input TYPE="TEXT" ID="HWType" NAME="HWType" SIZE="3" MAXLENGTH=1 VALUE="<%= m_sHWType%>"></td>
				<td ALIGN="RIGHT" class="White">Benefit Period:</td>
				<td><input TYPE="TEXT" ID="BenPeriod" NAME="BenPeriod" SIZE="5" MAXLENGTH=5 VALUE="<%= m_iBenPeriod%>"></td>
				<td ALIGN="RIGHT" class="White">Benefit Period Date From:</td>
				<td><input TYPE="TEXT" ID="BenFrom" NAME="BenFrom" SIZE="10" MAXLENGTH=10 VALUE="<%= m_dBenFrom%>"></td>
				<td ALIGN="RIGHT" class="White">Benefit Period Date Thru:</td>
				<td><input TYPE="TEXT" ID="BenThru" NAME="BenThru" SIZE="10" MAXLENGTH=10 VALUE="<%= m_dBenThru%>"></td>
			</tr>
			<tr>
				<td ALIGN="RIGHT" class="White">Admit Date From:</td>
				<td><input TYPE="TEXT" ID="AdmitFrom" NAME="AdmitFrom" SIZE="10" MAXLENGTH=10 VALUE="<%= m_dAdmitFrom%>"></td>
				<td ALIGN="RIGHT" class="White">Admit Date Thru:</td>
				<td><input TYPE="TEXT" ID="AdmitThru" NAME="AdmitThru" SIZE="10" MAXLENGTH=10 VALUE="<%= m_dAdmitThru%>"></td>
				<td ALIGN="RIGHT" class="White">Discharge Date From:</td>
				<td><input TYPE="TEXT" ID="DisFrom" NAME="DisFrom" SIZE="10" MAXLENGTH=10 VALUE="<%= m_dDisFrom%>"></td>
				<td ALIGN="RIGHT" class="White">Discharge Date Thru:</td>
				<td><input TYPE="TEXT" ID="DisThru" NAME="DisThru" SIZE="10" MAXLENGTH=10 VALUE="<%= m_dDisThru%>"></td>
			</tr>
			<tr>
				<td ALIGN="RIGHT" class="White">Admit Type:</td>
				<td><input TYPE="TEXT" ID="AdmitType" NAME="AdmitType" SIZE="3" MAXLENGTH=1 VALUE="<%= m_sAdmitType%>"></td>
				<td ALIGN="RIGHT" class="White">Bill Type:</td>
				<td><input TYPE="TEXT" ID="BillType" NAME="BillType" SIZE="5" MAXLENGTH=5 VALUE="<%= m_iBillType%>"></td>
				<td ALIGN="RIGHT" class="White">Bill Date From:</td>
				<td><input TYPE="TEXT" ID="BillFrom" NAME="BillFrom" SIZE="10" MAXLENGTH=10 VALUE="<%= m_dBillFrom%>"></td>
				<td ALIGN="RIGHT" class="White">Bill Date Thru:</td>
				<td><input TYPE="TEXT" ID="BillThru" NAME="BillThru" SIZE="10" MAXLENGTH=10 VALUE="<%= m_dBillThru%>"></td>
			</tr>
			<tr>
				<td ALIGN="RIGHT" class="White">Confine Type:</td>
				<td><input TYPE="TEXT" ID="Confine" NAME="Confine" SIZE="5" MAXLENGTH=5 VALUE="<%= m_iConfine%>"></td>
				<td ALIGN="RIGHT" class="White">ICD9 Proc. Code:</td>
				<td><input TYPE="TEXT" ID="ICD9_Proc" NAME="ICD9_Proc" SIZE="6" MAXLENGTH=5 VALUE="<%= m_iICD9_Proc%>"></td>
				<td ALIGN="RIGHT" class="White">DRG:</td>
				<td><input TYPE="TEXT" ID="DRG" NAME="DRG" SIZE="5" MAXLENGTH=5 VALUE="<%= m_iDRG%>"></td>
				<td ALIGN="RIGHT" class="White">Revenue Code:</td>
				<td><input TYPE="TEXT" ID="RevCode" NAME="RevCode" SIZE="5" MAXLENGTH=5 VALUE="<%= m_iRevCode%>"></td>
			</tr>
			<tr>
				<td ALIGN="RIGHT" class="White">Diagnosis:</td>
				<td><input TYPE="TEXT" ID="DiagCode" NAME="DiagCode" SIZE="8" MAXLENGTH=6 VALUE="<%= m_sDiagCode%>"></td>
				<td ALIGN="RIGHT" class="White">Procedure Code:</td>
				<td><input TYPE="TEXT" ID="ProcCode" NAME="ProcCode" SIZE="7" MAXLENGTH=5 VALUE="<%= m_sProcCode%>"></td>
				<td ALIGN="RIGHT" class="White">Modifier:</td>
				<td><input TYPE="TEXT" ID="Modifier" NAME="Modifier" SIZE="4" MAXLENGTH=2 VALUE="<%= m_sModifier%>"></td>
				<td ALIGN="RIGHT" class="White">Track Code:</td>
				<td><input TYPE="TEXT" ID="TrackCode" NAME="TrackCode" SIZE="5" MAXLENGTH=3 VALUE="<%= m_sTrackCode%>"></td>
			</tr>
			<tr>
				<td ALIGN="RIGHT" class="White">Coverage Code:</td>
				<td><input TYPE="TEXT" ID="CovCode" NAME="CovCode" SIZE="2" MAXLENGTH=2 VALUE="<%= m_sCovCode%>"></td>
				<td ALIGN="RIGHT" class="White">Pre Cal Code:</td>
				<td><input TYPE="TEXT" ID="PreCalCode" NAME="PreCalCode" SIZE="4" MAXLENGTH=2 VALUE="<%= m_sPreCalCode%>"></td>
				<td ALIGN="RIGHT" class="White">Place of Service Code:</td>
				<td><input TYPE="TEXT" ID="POSCode" NAME="POSCode" SIZE="5" MAXLENGTH=5 VALUE="<%= m_iPOSCode%>"></td>
				<td ALIGN="RIGHT" class="White">Pay To Code:</td>
				<td><input TYPE="TEXT" ID="PayCode" NAME="PayCode" SIZE="3" MAXLENGTH=1 VALUE="<%= m_sPayCode%>"></td>
			</tr>
			<tr>
				<td ALIGN="RIGHT" class="White">Charge:</td>
				<td><input TYPE="TEXT" ID="Charge" NAME="Charge" SIZE="8" MAXLENGTH=11 VALUE="<%= m_cCharge%>"></td>
				<td ALIGN="RIGHT" class="White">Check Account:</td>
				<td><input TYPE="TEXT" ID="CheckAcct" NAME="CheckAcct" SIZE="8" MAXLENGTH=6 VALUE="<%= m_sCheckAcct%>"></td>
				<td ALIGN="RIGHT" class="White">Check Number:</td>
				<td><input TYPE="TEXT" ID="CheckNum" NAME="CheckNum" SIZE="7" MAXLENGTH=10 VALUE="<%= m_iCheckNum%>"></td>
				<td COLSPAN="2"></td>
			</tr>
			<tr>
			</tr>
			<tr>
				<td ALIGN="RIGHT" class="White">PPO Code:</td>
				<td><input TYPE="TEXT" ID="PPO" NAME="PPO" SIZE="5" MAXLENGTH=3 VALUE="<%= m_sPPO%>"></td>
				<td ALIGN="RIGHT" class="White">PPO Override:</td>
				<td><input TYPE="CHECKBOX" ID="PPOOverride" NAME="PPOOverride" <% if m_bPPOCheck then Response.Write " CHECKED " end if %>></td>
				<td ALIGN="RIGHT" class="White">COB:</td>
				<td><input TYPE="CHECKBOX" ID="COB" NAME="COB" <% if m_bCOBCheck then Response.Write " CHECKED " end if %>></td>
				<td ALIGN="RIGHT" class="White">COB Override:</td>
				<td><input TYPE="CHECKBOX" ID="COBOver" NAME="COBOver" <% if m_bCOBOverCheck then Response.Write " CHECKED " end if %>></td>
			</tr>
			<tr>
				<td ALIGN="RIGHT" class="White">Overpayment Adjustments:</td>
				<td><input TYPE="CHECKBOX" ID="OverPay" NAME="OverPay" <% if m_bOverPayCheck then Response.Write " CHECKED " end if %>></td>
				<td ALIGN="RIGHT" class="White">Non-Overpayment Adjustments:</td>
				<td><input TYPE="CHECKBOX" ID="NonOverPay" NAME="NonOverPay" <% if m_bOtherAdjustCheck then Response.Write " CHECKED " end if %>></td>
				<td COLSPAN="4"></td>
			</tr>
		</table>


</body>
</html>
