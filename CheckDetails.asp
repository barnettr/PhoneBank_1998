<%@ Language=VBScript %>
<% Option Explicit %>
<% Response.Expires = 0 %>
<!--#include file="user.inc" -->
<!--#include file="adovbs_mod.inc" -->
<%
Dim i, adoConn, adoRS, sSQL, sTemp
Dim m_iCheckNum, m_sCheckAcct
Dim adoCmd, adoParam1, adoParam2, m_sHWType


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

function PayCodeDesc(sPayCodeChar)
dim sStatus

	Select case ucase(sPayCodeChar)
		case "J"
			sStatus = "Joint"
		case "N"
			sStatus = "Pharmacy"
		case "O"
			sStatus = "Insured"
		case "P"
			sStatus = "Provider"
		case "T"
			sStatus = "Third Party"
	end select
	PayCodeDesc = sPayCodeChar & "-" & sStatus
end function

function StatusDesc(sStatusChar)
dim sStatus

	Select case ucase(sStatusChar)
		case "B"
			sStatus = "Block Void"
		case "E"
			sStatus = "Electronic Funds Transfer"
		case "I"
			sStatus = "Issued"
		case "P"
			sStatus = "Paid"
		case "R"
			sStatus = "Reopened"
		case "S"
			sStatus = "Stop Pay"
		case "T"
			sStatus = "Stop Request"
		case "V"
			sStatus = "Void"
	end select
	StatusDesc = sStatusChar & "-" & sStatus
end function

function HWDesc(sHWChar)
dim sStatus

	Select case ucase(sHWChar)
		case "D"
			sStatus = "Dental"
		case "M"
			sStatus = "Medical"
		case "O"
			sStatus = "????"
		case "P"
			sStatus = "Prescription Drugs"
		case "T"
			sStatus = "Timeloss"
	end select
	HWDesc = sHWChar & "-" & sStatus
end function

m_iCheckNum = Request.QueryString("CheckNum")
m_sCheckAcct = Request.QueryString("CheckAcct")

if m_iCheckNum = "" or m_sCheckAcct = "" then
	Response.Write "<p><font color='red' size='2' face='verdana, arial, helvetica'><b>Check Account or Check Number missing; contact your network administrator</b></font>"
else
	set adoConn = Server.CreateObject("ADODB.Connection")
	adoConn.CommandTimeout = 0
  set adoCmd = Server.CreateObject("ADODB.Command")
	set adoRS = Server.CreateObject("ADODB.Recordset")
	adoConn.Open Application("DataConn")
  
  adoCmd.CommandText = "pb_CheckDetails"
	adoCmd.CommandType = adCmdStoredProc
	Set adoCmd.ActiveConnection = adoConn
  Set adoParam1 = adoCmd.CreateParameter("@P1_Acct", adVarChar, adParamInput, 6)
	adoCmd.Parameters.Append adoParam1
  Set adoParam2 = adoCmd.CreateParameter("@P2_Number", adInteger, adParamInput)
  adoCmd.Parameters.Append adoParam2
  adoCmd("@P1_Acct") = m_sCheckAcct
  adoCmd("@P2_Number") = m_iCheckNum
%>
	<html>
	<head>
<!--#include file="VBFuncs.inc" -->
	</head>
	<body LANGUAGE="VBScript" onLoad="UpdateScreen(2)">
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
				<font SIZE="+2" face="verdana, arial, helvetica"><strong>Check Information</strong></font>
			</td>
			<td ALIGN="CENTER" WIDTH="20%">
				<img SRC="images/bluebar2.gif" onClick="history.back">&nbsp;&nbsp;<img src="images/bluebar3.gif" onClick="history.go(+1)" border="0">
			</td>
		</tr>
	</table>
<%
	adoRS.Open adoCmd
	if adoRS.EOF then
		Response.Write ("<p><ul><font color='red' face='verdana, arial, helvetica'>Unable to find all necessary information; please contact your network administrator.</font></ul>")
	else
    m_sHWType = adoRS("HWType")
%>
    <br CLEAR="LEFT">
<!--		<table BORDER="2" WIDTH="100%" CELLPADDING="3" CELLSPACING="3">-->
		<table border="1" width="100%" COLS=4 RULES=GROUPS bgcolor="white" bordercolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			<COLGROUP SPAN=2>
			<COLGROUP SPAN=2>
		    <tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
				  <td WIDTH="25%" bgcolor="#cccccc"><strong>Check Account:</strong></td>
				  <td WIDTH="25%" STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#cccccc"><%= adoRS("Account")%></td>
		      <td WIDTH="25%" bgcolor="#cccccc"><strong>Amount:</strong></td>
		      <td WIDTH="25%" STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#cccccc"><%= formatcurrency(adoRS("Amount"))%></td>
		    </tr>
        <tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
		      <td bgcolor="#F0F0F0"><strong>Check Number:</strong></td>
		      <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= adoRS("Check #")%></td>
		      <td bgcolor="#F0F0F0"><strong>Issue Date:</strong></td>
		      <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= adoRS("IssueDate")%></td>
		    </tr>
		    <tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
		      <td bgcolor="#cccccc"><strong>Check Status:</strong></td>
		      <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#cccccc"><%= StatusDesc(adoRS("Status"))%></td>
		      <td bgcolor="#cccccc"><strong>Status Date:</strong></td>
		      <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#cccccc"><%= adoRS("StatusDate")%></td>
		    </tr>
		</table>
		<BR>
		<table WIDTH="100%" CELLPADDING="3" CELLSPACING="3" bgcolor="white" bordercolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
				<td WIDTH="20%" bgcolor="#F0F0F0"><strong>Issued to:</strong></td>
		    <td WIDTH="30%" STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><a href="ProviderSearch.asp?FullName=<%= adoRS("PayeeName") %>&TaxID=<%= adoRS("Payee") %>" onClick="WorkingStatus()"><%= adoRS("PayeeName")%></a></td>
		    <td WIDTH="20%" bgcolor="#F0F0F0"><strong>Group:</strong></td>
		    <td WIDTH="30%" STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= adoRS("Fund")%></td>
			</tr>
			<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
				<td bgcolor="#F0F0F0">&nbsp;</td>
		    <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= adoRS("Address")%></td>
		    <td bgcolor="#F0F0F0"><strong>HWType:</strong></td>
		    <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= HWDesc(adoRS("HWType"))%></td>
			</tr>
		  <tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
				<td bgcolor="#F0F0F0">&nbsp;</td>
		    <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= adoRS("Location")%></td>
		    <td bgcolor="#F0F0F0"><strong>Employer ID:</strong></td>
		    <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= adoRS("Employer")%></td>
		  </tr>
		  <tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
		    <td bgcolor="#F0F0F0"><strong>Payee ID:</strong></td>
		    <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= adoRS("Payee")%></td>
		    <td bgcolor="#F0F0F0"><strong>Employer Name:</strong></td>
		    <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= adoRS("Name")%></td>
		  </tr>
		  <tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
		    <td bgcolor="#F0F0F0"><strong>Associate #:</strong></td>
		    <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= adoRS("Associate")%></td>
		    <td bgcolor="#F0F0F0"><strong>Pay Code</strong></td>
		    <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= PayCodeDesc(adoRS("PayCode"))%></td>
		  </tr>
		</table>
    <br>
		<hr SIZE="3" NOSHADE>
<%
		adoRS.Close
'
' Claims
'
  if m_sHWType = "P" then
%>
		<br>
    <font SIZE="3" face="arial, helvetica"><center><b>Rx Claims Information</b></center></font>
    <table WIDTH="100%" BORDER="1" CELLPADDING="2" CELLSPACING="2" bgcolor="white" bordercolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
      <tr ALIGN="CENTER" BorderColorDark="#F0F0F0" BorderColorlight="#999999">
        <td bgcolor="#cccccc"><strong>Claim #</strong></td>
        <td bgcolor="#cccccc"><strong>Sequence</strong></td>
        <td bgcolor="#cccccc"><strong>SSN</strong></td>
        <td bgcolor="#cccccc"><strong>Dep Number</strong></td>
        <td bgcolor="#cccccc"><strong>Rx Number</strong></td>
        <td bgcolor="#cccccc"><strong>Date Filled</strong></td>
        <td bgcolor="#cccccc"><strong>Charge Amount</strong></td>
        <td bgcolor="#cccccc"><strong>Paid Amount</strong></td>
      </tr>
<% 
    sSQL = "pb_RxCheckClaimDetails '" & m_sCheckAcct & "', " & m_iCheckNum
    Set adoRS = adoConn.execute(sSQL)
    if not adoRS.EOF then
				Do While Not adoRS.EOF
      
%>    
      <tr ALIGN="CENTER" BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
        <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><a href="PrescriptionDetails.asp?SSN=<%= adoRS("SSN") %>&DepNum=<%= adoRS("DepNumber") %>&ClaimNumber=<%= adoRS("ClaimNumber") %>" onClick="WorkingStatus()"><%= adoRS("ClaimNumber") %></a></td>
        <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= adoRS("ClaimSequence") %></td>
        <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= adoRS("SSN") %></td>
        <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= adoRS("DepNumber") %></td>
        <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><%= adoRS("RxNumber") %></td>
        <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><% if adoRS("DateFilled") <> "" then %><%= formatdatetime(adoRS("DateFilled"),vbShortDate) %><% else %>&nbsp;<% end if %></td>
        <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><% if adoRS("ChargeAmount") <> "" then %><%= formatcurrency(adoRS("ChargeAmount")) %><% else %>&nbsp;<% end if %></td>
        <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><% if adoRS("PaidAmount") <> "" then %><%= formatcurrency(adoRS("PaidAmount")) %><% else %>&nbsp;<% end if %></td>
      </tr>
    
<%  
    	adoRS.MoveNext
		Loop
    adoRS.Close
    
    end if
  end if

    'if m_sHWType <> "P" then
%>
  </table>  
    <br>	
		<font SIZE="3" face="arial, helvetica"><center><b>Claims Information</b></center></font>
		<table WIDTH="100%" BORDER="1" CELLPADDING="2" CELLSPACING="2" bgcolor="white" bordercolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			<tr ALIGN="CENTER" BorderColorDark="#F0F0F0" BorderColorlight="#999999">
			    <td bgcolor="#cccccc"><strong>Claim #</strong></td>
			    <td bgcolor="#cccccc"><strong>Line #</strong></td>
			    <td bgcolor="#cccccc"><strong>SSN</strong></td>
			    <td bgcolor="#cccccc"><strong>Partic. Name</strong></td>
			    <td bgcolor="#cccccc">&nbsp;</td>
			    <td bgcolor="#cccccc"><strong>Dep #</strong></td>
			    <td bgcolor="#cccccc"><strong>Dep. Name</strong></td>
			    <td bgcolor="#cccccc"><strong>Cov. Code</strong></td>
			    <td bgcolor="#cccccc"><strong>From</strong></td>
			    <td bgcolor="#cccccc"><strong>Thru</strong></td>
			    <td bgcolor="#cccccc"><strong>Charge</strong></td>
			    <td bgcolor="#cccccc"><strong>Paid</strong></td>
			</tr>
<%
      sSQL = "pb_CheckClaimDetails '" & m_sCheckAcct & "', " & m_iCheckNum
      Set adoRS = adoConn.execute(sSQL)
			if not adoRS.EOF then
				Do While Not adoRS.EOF
%>
					<tr ALIGN="CENTER" BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
					  <td ALIGN="LEFT" NOWRAP bgcolor="#F0F0F0"><a HREF="aspexec.asp?ClaimNum=<%= ConvertToClaimNum(adoRS("LogDate"),adoRS("Sequence"),adoRS("Distr")) %>" onClick="WorkingStatus()"><%= ConvertToClaimNum(adoRS("LogDate"),adoRS("Sequence"),adoRS("Distr"))%></a></td>
					  <td bgcolor="#F0F0F0"><%= adoRS("Line #") %></td>
					  <td bgcolor="#F0F0F0"><%= adoRS("SSN") %></td>
<%
						sTemp = "<TD ALIGN=LEFT bgcolor='#F0F0F0'><A HREF='PersonDetails.asp?SSN=" & adoRS("SSN") & "&DepNum=0&locktype=" & adoRS("ParticipantLockType") & "&Admin=" & adoRS("ParticipantAdmin")
						sTemp = sTemp & "' onClick='javascript:return LogCheck(" & adoRS("SSN") & ")'>"
						sTemp = sTemp & adoRS("ParticName") & "</A></TD>"
						Response.Write sTemp
%>					    
					  <td bgcolor="#F0F0F0"><input TYPE="BUTTON" ID="Search" VALUE="Calls" onClick="ShowCalls(<%=adoRS("SSN")%>)"></td>
					  <td bgcolor="#F0F0F0"><%= adoRS("DepNumber") %></td>
<%
						if adoRS("DepNumber") <> 0 then
							sTemp = "<TD ALIGN=LEFT bgcolor='#F0F0F0'><A HREF='PersonDetails.asp?SSN=" & adoRS("SSN") & "&DepNum=" & adoRS("DepNumber") & "&locktype=" & adoRS("locktype") & "&Admin=" & adoRS("SupervisorAccessOnly")
							sTemp = sTemp & "' onClick='javascript:return LogCheck(" & adoRS("SSN") & ")'>"
							sTemp = sTemp & adoRS("DepName") & "</A></TD>"
						else
							sTemp = "<TD bgcolor='#F0F0F0'>&nbsp;</TD>"
						end if
						Response.Write sTemp
%>					    
					    <td bgcolor="#F0F0F0"><%= adoRS("CovCode") %></td>
					    <td bgcolor="#F0F0F0"><% if adoRS("FromDate") <> "" then %><%= formatdatetime(adoRS("FromDate")) %><% else %>&nbsp;<%= adoRS("FromDate") %><% end if %></td>
					    <td bgcolor="#F0F0F0"><% if adoRS("ThruDate") <> "" then %><%= formatdatetime(adoRS("ThruDate")) %><% else %>&nbsp;<%= adoRS("ThruDate") %><% end if %></td>
					    <td bgcolor="#F0F0F0"><% if adoRS("ChargeAmount") <> "" then %><%= formatcurrency(adoRS("ChargeAmount")) %><% else %>&nbsp;<%= adoRS("ChargeAmount") %><% end if %></td>
					    <td bgcolor="#F0F0F0"><% if adoRS("PaidAmount") <> "" then %><%= formatcurrency(adoRS("PaidAmount")) %><% else %>&nbsp;<%= adoRS("PaidAmount") %><% end if %></td>
					</tr>
<% 
					adoRS.MoveNext
				Loop 
			else
%>
				<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
				    <td COLSPAN="12" bgcolor="#F0F0F0"><font color="red" face="verdana, arial, helvetica" size="2"><b>No Claims found.</b></font></td>
				</tr>
<%
			end if
%>
		</table>
<%
		adoRS.Close
		adoConn.Close
		set adoRS = nothing
		set adoConn = nothing
    end if
   
%>	
		<form ID="Criteria" METHOD="GET" ACTION="PhoneSearch.asp" TARGET="Details">
			<input TYPE="HIDDEN" ID="SSN" NAME="SSN" VALUE>
		</form>
		<script LANGUAGE="VBScript">
			
			sub ShowCalls(iSSN)
			
				WorkingStatus
				Criteria.SSN.value = iSSN
				Criteria.submit
				
			end sub
			
		</script>
		</body>
		</html>
<%
	end if


%>