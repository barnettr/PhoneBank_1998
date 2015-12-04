<%@ Language=VBScript %>
<% Option Explicit %>
<% Response.Expires = 0 %>
<!--#include file="user.inc" -->
<!--#include file="adovbs_mod.inc" -->
<%
dim m_iSSN, adoConn, adoRS, sSQL, m_iDepNum, adoCmd, adoParam1, adoParam2
dim iPageNo, m_dFromDate, m_dThruDate, m_dStatusFrom, m_dStatusThru, m_iRxNumber, m_iNDC
dim adoParam3, adoParam4, adoParam5, adoParam6, adoParam7, adoParam8
dim m_sHWType, m_sFund, adoParam9, m_sStatus
dim Manager, Supervisor, Auditor

  m_iSSN = request.querystring("SSN")
  m_iDepNum = request.querystring("DepNum")
  
  Set adoConn = Server.CreateObject("ADODB.Connection")
  adoConn.ConnectionTimeout = 300
  adoConn.CommandTimeout = 300
  adoConn.Open Application("DataConn")

  Set adoCmd = Server.CreateObject("ADODB.Command")
  adoCmd.CommandTimeout = 300
  adoCmd.ActiveConnection = adoConn

  adoCmd.CommandType = &H0001
  adoCmd.CommandText = "select * from UserInformation where LogonID='" & Session("User") & "'"

  Set adoRS = Server.CreateObject("ADODB.Recordset")
  adoRS.Open adoCmd,,3,3

  if adoRS.state = 1 then
	  if adoRS.EOF then
		  norecords = 1
	  end if
  end if

  if not adoRS.EOF then
    Manager = adoRS("IsManager")
    Supervisor = adoRS("IsSupervisor")
    Auditor = adoRS("IsAuditor")
  end if
  adoRS.Close
  

  set adoCmd = Server.CreateObject("ADODB.Command")
  adoCmd.CommandTimeout = 0
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
  if request.querystring("SSN") <> "" then
    m_iSSN = request.querystring("SSN")
  end if
  if request.querystring("DepNum") <> "" then
    m_iDepNum = request.querystring("DepNum")
  else
    m_iDepNum = 99
  end if
  if Request.Querystring("FromDate") <> "" then
		m_dFromDate = Request.Querystring("FromDate")
	end if
	if Request.Querystring("ThruDate") <> "" then
		m_dThruDate = Request.Querystring("ThruDate")
	end if
	if Request.Querystring("StatusFrom") <> "" then
		m_dStatusFrom = Request.Querystring("StatusFrom")
	end if
	if Request.Querystring("StatusThru") <> "" then
		m_dStatusThru = Request.Querystring("StatusThru")
	end if  
	if Request.Querystring("RxNumber") <> "" then
		m_iRxNumber = Request.Querystring("RxNumber")
	end if
	if Request.Querystring("NDC") <> "" then
		m_iNDC = Request.Querystring("NDC")
	end if
  if Request.Querystring("Status") <> "" then
		m_sStatus = Request.Querystring("Status")
	end if
  
  adoCmd.CommandText = "PB_GetRxClaims"
  adoCmd.CommandType = adCmdStoredProc
  Set adoCmd.ActiveConnection = adoConn
  Set adoParam1 = adoCmd.CreateParameter("@P1_SSN", adInteger, adParamInput)
  adoCmd.Parameters.Append adoParam1
  Set adoParam2 = adoCmd.CreateParameter("@P2_DepNumber", adInteger, adParamInput)
  adoCmd.Parameters.Append adoParam2
  Set adoParam3 = adoCmd.CreateParameter("@P3_RxNumber", adInteger, adParamInput)
  adoCmd.Parameters.Append adoParam3
  Set adoParam4 = adoCmd.CreateParameter("@P4_NDC", adChar, adParamInput, 11)
  adoCmd.Parameters.Append adoParam4
  Set adoParam5 = adoCmd.CreateParameter("@P5_FromDate", adDBTimeStamp, adParamInput)
  adoCmd.Parameters.Append adoParam5
  Set adoParam6 = adoCmd.CreateParameter("@P6_ThruDate", adDBTimeStamp, adParamInput)
  adoCmd.Parameters.Append adoParam6
  Set adoParam7 = adoCmd.CreateParameter("@P7_StatusFrom", adDBTimeStamp, adParamInput)
  adoCmd.Parameters.Append adoParam7
  Set adoParam8 = adoCmd.CreateParameter("@P8_StatusThru", adDBTimeStamp, adParamInput)
  adoCmd.Parameters.Append adoParam8
  Set adoParam9 = adoCmd.CreateParameter("@P9_Status", adChar, adParamInput, 1)
  adoCmd.Parameters.Append adoParam9
  adoCmd("@P1_SSN") = m_iSSN
  adoCmd("@P2_DepNumber") = m_iDepNum
  adoCmd("@P3_RxNumber") = m_iRxNumber
  adoCmd("@P4_NDC") = m_iNDC
  adoCmd("@P5_FromDate") = m_dFromDate
  adoCmd("@P6_ThruDate") = m_dThruDate
  adoCmd("@P7_StatusFrom") = m_dStatusFrom
  adoCmd("@P8_StatusThru") = m_dStatusThru
  adoCmd("@P9_Status") = m_sStatus

%>
	
  <html>
	<head>
<!--#include file="VBFuncs.inc" -->
	</head>
	<body TOPMARGIN="2" LEFTMARGIN="2" RIGHTMARGIN="0" LANGUAGE="VBScript" onLoad="UpdateScreen(3)">
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
				<font SIZE="+2" face="verdana, arial, helvetica"><strong>Prescription Information</strong></font>
			</td>
			<td ALIGN="CENTER" WIDTH="20%">
				<img SRC="images/bluebar2.gif" onClick="history.back">&nbsp;&nbsp;<img src="images/bluebar3.gif" onClick="history.go(+1)" border="0">
			</td>
		</tr>
	</table>
<%
  adoRS.Open adoCmd
	if adoRS.EOF then
		Response.Write ("<p><ul><font color='red' face='verdana, arial, helvetica' size='2'><b>Unable to find any prescription information on this individual.</b></font></ul>")
	else
%>

<table COLS="2" WIDTH="100%" CELLPADDING="0" CELLSPACING="0" BORDER="0">
	<tr>
		<td WIDTH="10%"><input TYPE="IMAGE" SRC="images/query_button.gif" ACCESSKEY="S" NAME="NewQuery" VALUE="Start Query" onClick="ValidSend(0)"></td>
		<td WIDTH="65%"><font face="verdana, arial, helvetica" size="2"><b>Enter your criteria, and then click the Start Query button or key (ALT+S).</b></font></td>
  </tr>
</table>

<div STYLE="height=13%; width:100%; overflow=auto;">
			<form ID="Criteria" METHOD="GET" ACTION="Prescription.asp" TARGET="Details">
				<table CLASS="CriteriaTable" COLS="4" CELLPADDING="0" CELLSPACING="0" border="0">
					<tr>
<!-- These TH's were required in order to get IE to deal properly with the DIV...among other things, without them and with BORDER=1, things were displayed OK, butwith the desired BORDER=0 IE seemed to 'lose' a number of the elements-->
						<th>
						<th>
						<th>
						<th>
					</tr>
          <tr>
						<td class="White">SSN:</td>
						<td><input TYPE="TEXT" ID="SSN" NAME="SSN" SIZE="10" MAXLENGTH=10 VALUE="<%= m_iSSN%>"></td>
						<td class="White">Dependent Number:</td>
						<td><input TYPE="TEXT" ID="DepNum" NAME="DepNum" SIZE="2" MAXLENGTH=2 VALUE="<%= m_iDepNum%>"></td>
            <td class="White">Status:</td>
            <td><input TYPE="TEXT" ID="Status" NAME="Status" SIZE="2" MAXLENGTH=2 VALUE="<%= m_sStatus%>"></td>
					</tr>
					<tr>
						<td class="White">Rx Number:</td>
						<td><input TYPE="TEXT" ID="RxNumber" NAME="RxNumber" SIZE="10" MAXLENGTH=10 VALUE="<%= m_iRxNumber%>"></td>
						<td class="White">NDC:</td>
						<td><input TYPE="TEXT" ID="NDC" NAME="NDC" SIZE="20" MAXLENGTH=20 VALUE="<%= m_iNDC%>"></td>
					</tr>
					<tr>
						<td class="White">Rx Date Filled From:</td>
						<td><input TYPE="TEXT" ID="FromDate" NAME="FromDate" SIZE="10" MAXLENGTH=10 VALUE="<%= m_dFromDate%>"></td>
						<td class="White">Rx Date Filled Thru:</td>
						<td><input TYPE="TEXT" ID="ThruDate" NAME="ThruDate" SIZE="10" MAXLENGTH=10 VALUE="<%= m_dThruDate%>"></td>
					</tr>
          <tr>
						<td class="White">Status Date From:</td>
						<td><input TYPE="TEXT" ID="StatusFrom" NAME="StatusFrom" SIZE="10" MAXLENGTH=10 VALUE="<%= m_dStatusFrom%>"></td>
						<td class="White">Status Date Thru:</td>
						<td><input TYPE="TEXT" ID="StatusThru" NAME="StatusThru" SIZE="10" MAXLENGTH=10 VALUE="<%= m_dStatusThru%>"></td>
					</tr>
          <tr>
            <td colspan="4">&nbsp;</td>
          </tr>
				</table>
				<input TYPE="HIDDEN" ID="ActionType" NAME="ActionType" VALUE>
			</form>
		</div>
		<br CLEAR="LEFT">
<%  
      if adoRS("locktype") <> "" then
        if Manager = "True" or Supervisor = "True" or Auditor = "True" then 
%>  

<div style="height:68%; width:100%;">
		<p>&nbsp;
    <table WIDTH="100%" BORDER="1" bordercolor="white" bgcolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			<tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
				<td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>GRP</b></font></td>
        <td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>Tracking</b></font></td>
        <td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>Patient</b></font></td>
        <td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>SSN</b></font></td>
        <td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>RX #</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Date Filled</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Charge</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Paid</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Status</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Status Date</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Claim #</b></font></td>
      </tr>
<% 
			
			  Do While Not adoRS.EOF
%>

        <tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("Fund") %></td>
          <td nowrap bgcolor="#F0F0F0"><a href="Tracking.asp?SSN=<%= m_iSSN %>&DepNum=<%= m_iDepNum %>&HWType=P&Fund=<%= adoRS("Fund") %>&Years=<%= Year(Date) %>" onClick="WorkingStatus()">Tracking</a></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("Patient") %></td>
          <td nowrap bgcolor="#F0F0F0"><%= m_iSSN %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("RxNumber") %></td>
          <td nowrap bgcolor="#F0F0F0"><% if adoRS("DateFilled") <> "" then %><%= formatdatetime(adoRS("DateFilled"),vbShortDate) %><% else %>&nbsp;<% end if %></td>
          <td nowrap bgcolor="#F0F0F0"><% if adoRS("ChargeAmount") <> "" then %><%= formatcurrency(adoRS("ChargeAmount")) %><% else %>&nbsp;<% end if %></td>
          <td nowrap bgcolor="#F0F0F0"><% if adoRS("PaidAmount") <> "" then %><%= formatcurrency(adoRS("PaidAmount")) %><% else %>&nbsp;<% end if %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("Status") %></td>
          <td nowrap bgcolor="#F0F0F0"><% if adoRS("DateStatus") <> "" then %><%= formatdatetime(adoRS("DateStatus"),vbShortDate) %><% else %>&nbsp;<% end if %></td>
          <td nowrap bgcolor="#F0F0F0"><a href="PrescriptionDetails.asp?SSN=<%= m_iSSN %>&DepNum=<%= m_iDepNum %>&ClaimNumber=<%= adoRS("ClaimNumber") %>" onClick="WorkingStatus()"><%= adoRS("ClaimNumber") %></a></td>
        </tr>
        
<%
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
      else 'Admin, Manager, Auditor not true
%>

<div style="height:68%; width:100%;">
		<p>&nbsp;
    <table WIDTH="100%" BORDER="1" bordercolor="white" bgcolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			<tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
				<td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>GRP</b></font></td>
        <td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>Tracking</b></font></td>
        <td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>Patient</b></font></td>
        <td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>SSN</b></font></td>
        <td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>RX #</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Date Filled</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Charge</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Paid</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Status</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Status Date</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Claim #</b></font></td>
      </tr>
<% 
			
			  Do While Not adoRS.EOF
          if instr(adoRS("locktype"),adoRS("ClaimType")) = 0 then 
%>

        <tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("Fund") %></td>
          <td nowrap bgcolor="#F0F0F0"><a href="Tracking.asp?SSN=<%= m_iSSN %>&DepNum=<%= m_iDepNum %>&HWType=P&Fund=<%= adoRS("Fund") %>&Years=<%= Year(Date) %>" onClick="WorkingStatus()">Tracking</a></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("Patient") %></td>
          <td nowrap bgcolor="#F0F0F0"><%= m_iSSN %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("RxNumber") %></td>
          <td nowrap bgcolor="#F0F0F0"><% if adoRS("DateFilled") <> "" then %><%= formatdatetime(adoRS("DateFilled"),vbShortDate) %><% else %>&nbsp;<% end if %></td>
          <td nowrap bgcolor="#F0F0F0"><% if adoRS("ChargeAmount") <> "" then %><%= formatcurrency(adoRS("ChargeAmount")) %><% else %>&nbsp;<% end if %></td>
          <td nowrap bgcolor="#F0F0F0"><% if adoRS("PaidAmount") <> "" then %><%= formatcurrency(adoRS("PaidAmount")) %><% else %>&nbsp;<% end if %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("Status") %></td>
          <td nowrap bgcolor="#F0F0F0"><% if adoRS("DateStatus") <> "" then %><%= formatdatetime(adoRS("DateStatus"),vbShortDate) %><% else %>&nbsp;<% end if %></td>
          <td nowrap bgcolor="#F0F0F0"><a href="PrescriptionDetails.asp?SSN=<%= m_iSSN %>&DepNum=<%= m_iDepNum %>&ClaimNumber=<%= adoRS("ClaimNumber") %>" onClick="WorkingStatus()"><%= adoRS("ClaimNumber") %></a></td>
        </tr>
        
<%
			    else
            response.write "<tr bordercolordark='#FCFCFC' bordercolorlight='#CCCCCC'><td nowrap bgcolor='#F0F0F0' colspan='11'><center><b>Data is not available because file is locked---contact your Administrator!</b></center></td></tr>" 
          end if
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
      end if
    else 'adoRS("locktype") is empty
%>

<div style="height:68%; width:100%;">
		<p>&nbsp;
    <table WIDTH="100%" BORDER="1" bordercolor="white" bgcolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			<tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
				<td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>GRP</b></font></td>
        <td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>Tracking</b></font></td>
        <td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>Patient</b></font></td>
        <td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>SSN</b></font></td>
        <td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>RX #</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Date Filled</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Charge</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Paid</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Status</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Status Date</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Claim #</b></font></td>
      </tr>
<% 
			
			  Do While Not adoRS.EOF
%>

        <tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("Fund") %></td>
          <td nowrap bgcolor="#F0F0F0"><a href="Tracking.asp?SSN=<%= m_iSSN %>&DepNum=<%= m_iDepNum %>&HWType=P&Fund=<%= adoRS("Fund") %>&Years=<%= Year(Date) %>" onClick="WorkingStatus()">Tracking</a></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("Patient") %></td>
          <td nowrap bgcolor="#F0F0F0"><%= m_iSSN %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("RxNumber") %></td>
          <td nowrap bgcolor="#F0F0F0"><% if adoRS("DateFilled") <> "" then %><%= formatdatetime(adoRS("DateFilled"),vbShortDate) %><% else %>&nbsp;<% end if %></td>
          <td nowrap bgcolor="#F0F0F0"><% if adoRS("ChargeAmount") <> "" then %><%= formatcurrency(adoRS("ChargeAmount")) %><% else %>&nbsp;<% end if %></td>
          <td nowrap bgcolor="#F0F0F0"><% if adoRS("PaidAmount") <> "" then %><%= formatcurrency(adoRS("PaidAmount")) %><% else %>&nbsp;<% end if %></td>
          <td nowrap bgcolor="#F0F0F0"><%= adoRS("Status") %></td>
          <td nowrap bgcolor="#F0F0F0"><% if adoRS("DateStatus") <> "" then %><%= formatdatetime(adoRS("DateStatus"),vbShortDate) %><% else %>&nbsp;<% end if %></td>
          <td nowrap bgcolor="#F0F0F0"><a href="PrescriptionDetails.asp?SSN=<%= m_iSSN %>&DepNum=<%= m_iDepNum %>&ClaimNumber=<%= adoRS("ClaimNumber") %>" onClick="WorkingStatus()"><%= adoRS("ClaimNumber") %></a></td>
        </tr>
        
<%
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
  
</body>
<!--#include file="VBFuncs.inc" -->
<script LANGUAGE="VBSCRIPT">
	
	sub ValidSend(iIndex)
	dim bValidData
	dim vbOKCancel, vbOk
	
		'vbOkCancel = 1
		'vbOk = 1

		Criteria.SSN.Value = trim(Criteria.SSN.Value)
    Criteria.DepNum.Value = trim(Criteria.DepNum.Value)
    Criteria.RxNumber.Value = trim(Criteria.RxNumber.Value)
		Criteria.NDC.Value = trim(Criteria.NDC.Value)
		Criteria.FromDate.Value = trim(Criteria.FromDate.Value)
		Criteria.ThruDate.Value = trim(Criteria.ThruDate.Value)
    Criteria.StatusFrom.Value = trim(Criteria.StatusFrom.Value)
		Criteria.StatusThru.Value = trim(Criteria.StatusThru.Value)
    
    'if iIndex = 0 and (Criteria.DepNum.Value = "" then
			'msgbox "You must enter a Dependent Number for your query to run!",16,"Missing Data"
      'exit sub
		'end if

		
		WorkingStatus
		if "<%= m_iSSN%>" <> Criteria.SSN.Value or _
        "<%= m_iDepNum%>" <> Criteria.DepNum.Value or _
        "<%= m_iRxNumber%>" <> Criteria.RxNumber.Value or _
				"<%= m_iNDC%>" <> Criteria.NDC.Value or _
				"<%= m_dFromDate%>" <> Criteria.FromDate.Value or _
				"<%= m_dThruDate%>" <> Criteria.ThruDate.Value or _
        "<%= m_dStatusFrom%>" <> Criteria.StatusFrom.Value or _
				"<%= m_dStatusThru%>" <> Criteria.StatusThru.Value then
			Criteria.ActionType.value=""
		else
			Criteria.ActionType.value=iIndex
		end if
		Criteria.submit
	end sub
	
</script>

</html>