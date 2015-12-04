<%@ Language=VBScript %>
<% Option Explicit %>
<% Response.Expires = 0 %>
<!--#include file="user.inc" -->
<!--#include file="adovbs_mod.inc" -->
<%
dim m_iSSN, adoConn, adoRS, sSQL, m_iDepNum, m_iClaimNumber, adoCmd
dim adoParam1, adoParam2, adoParam3

  m_iSSN = request.querystring("SSN")
  m_iDepNum = request.querystring("DepNum")
  m_iClaimNumber = request.querystring("ClaimNumber")
  
  set adoConn = Server.CreateObject("ADODB.Connection")
  adoConn.ConnectionTimeout = 0
	adoConn.CommandTimeout = 0
  set adoCmd = Server.CreateObject("ADODB.Command")
  adoCmd.CommandTimeout = 0
	set adoRS = Server.CreateObject("ADODB.Recordset")
	adoConn.Open Application("DataConn")
  
  adoCmd.CommandText = "PB_RxListClaimsDetails"
  adoCmd.CommandType = adCmdStoredProc
  Set adoCmd.ActiveConnection = adoConn
  Set adoParam1 = adoCmd.CreateParameter("@P1_SSN", adInteger, adParamInput)
  adoCmd.Parameters.Append adoParam1
  Set adoParam2 = adoCmd.CreateParameter("@P2_DepNumber", adSmallInt, adParamInput)
  adoCmd.Parameters.Append adoParam2
  Set adoParam3 = adoCmd.CreateParameter("@P3_ClaimNumber", adChar, adParamInput, 14)
  adoCmd.Parameters.Append adoParam3
  adoCmd("@P1_SSN") = m_iSSN
  adoCmd("@P2_DepNumber") = m_iDepNum
  adoCmd("@P3_ClaimNumber") = m_iClaimNumber
%>
	
  <html>
	<head>
<!--#include file="VBFuncs.inc" -->
	</head>
	<body LANGUAGE="VBScript" onLoad="UpdateScreen(3)">
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
				<font SIZE="+2" face="verdana, arial, helvetica"><strong>Prescription Detail Information</strong></font>
			</td>
			<td ALIGN="CENTER" WIDTH="20%">
				<img SRC="images/bluebar2.gif" onClick="history.back">&nbsp;&nbsp;<img src="images/bluebar3.gif" onClick="history.go(+1)" border="0">
			</td>
		</tr>
	</table>
<%
  adoRS.Open adoCmd
  'adoRS.Open sSQL, adoConn, adOpenForwardOnly, adLockOptimistic
	if adoRS.EOF then
		Response.Write ("<p><ul><font color='red' face='verdana, arial, helvetica' size='2'><b>Unable to find all necessary information; please contact your network administrator.</b></font></ul>")
	else
%>
    <p>&nbsp;  
		<table WIDTH="100%" BORDER="1" bordercolor="white" bgcolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			<tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
				<td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>Claim #</b></font></td>
        <td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>SSN</b></font></td>
        <td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>Patient</b></font></td>
        <td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>Dep Number</b></font></td>
        <td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>Birthdate</b></font></td>
        <td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>Age</b></font></td>
        <td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>Gender</b></font></td>
        <td ALIGN="CENTER" bgcolor="#cccccc"><font COLOR="BLUE"><b>Relationship</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>GRP</b></font></td>
        <td align="center" bgcolor="#cccccc"><font COLOR="BLUE"><b>Plan</b></font></td>
      </tr>
      <tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
        <td align="center" nowrap bgcolor="#F0F0F0"><b><%= adoRS("ClaimNumber") %></b></td>
        <td align="center" nowrap bgcolor="#F0F0F0"><b><%= m_iSSN %></b></td>
        <td align="center" nowrap bgcolor="#F0F0F0"><b><%= adoRS("Patient") %></b></td>
        <td align="center" nowrap bgcolor="#F0F0F0"><b><%= m_iDepNum %></b></td>
        <td align="center" nowrap bgcolor="#F0F0F0"><b><%= adoRS("BirthDate") %></b></td>
        <td align="center" nowrap bgcolor="#F0F0F0"><b><%= DateDiff("d", adoRS("BirthDate"), adoRS("DateFilled")) \ 365%></b></td>
        <td align="center" nowrap bgcolor="#F0F0F0"><b><%= adoRS("Gender") %></b></td>
        <td align="center" nowrap bgcolor="#F0F0F0"><b>&nbsp;<%= adoRS("RelationCode") %></b></td>
        <td align="center" nowrap bgcolor="#F0F0F0"><b><%= adoRS("Fund") %></b></td>
        <td align="center" nowrap bgcolor="#F0F0F0"><b><%= adoRS("PlanCode") %></b></td>
      </tr>
    </table>
      
    <p>&nbsp;
    <center>
    <table WIDTH="50%" BORDER="1" bordercolor="white" bgcolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
      <tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
        <td align="right" nowrap bgcolor="#F0F0F0"><b>RX #:</b></td>
        <td align="left" STYLE='color:<%= Session("EmphColor")%>;' nowrap bgcolor="#F0F0F0"><b><%= adoRS("RxNumber") %></b></td>
        <td align="right" nowrap bgcolor="#F0F0F0"><b>Filled Date:</b></td>
        <td colspan="3" align="left" STYLE='color:<%= Session("EmphColor")%>;' nowrap bgcolor="#F0F0F0"><b><%= adoRS("DateFilled") %></b></td>
        <td></td>
        <td></td>
      </tr>
      <tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
        <td align="right" nowrap bgcolor="#F0F0F0"><b>Days Supply:</b></td>
        <td align="left" STYLE='color:<%= Session("EmphColor")%>;' nowrap bgcolor="#F0F0F0"><b><%= adoRS("DaysSupply") %></b></td>
        <td align="right" nowrap bgcolor="#F0F0F0"><b>Quantity:</b></td>
        <td colspan="3" align="left" STYLE='color:<%= Session("EmphColor")%>;' nowrap bgcolor="#F0F0F0"><b><%= adoRS("Quantity") %></b></td>
        <td></td>
        <td></td>
      </tr>
      <tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
        <td align="right" nowrap bgcolor="#F0F0F0"><b>DAW:</b></td>
        <td align="left" STYLE='color:<%= Session("EmphColor")%>;' nowrap bgcolor="#F0F0F0"><b><%= adoRS("DAW") %></b></td>
        <td align="right" nowrap bgcolor="#F0F0F0"><b>Generic:</b></td>
        <td align="left" STYLE='color:<%= Session("EmphColor")%>;' nowrap bgcolor="#F0F0F0"><b><%= adoRS("Generic") %></b></td>
        <td align="right" nowrap bgcolor="#F0F0F0"><b>Name:</b></td>
        <td align="left" STYLE='color:<%= Session("EmphColor")%>;' nowrap bgcolor="#F0F0F0"><b><% if adoRS("Description") <> "" then %>&nbsp;<%= adoRS("Description") %><% else %>&nbsp;NO NAME GIVEN<% end if %></b></td>
      </tr>
      <tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
        <td align="right" nowrap bgcolor="#F0F0F0"><b>Claim Type:</b></td>
        <td align="left" STYLE='color:<%= Session("EmphColor")%>;' nowrap bgcolor="#F0F0F0"><b><%= adoRS("ClaimType") %></b></td>
        <td align="right" nowrap bgcolor="#F0F0F0"><b>Pharmacy:</b></td>
        <td colspan="3" align="left" STYLE='color:<%= Session("EmphColor")%>;' nowrap bgcolor="#F0F0F0"><b><%= adoRS("NABP") %>&nbsp;&nbsp;<%= adoRS("Name") %></b></td>
        <td></td>
        <td></td>
      </tr>
      <tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
        <td align="right" nowrap bgcolor="#F0F0F0"><b>NDC:</b></td>
        <td align="left" STYLE='color:<%= Session("EmphColor")%>;' nowrap bgcolor="#F0F0F0"><b><%= adoRS("NDC") %></b></td>
        <td align="right" nowrap bgcolor="#F0F0F0"><b>Status/Status Date:</b></td>
        <td colspan="3" align="left" STYLE='color:<%= Session("EmphColor")%>;' nowrap bgcolor="#F0F0F0"><b><%= adoRS("Status") %><% if adoRS("DateStatus") <> "" then %>&nbsp;&nbsp;&nbsp;<%= formatdatetime(adoRS("DateStatus"),vbShortDate) %><% else %>&nbsp;</b><% end if %></td>
        <td></td>
        <td></td>
      </tr>
    </table>

    <table WIDTH="50%" BORDER="1" bordercolor="white" bgcolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
      <tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
        <td align="right" nowrap bgcolor="#F0F0F0"><b>PayCode:</b></td>
        <td colspan="3" align="left" STYLE='color:<%= Session("EmphColor")%>;' nowrap bgcolor="#F0F0F0"><b><%= adoRS("PayCode") %></b></td>
      </tr>
      <tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
        <td align="right" nowrap bgcolor="#F0F0F0"><b>Payee:</b></td>
        <td colspan="3" align="center" STYLE='color:<%= Session("EmphColor")%>;' nowrap bgcolor="#F0F0F0"><% if adoRS("PayeeName") <> "" then %><b><%= adoRS("PayeeName") %></b><% else %><b>No Check Issued!!</b><% end if %></td>
      </tr>
      <tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
        <td align="right" nowrap bgcolor="#F0F0F0"><b>Address:</b></td>
        <td colspan="3" align="center" STYLE='color:<%= Session("EmphColor")%>;' nowrap bgcolor="#F0F0F0"><b><%= adoRS("Street") %></b></td>
      </tr>
      <tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
        <td align="right" nowrap bgcolor="#F0F0F0"><b>City/State/Zip:</b></td>
        <td colspan="3" align="center" STYLE='color:<%= Session("EmphColor")%>;' nowrap bgcolor="#F0F0F0"><b><%= adoRS("City") %>&nbsp;&nbsp;<%= adoRS("State") %>&nbsp;&nbsp;<%= adoRS("PostalCode") %></b></td>
      </tr>
      <tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
        <td align="right" nowrap bgcolor="#F0F0F0"><b>Check Account:</b></td>
        <td align="center" STYLE='color:<%= Session("EmphColor")%>;' nowrap bgcolor="#F0F0F0"><b><%= adoRS("CheckAccount") %></b></td>
        <td align="right" nowrap bgcolor="#F0F0F0"><b>Check Number:</b></td>
        <td align="center" STYLE='color:<%= Session("EmphColor")%>;' nowrap bgcolor="#F0F0F0"><a href="CheckDetails.asp?CheckNum=<%= adoRS("CheckNumber") %>&CheckAcct=<%= adoRS("CheckAccount") %>" onClick="WorkingStatus()"><b><%= adoRS("CheckNumber") %></b></a></td>
      </tr>
    </table>
    </center>
    
    <p>&nbsp;
    <center>
    <table WIDTH="40%" BORDER="1" bordercolor="white" bgcolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			<tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
        <td width="50%" align="right" nowrap bgcolor="#F0F0F0"><b>Charge:</b></td>
        <td width="50%" align="left" STYLE='color:<%= Session("EmphColor")%>;' nowrap bgcolor="#F0F0F0"><% if adoRS("ChargeAmount") <> "" then %><b><%= formatcurrency(adoRS("ChargeAmount")) %><% else %>&nbsp;<%= adoRS("ChargeAmount") %></b><% end if %></td>
      </tr>
      <tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
        <td align="right" nowrap bgcolor="#F0F0F0"><b>Non-Pay:</b></td>
        <td align="left" STYLE='color:<%= Session("EmphColor")%>;' nowrap bgcolor="#F0F0F0"><% if adoRS("NonPayAmount") <> "" then %><b><%= formatcurrency(adoRS("NonPayAmount")) %><% else %>&nbsp;<%= adoRS("NonPayAmount") %></b><% end if %></td>
      </tr>
      <tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
        <td align="right" nowrap bgcolor="#F0F0F0"><b>Deductible:</b></td>
        <td align="left" STYLE='color:<%= Session("EmphColor")%>;' nowrap bgcolor="#F0F0F0"><% if adoRS("Deductible") <> "" then %><b><%= formatcurrency(adoRS("Deductible")) %><% else %>&nbsp;<%= adoRS("Deductible") %></b><% end if %></td>
      </tr>
      <tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
        <td align="right" nowrap bgcolor="#F0F0F0"><b>Co-Pay:</b></td>
        <td align="left" STYLE='color:<%= Session("EmphColor")%>;' nowrap bgcolor="#F0F0F0"><% if adoRS("Copay") <> "" then %><b><%= formatcurrency(adoRS("Copay")) %><% else %>&nbsp;<%= adoRS("Copay") %></b><% end if %></td>
      </tr>
      <tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
        <td align="right" nowrap bgcolor="#F0F0F0"><b>Paid:</b></td>
        <td align="left" STYLE='color:<%= Session("EmphColor")%>;' nowrap bgcolor="#F0F0F0"><% if adoRS("PaidAmount") <> "" then %><b><%= formatcurrency(adoRS("PaidAmount")) %><% else %>&nbsp;<%= adoRS("PaidAmount") %></b><% end if %></td>
      </tr>
    </table>
    </center>
        
<%
      adoRS.Close
			adoConn.Close
			set adoRS = nothing
			set adoConn = nothing
  end if
%>
  

  
</body>
</html>
