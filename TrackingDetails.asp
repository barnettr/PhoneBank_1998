<%@ Language=VBScript %>
<% Option Explicit %>
<% Response.Expires = 0 %> 
<!--#include file="user.inc" -->
<!--#include file="adovbs_mod.inc" -->
<%
dim m_iSSN, adoConn, adoRS, sSQL, m_iDepNum, adoCmd, adoParam1, adoParam2, adoParam3
dim m_sTrackCode

  m_iSSN = request.querystring("SSN")
  m_iDepNum = request.querystring("DepNumber")
  m_sTrackCode = request.querystring("TrackCode")
  
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
  
  
  set adoConn = Server.CreateObject("ADODB.Connection")
  adoConn.ConnectionTimeout = 0
	adoConn.CommandTimeout = 0
  set adoCmd = Server.CreateObject("ADODB.Command")
  adoCmd.CommandTimeout = 0
	set adoRS = Server.CreateObject("ADODB.Recordset")
	adoConn.Open Application("DataConn")
  
  adoCmd.CommandText = "LookUpChargeLineTracking"
	adoCmd.CommandType = adCmdStoredProc
	Set adoCmd.ActiveConnection = adoConn
  Set adoParam1 = adoCmd.CreateParameter("@P1_SSN", adInteger, adParamInput)
	adoCmd.Parameters.Append adoParam1
  Set adoParam2 = adoCmd.CreateParameter("@P2_DepNumber", adInteger, adParamInput)
  adoCmd.Parameters.Append adoParam2
  Set adoParam3 = adoCmd.CreateParameter("@P3_TrackCode", adChar, adParamInput, 3)
  adoCmd.Parameters.Append adoParam3
  adoCmd("@P1_SSN") = m_iSSN
  adoCmd("@P2_DepNumber") = m_iDepNum
  adoCmd("@P3_TrackCode") = m_sTrackCode

%>

  <html>
  <head>
  <!--#include file="VBFuncs.inc" -->
  </head>
  
  <body onload="UpdateScreen(3)">
    <table width="100%" border="0">
      <tr>
        <td align="right"><img SRC="images/bluebar2.gif" onClick="history.go(-1)">&nbsp;&nbsp;<img src="images/bluebar3.gif" onClick="history.go(+1)" border="0"></td>
      </tr>
    </table>  

		<br>	
		<font SIZE="3" face="arial, helvetica"><center><b>Charge Line Tracking</b></center></font>
    <br>
		<table WIDTH="100%" BORDER="1" CELLPADDING="2" CELLSPACING="2" bgcolor="white" bordercolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			<tr ALIGN="CENTER" BorderColorDark="#F0F0F0" BorderColorlight="#999999">
	      <td bgcolor="#cccccc"><font COLOR="BLUE"><strong>Fund</strong></font></td>
	      <td bgcolor="#cccccc"><font COLOR="BLUE"><strong>Status</strong></font></td>
	      <td bgcolor="#cccccc"><font COLOR="BLUE"><strong>Diagnosis</strong></font></td>
	      <td bgcolor="#cccccc"><font COLOR="BLUE"><strong>PPO</strong></font></td>
	      <td bgcolor="#cccccc"><font COLOR="BLUE"><strong>BP</strong></font></td>
	      <td bgcolor="#cccccc"><font COLOR="BLUE"><strong>Tracked</strong></font></td>
	      <td bgcolor="#cccccc"><font COLOR="BLUE"><strong>Units</strong></font></td>
	      <td bgcolor="#cccccc"><font COLOR="BLUE"><strong>From</strong></font></td>
	      <td bgcolor="#cccccc"><font COLOR="BLUE"><strong>Thru</strong></font></td>
	      <td bgcolor="#cccccc"><font COLOR="BLUE"><strong>Procedure</strong></font></td>
	      <td bgcolor="#cccccc"><font COLOR="BLUE"><strong>Provider</strong></font></td>
	      <td bgcolor="#cccccc"><font COLOR="BLUE"><strong>Claim</strong></font></td>
			</tr>
<%
  adoRS.Open adoCmd
  if not adoRS.EOF then
		Do While Not adoRS.EOF
%>

      <tr ALIGN="CENTER" BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
        <td bgcolor="#F0F0F0">&nbsp;<%= adoRS("Fund") %></td>
        <td bgcolor="#F0F0F0">&nbsp;<%= adoRS("Status")%></td>
        <td bgcolor="#F0F0F0">&nbsp;<%= adoRS("DiagnosisCode") %></td>
        <td bgcolor="#F0F0F0">&nbsp;<%= adoRS("PPOSponsor")%></td>
        <td bgcolor="#F0F0F0">&nbsp;<%= adoRS("IncidentNumber")%></td>
        <td bgcolor="#F0F0F0">&nbsp;<% if adoRS("TrackedAmount") <> "" then %><%= formatcurrency(adoRS("TrackedAmount"))%><% else %>&nbsp;<% end if %></td>
        <td bgcolor="#F0F0F0">&nbsp;<%= adoRS("TrackedUnits")%></td>
        <td bgcolor="#F0F0F0">&nbsp;<% if adoRS("FromDate") <> "" then %><%= FORMATDATETIME(adoRS("FromDate"),vbShortDate) %><% else %>&nbsp;<%= adoRS("FromDate") %><% end if %></td>
        <td bgcolor="#F0F0F0">&nbsp;<% if adoRS("ThruDate") <> "" then %><%= FORMATDATETIME(adoRS("ThruDate"),vbShortDate)%><% else %>&nbsp;<%= adoRS("ThruDate") %><% end if %></td>
        <td bgcolor="#F0F0F0">&nbsp;<%= adoRS("ProcedureCode") %></td>
        <td bgcolor="#F0F0F0">&nbsp;<%= adoRS("FullName") %></td>
        <td bgcolor="#F0F0F0">&nbsp;<a HREF="aspexec.asp?LogDate=<%= adoRS("LogDate")%>&amp;Sequence=<%= adoRS("Sequence")%>&amp;Distr=<%=  adoRS("Distr")%>&ClaimNum=<%= ConvertToClaimNum(adoRS("LogDate"),adoRS("Sequence"),adoRS("Distr")) %>" onClick="WorkingStatus()"><%= ConvertToClaimNum(adoRS("LogDate"),adoRS("Sequence"),adoRS("Distr")) %></a></td>
      </tr>
      
<%  
      adoRS.MoveNext
    Loop
  else
%>    
      <tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC" align="center">
				 <td COLSPAN="12" bgcolor="#F0F0F0"><font color="red" face="verdana, arial, helvetica" size="2"><b>No ChargeLine Tracking Information Available.</b></font></td>
			</tr>
    
    </table>
 <% end if 
    adoRS.Close
		adoConn.Close
		set adoRS = nothing
		set adoConn = nothing
%>
        
  </body>
  </html>