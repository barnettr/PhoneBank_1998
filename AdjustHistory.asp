<%@ Language=VBScript %>
<% Option Explicit %>
<% Response.Expires = 0 %>
<!--#include file="user.inc" -->
<!--#include file="adovbs_mod.inc" -->
<%
Dim m_iLogDate, m_iSequence, m_iDistr, m_iSSN
Dim i, adoConn, adoRS, sSQL, sTemp, adoCmd, adoParam

m_iLogDate = request.querystring("LogDate")
m_iSequence = request.querystring("Sequence")
m_iDistr = request.querystring("Distr")
m_iSSN = request.querystring("SSN")

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

if m_iLogDate = "" or m_iSequence = "" or m_iDistr = "" then
  response.write "<p><font color='red' size='2' face='verdana, arial, helvetica'><b>Claim Number is missing; contact your network administrator</b></font>"
else
  
  Set adoConn = Server.CreateObject("ADODB.Connection")
  adoConn.CommandTimeout = 0
  adoConn.Open Application("DataConn")

  Set adoCmd = Server.CreateObject("ADODB.Command")
  adoCmd.ActiveConnection = adoConn
  adoCmd.CommandType = &H0004 'adParamReturnValue
  adoCmd.CommandText = "FindAllAdjByClaim"
  
    Set adoParam			  = adoCmd.CreateParameter("@P1_LogDate")
		adoParam.Type			  = 135
		adoParam.Direction	= &H0001
		adoParam.Size			  = 8
		adoParam.Value			= m_iLogDate
		adoCmd.Parameters.Append adoParam
    
    Set adoParam			  = adoCmd.CreateParameter("@P1_Sequence")
		adoParam.Type			  = 3
		adoParam.Direction	= &H0001
		adoParam.Size			  = 4
		adoParam.Value			= m_iSequence
		adoCmd.Parameters.Append adoParam
    
    Set adoParam			  = adoCmd.CreateParameter("@P1_Distr")
		adoParam.Type			  = 3
		adoParam.Direction	= &H0001
		adoParam.Size			  = 4
		adoParam.Value			= m_iDistr
		adoCmd.Parameters.Append adoParam
    
  Set adoRS = Server.CreateObject("ADODB.Recordset")
  

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
		<font SIZE="3" face="arial, helvetica"><center><b>Adjust History Information</b></center></font>
    <br>
		<table WIDTH="100%" BORDER="1" CELLPADDING="2" CELLSPACING="2" bgcolor="white" bordercolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			<tr ALIGN="CENTER" BorderColorDark="#F0F0F0" BorderColorlight="#999999">
	      <td bgcolor="#cccccc"><font COLOR="BLUE"><strong>StatusDate</strong></font></td>
	      <td bgcolor="#cccccc"><font COLOR="BLUE"><strong>Status</strong></font></td>
	      <td bgcolor="#cccccc"><font COLOR="BLUE"><strong>ClaimNumber</strong></font></td>
	      <td bgcolor="#cccccc"><font COLOR="BLUE"><strong>Line</strong></font></td>
	      <td bgcolor="#cccccc"><font COLOR="BLUE"><strong>Dep#</strong></font></td>
	      <td bgcolor="#cccccc"><font COLOR="BLUE"><strong>Year</strong></font></td>
	      <td bgcolor="#cccccc"><font COLOR="BLUE"><strong>FromDate</strong></font></td>
	      <td bgcolor="#cccccc"><font COLOR="BLUE"><strong>Procedure</strong></font></td>
	      <td bgcolor="#cccccc"><font COLOR="BLUE"><strong>Charge</strong></font></td>
	      <td bgcolor="#cccccc"><font COLOR="BLUE"><strong>Benefit</strong></font></td>
	      <td bgcolor="#cccccc"><font COLOR="BLUE"><strong>Paid</strong></font></td>
	      <td bgcolor="#cccccc"><font COLOR="BLUE"><strong>Code</strong></font></td>
        <td bgcolor="#cccccc"><font COLOR="BLUE"><strong>Units</strong></font></td>
        <td bgcolor="#cccccc"><font COLOR="BLUE"><strong>Tracked</strong></font></td>
        <td bgcolor="#cccccc"><font COLOR="BLUE"><strong>Reason</strong></font></td>
        <td bgcolor="#cccccc"><font COLOR="BLUE"><strong>Provider</strong></font></td>
			</tr>
<%
  adoRS.Open adoCmd
	if not adoRS.EOF then
		Do While Not adoRS.EOF
%>      
      
      <tr ALIGN="CENTER" BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
        <td bgcolor="#F0F0F0">&nbsp;<% if adoRS("StatusDate") <> "" then %><%= FORMATDATETIME(adoRS("StatusDate"),vbShortDate)  %><% else %>&nbsp;<% end if %></td>
        <td bgcolor="#F0F0F0">&nbsp;<%= adoRS("Status")%></td>
        <td bgcolor="#F0F0F0">&nbsp;<%= ConvertToClaimNum(adoRS("LogDate"),adoRS("Sequence"),adoRS("Distr")) %></td>
        <td bgcolor="#F0F0F0">&nbsp;<%= adoRS("LineNumber")%></td>
        <td bgcolor="#F0F0F0">&nbsp;<%= adoRS("DepNumber")%></td>
        <td bgcolor="#F0F0F0">&nbsp;<%= adoRS("Year")%></td>
        <td bgcolor="#F0F0F0">&nbsp;<%= adoRS("FromDate")%></td>
        <td bgcolor="#F0F0F0">&nbsp;<%= adoRS("ProcedureCode")%></td>
        <td bgcolor="#F0F0F0">&nbsp;<% if adoRS("ChargeAmount") <> "" then %><%= formatcurrency(adoRS("ChargeAmount"))%><% else %>&nbsp;<%= adoRS("ChargeAmount") %><% end if %></td>
        <td bgcolor="#F0F0F0">&nbsp;<% if adoRS("AllowableBenefit") <> "" then %><%= formatcurrency(adoRS("AllowableBenefit"))%><% else %>&nbsp;<%= adoRS("AllowableBenefit") %><% end if %></td>
        <td bgcolor="#F0F0F0">&nbsp;<% if adoRS("PaidAmount") <> "" then %><%= formatcurrency(adoRS("PaidAmount"))%><% else %>&nbsp;<%= adoRS("PaidAmount") %><% end if %></td>
        <td bgcolor="#F0F0F0">&nbsp;<%= adoRS("TrackCode")%></td>
        <td bgcolor="#F0F0F0">&nbsp;<%= adoRS("TrackedUnits")%></td>
        <td bgcolor="#F0F0F0">&nbsp;<% if adoRS("TrackedAmount") <> "" then %><%= formatcurrency(adoRS("TrackedAmount"))%><% else %>&nbsp;<%= adoRS("TrackedAmount") %><% end if %></td>
        <td bgcolor="#F0F0F0">&nbsp;<%= adoRS("Reason")%></td>
        <td bgcolor="#F0F0F0">&nbsp;<%= adoRS("ProviderName")%></td>
      </tr>
<%  
      adoRS.MoveNext
    Loop
  else
%>    
      <tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC" align="center">
				 <td COLSPAN="16" bgcolor="#F0F0F0"><font color="red" face="verdana, arial, helvetica" size="2"><b>No Adjust History Information Available.</b></font></td>
			</tr>
    
    </table>
<%  end if 
    adoRS.Close
		adoConn.Close
		set adoRS = nothing
		set adoConn = nothing
%>
        
  </body>
  </html>
  
<% 
  end if 
%>