<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Untitled</title>
</head>

<body>

  set adoConn = Server.CreateObject("ADODB.Connection")
  adoConn.ConnectionTimeout = 0
	adoConn.CommandTimeout = 0
  set adoCmd = Server.CreateObject("ADODB.Command")
  adoCmd.CommandTimeout = 0
	set adoRS = Server.CreateObject("ADODB.Recordset")
	adoConn.Open Application("DataConn")
  
  sSQL = "select Fund FROM RX where SSN =" & m_iSSN & " AND DepNumber =" & m_iDepNum
  
  adoRS.Open sSQL, adoConn, adOpenForwardOnly, adLockOptimistic
  if not adoRS.EOF then
    'm_sHWType = adoRS("HWType")
    m_sFund = adoRS("Fund")
  end if
  
  adoRS.Close
  adoConn.Close
	set adoRS = nothing
	set adoConn = nothing
  
  '******************
' Medical Tracking
'******************
%>
		<br>
		<font SIZE="3" face="arial, helvetica"><center><b>Medical Tracking</b></center></font>
		<table border="1" cellpadding="2" cellspacing="2" width="100%" bgcolor="white" bordercolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			<tr ALIGN="CENTER" BorderColorDark="#F0F0F0" BorderColorlight="#999999">
			    <td bgcolor="#CCCCCC"><strong>Year</strong></td>
			    <td bgcolor="#CCCCCC"><strong>Dep #</strong></td>
          <td bgcolor="#CCCCCC"><strong>Fund</strong></td>
			    <td bgcolor="#CCCCCC"><strong>PlanCat</strong></td>
          <td bgcolor="#CCCCCC"><strong>HWType</strong></td>
          <td bgcolor="#CCCCCC"><strong>TrackCode</strong></td>
          <td bgcolor="#CCCCCC"><strong>Amount</strong></td>
          <td bgcolor="#CCCCCC"><strong>Services</strong></td>
			</tr>
<%  if not m_bDependent then
			sSQL = "select Year, DepNumber, Fund, PlanCat, HWType, TrackCode, Amount, Services"
			sSQL = sSQL & " FROM Tracking"
			sSQL = sSQL & " WHERE SSN = " & m_iSSN & " AND DepNumber in (99,0) AND HWType='M'"
			sSQL = sSQL & " order by Year DESC, DepNumber, HWType, PlanCat, TrackCode"
    else
      sSQL = "select Year, DepNumber, Fund, PlanCat, HWType, TrackCode, Amount, Services"
			sSQL = sSQL & " FROM Tracking"
			sSQL = sSQL & " WHERE SSN = " & m_iSSN & " AND (DepNumber = 99 or DepNumber=" & m_iDepNo & ") AND HWType='M'"
			sSQL = sSQL & " order by Year DESC, DepNumber, PlanCat, TrackCode"
    end if	
      'response.write sSQL				

			adoRS.Open sSQL, adoConn, adOpenForwardOnly, adLockOptimistic
			if not adoRS.EOF then
				Do While Not adoRS.EOF
%>
					<tr align="center" BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
					    <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0">&nbsp;<b><%= adoRS("Year")%></b></td>
					    <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0">&nbsp;<b><%= adoRS("DepNumber")%></b></td>
              <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0">&nbsp;<b><%= adoRS("Fund")%></b></td>
					    <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0">&nbsp;<b><%= adoRS("PlanCat")%></b></td>
              <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0">&nbsp;<b><%= adoRS("HWType")%></b></td>
              <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0">&nbsp;<b><%= adoRS("TrackCode")%></b></td>
              <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0">&nbsp;<b><%= FORMATCURRENCY(adoRS("Amount"))%></b></td>
              <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0">&nbsp;<b><%= adoRS("Services")%></b></td>
					</tr>
<% 
					adoRS.MoveNext
				Loop 
			else
%>
				<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC" align="center">
				    <td COLSPAN="8" bgcolor="#F0F0F0"><font color="red" face="verdana, arial, helvetica" size="2"><b>No Medical Tracking Data Found.</b></font></td>
				</tr>
<%
			end if  
%>
		    <tr>
          <td colspan="8" align="right"><img SRC="images/bluebar2.gif" onClick="history.go(-1)">&nbsp;&nbsp;<img src="images/bluebar3.gif" onClick="history.go(+1)" border="0"></td>
        </tr>
    
    </table>
<%
		adoRS.Close

'*****************
' Dental Tracking
'*****************
%>
		<br>
		<font SIZE="3" face="arial, helvetica"><center><b>Dental Tracking</b></center></font>
		<table border="1" cellpadding="2" cellspacing="2" width="100%" bgcolor="white" bordercolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			<tr ALIGN="CENTER" BorderColorDark="#F0F0F0" BorderColorlight="#999999">
			    <td bgcolor="#CCCCCC"><strong>Year</strong></td>
			    <td bgcolor="#CCCCCC"><strong>Dep #</strong></td>
          <td bgcolor="#CCCCCC"><strong>Fund</strong></td>
			    <td bgcolor="#CCCCCC"><strong>PlanCat</strong></td>
          <td bgcolor="#CCCCCC"><strong>HWType</strong></td>
          <td bgcolor="#CCCCCC"><strong>TrackCode</strong></td>
          <td bgcolor="#CCCCCC"><strong>Amount</strong></td>
          <td bgcolor="#CCCCCC"><strong>Services</strong></td>
			</tr>
<%  if not m_bDependent then
			sSQL = "select Year, DepNumber, Fund, PlanCat, HWType, TrackCode, Amount, Services"
			sSQL = sSQL & " FROM Tracking"
			sSQL = sSQL & " WHERE SSN = " & m_iSSN & " AND DepNumber in (99,0) AND HWType='D'"
			sSQL = sSQL & " order by Year DESC, DepNumber, HWType, PlanCat, TrackCode"
    else
      sSQL = "select Year, DepNumber, Fund, PlanCat, HWType, TrackCode, Amount, Services"
			sSQL = sSQL & " FROM Tracking"
			sSQL = sSQL & " WHERE SSN = " & m_iSSN & " AND (DepNumber = 99 or DepNumber=" & m_iDepNo & ") AND HWType='D'"
			sSQL = sSQL & " order by Year DESC, DepNumber, PlanCat, TrackCode"
    end if	
      'response.write sSQL				

			adoRS.Open sSQL, adoConn, adOpenForwardOnly, adLockOptimistic
			if not adoRS.EOF then
				Do While Not adoRS.EOF
%>
					<tr align="center" BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC">
					    <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0">&nbsp;<b><%= adoRS("Year")%></b></td>
					    <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0">&nbsp;<b><%= adoRS("DepNumber")%></b></td>
              <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0">&nbsp;<b><%= adoRS("Fund")%></b></td>
					    <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0">&nbsp;<b><%= adoRS("PlanCat")%></b></td>
              <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0">&nbsp;<b><%= adoRS("HWType")%></b></td>
              <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0">&nbsp;<b><%= adoRS("TrackCode")%></b></td>
              <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0">&nbsp;<b><%= FORMATCURRENCY(adoRS("Amount"))%></b></td>
              <td STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0">&nbsp;<b><%= adoRS("Services")%></b></td>
					</tr>
<% 
					adoRS.MoveNext
				Loop 
			else
%>
				<tr BorderColorDark="#FCFCFC" BorderColorlight="#CCCCCC" align="center">
				    <td COLSPAN="8" bgcolor="#F0F0F0"><font color="red" face="verdana, arial, helvetica" size="2"><b>No Dental Tracking Data Found.</b></font></td>
				</tr>
<%
			end if  
%>
		    <tr>
          <td colspan="8" align="right"><img SRC="images/bluebar2.gif" onClick="history.go(-1)">&nbsp;&nbsp;<img src="images/bluebar3.gif" onClick="history.go(+1)" border="0"></td>
        </tr>
    
    </table>
<%
		adoRS.Close
    
    if msgbox("If both Check Account and Check Number are not specified, the search can be slow.  Search anyway?",vbOkCancel,"Caution") <> vbOk then
				exit sub
			end if


</body>
</html>
