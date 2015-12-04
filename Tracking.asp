<%@ Language=VBScript %>
<% Option Explicit %>
<% Response.Expires = 0 %>
<!--#include file="user.inc" -->
<!--#include file="adovbs_mod.inc" -->
<%
dim m_iSSN, adoConn, adoRS, sSQL, m_iDepNum, adoCmd, adoParam1, adoParam2
dim iPageNo, m_iLife, m_i1999, m_i1998, m_i1997, m_i1996, m_i1995, m_i1994
dim adoParam3, adoParam4, adoParam5, adoParam6, adoParam7, adoParam8
dim m_sMedical, m_sDental, m_sVision, m_sCSB, m_sPrescription, OOP
dim m_iYears, m_sHWType, m_sFund, m_bDependent
dim m_iOOPIND, m_iOOPFAM, m_iDEDIND, m_iDEDFAM, m_iPPOIND, m_iPPOFAM, m_iDE4IND, m_iDE4FAM, m_iGrpLife, m_iGrpLife2
dim m_iOVP, m_iCOB, m_sName, m_iPREYTD, bUseSQL
dim sTemp

  m_iSSN = request.querystring("SSN")
  m_iDepNum = request.querystring("DepNum")
  m_iYears = request.querystring("Years")
  m_sHWType = request.querystring("HWType")
  m_sFund = request.querystring("Fund")
  
  	set adoConn = Server.CreateObject("ADODB.Connection")
  	adoConn.ConnectionTimeout = 0
	adoConn.CommandTimeout = 0
  	set adoCmd = Server.CreateObject("ADODB.Command")
  	adoCmd.CommandTimeout = 0
	set adoRS = Server.CreateObject("ADODB.Recordset")
	adoConn.Open Application("DataConn")
  
  	adoCmd.CommandText = "pb_GetTrackingName"
	adoCmd.CommandType = &H0004
	Set adoCmd.ActiveConnection = adoConn
  	Set adoParam1 = adoCmd.CreateParameter("@P1_SSN", adInteger, &H0001)
	adoCmd.Parameters.Append adoParam1
  	Set adoParam2 = adoCmd.CreateParameter("@P2_DepNumber", adInteger, &H0001)
  	adoCmd.Parameters.Append adoParam2
  	adoCmd("@P1_SSN") = m_iSSN
  	adoCmd("@P2_DepNumber") = m_iDepNum
    
  adoRS.Open adoCmd
  if not adoRS.EOF then
    m_sName = adoRS("Name")
  end if
  
  adoRS.Close
  
  set adoCmd = Server.CreateObject("ADODB.Command")
	set adoRS = Server.CreateObject("ADODB.Recordset")
  
  if m_iDepNum = 0 then
		m_bDependent = false
  
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
  if Request.Querystring("Years") <> "" then
		m_iYears = Request.Querystring("Years")
  else
    m_iYears = Year(Date)
	end if
	if Request.Querystring("HWType") <> "" then
		m_sHWType = Request.Querystring("HWType")
  end if
  if Request.Querystring("Fund") <> "" then
		m_sFund = Request.Querystring("Fund")
  end if
  
  adoCmd.CommandText = "pb_GetTracking"
	adoCmd.CommandType = &H0004
	Set adoCmd.ActiveConnection = adoConn
  Set adoParam1 = adoCmd.CreateParameter("@P1_SSN", adInteger, &H0001)
	adoCmd.Parameters.Append adoParam1
  Set adoParam2 = adoCmd.CreateParameter("@P2_Year", adInteger, &H0001)
  adoCmd.Parameters.Append adoParam2
  Set adoParam3 = adoCmd.CreateParameter("@P3_HWType", adChar, &H0001, 1)
  adoCmd.Parameters.Append adoParam3
  Set adoParam4 = adoCmd.CreateParameter("@P4_Fund", adChar, &H0001, 3)
  adoCmd.Parameters.Append adoParam4
  adoCmd("@P1_SSN") = m_iSSN
  adoCmd("@P2_Year") = m_iYears
  adoCmd("@P3_HWType") = m_sHWType
  adoCmd("@P4_Fund") = m_sFund
  
  else m_bdependent = true
  
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
  end if
  if Request.Querystring("Years") <> "" then
		m_iYears = Request.Querystring("Years")
  else
    m_iYears = Year(Date)
	end if
	if Request.Querystring("HWType") <> "" then
	  m_sHWType = Request.Querystring("HWType")
  end if
  if Request.Querystring("Fund") <> "" then
		m_sFund = Request.Querystring("Fund")
  end if
  
  adoCmd.CommandText = "pb_GetTrackingDep"
	adoCmd.CommandType = &H0004
	Set adoCmd.ActiveConnection = adoConn
  Set adoParam1 = adoCmd.CreateParameter("@P1_SSN", adInteger, &H0001)
	adoCmd.Parameters.Append adoParam1
  Set adoParam2 = adoCmd.CreateParameter("@P2_Year", adInteger, &H0001)
  adoCmd.Parameters.Append adoParam2
  Set adoParam3 = adoCmd.CreateParameter("@P3_HWType", adChar, &H0001, 1)
  adoCmd.Parameters.Append adoParam3
  Set adoParam4 = adoCmd.CreateParameter("@P4_Fund", adChar, &H0001, 3)
  adoCmd.Parameters.Append adoParam4
  Set adoParam5 = adoCmd.CreateParameter("@P5_DepNumber", adInteger, &H0001)
  adoCmd.Parameters.Append adoParam5
  adoCmd("@P1_SSN") = m_iSSN
  adoCmd("@P2_Year") = m_iYears
  adoCmd("@P3_HWType") = m_sHWType
  adoCmd("@P4_Fund") = m_sFund
  adoCmd("@P5_DepNumber") = m_iDepNum
  
  end if
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
				<font SIZE="+2" face="verdana, arial, helvetica"><strong>Tracking Information</strong></font>
			</td>
			<td ALIGN="CENTER" WIDTH="20%">
				<img SRC="images/bluebar2.gif" onClick="history.back">&nbsp;&nbsp;<img src="images/bluebar3.gif" onClick="history.go(+1)" border="0">
			</td>
		</tr>
	</table>


<table COLS="2" WIDTH="100%" CELLPADDING="0" CELLSPACING="0" BORDER="0">
	<tr>
		<td WIDTH="10%"><input TYPE="IMAGE" SRC="images/query_button.gif" ACCESSKEY="S" NAME="NewQuery" VALUE="Start Query" onClick="ValidSend(0)"></td>
		<td WIDTH="65%"><font face="verdana, arial, helvetica" size="2"><b>Enter your criteria, and then click the Start Query button or key (ALT+S).</b></font></td>
  </tr>
</table>

<div STYLE="height=13%; width:100%; overflow=auto;">
			<form ID="Criteria" METHOD="GET" ACTION="Tracking.asp" TARGET="Details">
				<table CLASS="CriteriaTable" COLS="4" CELLPADDING="0" CELLSPACING="0" border="0"> 
					<tr>
<!-- These TH's were required in order to get IE to deal properly with the DIV...among other things, without them and with BORDER=1, things were displayed OK, butwith the desired BORDER=0 IE seemed to 'lose' a number of the elements-->
						<th>
						<th>
						<th>
						<th>
            <th>
            <th>
            <th>
            <th>
            <th>
            <th>
          </tr>
          <tr>
            <td class="White" align="right">SSN:</td>
						<td><td><input TYPE="TEXT" ID="SSN" NAME="SSN" SIZE="10" MAXLENGTH=10 VALUE="<%= m_iSSN%>"></td></td>
						<td class="White" align="right">Dependent Number:</td>
						<td><input TYPE="TEXT" ID="DepNum" NAME="DepNum" SIZE="2" MAXLENGTH=2 VALUE="<%= m_iDepNum%>"></td>
            <td class="White" align="right">Year:</td>
            <td><!--- <input TYPE="TEXT" ID="Years" NAME="Years" SIZE="4" MAXLENGTH=4 VALUE="<%= m_iYears%>"> --->
              <select name="Years" ID="Years">
                <option value="">--Choose One--</option>
                <option value="9999" <% if request.querystring("m_iYears") = "9999" then %>Selected<% end if %>>9999</option>
                <option value="<%= Year(Date) %>" <% if Request.QueryString("m_iYears") = "Year(Date)" then %>Selected<% end if %>><%= Year(Date) %></option>
                <option value="<%= (Year(Date)-1) %>" <% if Request.QueryString("m_iYears") = "(Year(Date)-1)" then %>Selected<% end if %>><%= (Year(Date)-1) %></option>
                <option value="<%= (Year(Date)-2) %>" <% if Request.QueryString("m_iYears") = "(Year(Date)-2)" then %>Selected<% end if %>><%= (Year(Date)-2) %></option>
                <option value="<%= (Year(Date)-3) %>" <% if Request.QueryString("m_iYears") = "(Year(Date)-3)" then %>Selected<% end if %>><%= (Year(Date)-3) %></option>
                <option value="<%= (Year(Date)-4) %>" <% if Request.QueryString("m_iYears") = "(Year(Date)-4)" then %>Selected<% end if %>><%= (Year(Date)-4) %></option>
                <option value="<%= (Year(Date)-5) %>" <% if Request.QueryString("m_iYears") = "(Year(Date)-5)" then %>Selected<% end if %>><%= (Year(Date)-5) %></option>  
              </select>
                </td>
            <td class="White" align="right">HWType:</td>
            <td><input TYPE="TEXT" ID="HWType" NAME="HWType" SIZE="1" MAXLENGTH=1 VALUE="<%= m_sHWType%>"></td>
            <td class="White" align="right">Fund:</td>
            <td><input TYPE="TEXT" ID="Fund" NAME="Fund" SIZE="3" MAXLENGTH=3 VALUE="<%= m_sFund%>"></td>
					</tr>
					<tr>
            <td colspan="10">&nbsp;</td>
          </tr>
				</table>
				<input TYPE="HIDDEN" ID="ActionType" NAME="ActionType" VALUE>
			</form>
		</div>
		<br CLEAR="LEFT">
    
<%
  sTemp = Year(Date)
  adoRS.Open adoCmd
	if adoRS.EOF then
		Response.Write ("<p><center><font color='red' face='verdana, arial, helvetica' size='2'><b>The year defaults to " & sTemp &".</b></font></center>")
	else
  Do While Not adoRS.EOF
    if adoRS("TrackCode") = "OOP" and adoRS("DepNumber") <> 99 then
      m_iOOPIND = formatcurrency(adoRS("Amount"))
    end if
    if adoRS("TrackCode") = "OOP" and adoRS("DepNumber") = 99 then
      m_iOOPFAM = formatcurrency(adoRS("Amount"))
    end if
    if adoRS("TrackCode") = "DED" and adoRS("DepNumber") <> 99 then
      m_iDEDIND = formatcurrency(adoRS("Amount"))
    end if
    if adoRS("TrackCode") = "DED" and adoRS("DepNumber") = 99 then
      m_iDEDFAM = formatcurrency(adoRS("Amount"))
    end if
    if adoRS("TrackCode") = "PPO" and adoRS("DepNumber") <> 99 then
      m_iPPOIND = formatcurrency(adoRS("Amount"))
    end if
    if adoRS("TrackCode") = "PPO" and adoRS("DepNumber") = 99 then
      m_iPPOFAM = formatcurrency(adoRS("Amount"))
    end if
    if adoRS("TrackCode") = "RSV" and adoRS("DepNumber") <> 99 then
      m_iCOB = formatcurrency(adoRS("Amount"))
    end if
    
		  adoRS.MoveNext
			Loop
      adoRS.Close
  end if
      
      set adoCmd = Server.CreateObject("ADODB.Command")
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
      end if
      if Request.Querystring("HWType") <> "" then
	      m_sHWType = Request.Querystring("HWType")
      end if
      if Request.Querystring("Fund") <> "" then
		    m_sFund = Request.Querystring("Fund")
      end if
      
      adoCmd.CommandText = "pb_GetGroupLife"
    	adoCmd.CommandType = &H0004
    	Set adoCmd.ActiveConnection = adoConn
      Set adoParam1 = adoCmd.CreateParameter("@P1_SSN", adInteger, &H0001)
    	adoCmd.Parameters.Append adoParam1
      Set adoParam3 = adoCmd.CreateParameter("@P3_HWType", adChar, &H0001, 1)
      adoCmd.Parameters.Append adoParam3
      Set adoParam4 = adoCmd.CreateParameter("@P4_Fund", adChar, &H0001, 3)
      adoCmd.Parameters.Append adoParam4
      Set adoParam5 = adoCmd.CreateParameter("@P5_DepNumber", adInteger, &H0001)
      adoCmd.Parameters.Append adoParam5
      adoCmd("@P1_SSN") = m_iSSN
      adoCmd("@P3_HWType") = m_sHWType
      adoCmd("@P4_Fund") = m_sFund
      adoCmd("@P5_DepNumber") = m_iDepNum
      
      adoRS.Open adoCmd
      
      Do While Not adoRS.EOF
        if adoRS("TrackCode") = "LIF" and adoRS("DepNumber") <> 99 then
          m_iGrpLife = formatcurrency(adoRS("Amount"))
        end if
    
		  adoRS.MoveNext
			Loop
      adoRS.Close
      
      set adoCmd = Server.CreateObject("ADODB.Command")
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
      end if
      if Request.Querystring("HWType") <> "" then
	      m_sHWType = Request.Querystring("HWType")
      end if
      if Request.Querystring("Fund") <> "" then
		    m_sFund = Request.Querystring("Fund")
      end if
      
      adoCmd.CommandText = "pb_GetOVP"
    	adoCmd.CommandType = &H0004
    	Set adoCmd.ActiveConnection = adoConn
      Set adoParam1 = adoCmd.CreateParameter("@P1_SSN", adInteger, &H0001)
    	adoCmd.Parameters.Append adoParam1
      Set adoParam3 = adoCmd.CreateParameter("@P3_HWType", adChar, &H0001, 1)
      adoCmd.Parameters.Append adoParam3
      Set adoParam4 = adoCmd.CreateParameter("@P4_Fund", adChar, &H0001, 3)
      adoCmd.Parameters.Append adoParam4
      adoCmd("@P1_SSN") = m_iSSN
      adoCmd("@P3_HWType") = m_sHWType
      adoCmd("@P4_Fund") = m_sFund
      
      adoRS.Open adoCmd
      
      Do While Not adoRS.EOF
        if adoRS("TrackCode") = "OVP" and adoRS("DepNumber") = 99 then
          m_iOVP = formatcurrency(adoRS("Amount"))
        end if
    
		  adoRS.MoveNext
			Loop
      adoRS.Close
      
      set adoCmd = Server.CreateObject("ADODB.Command")
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
      end if
      if Request.Querystring("HWType") <> "" then
	      m_sHWType = Request.Querystring("HWType")
      end if
      if Request.Querystring("Fund") <> "" then
		    m_sFund = Request.Querystring("Fund")
      end if
      if Request.Querystring("Years") <> "" then
		    m_iYears = (Request.Querystring("Years") - 1)
      else
        m_iYears = (Year(Date)-1)
	    end if
      
      adoCmd.CommandText = "pb_GetQuarter"
    	adoCmd.CommandType = &H0004
    	Set adoCmd.ActiveConnection = adoConn
      Set adoParam1 = adoCmd.CreateParameter("@P1_SSN", adInteger, &H0001)
    	adoCmd.Parameters.Append adoParam1
      Set adoParam2 = adoCmd.CreateParameter("@P2_DepNumber", adInteger, &H0001)
      adoCmd.Parameters.Append adoParam2
      Set adoParam3 = adoCmd.CreateParameter("@P3_HWType", adChar, &H0001, 1)
      adoCmd.Parameters.Append adoParam3
      Set adoParam4 = adoCmd.CreateParameter("@P4_Fund", adChar, &H0001, 3)
      adoCmd.Parameters.Append adoParam4
      Set adoParam5 = adoCmd.CreateParameter("@P5_Year", adInteger, &H0001)
      adoCmd.Parameters.Append adoParam5
      adoCmd("@P1_SSN") = m_iSSN
      adoCmd("@P2_DepNumber") = m_iDepNum
      adoCmd("@P3_HWType") = m_sHWType
      adoCmd("@P4_Fund") = m_sFund
      adoCmd("@P5_Year") = m_iYears
      
      adoRS.Open adoCmd
      
      Do While Not adoRS.EOF
        if adoRS("DepNumber") <> 99 then
          m_iDE4IND = formatcurrency(adoRS("Amount"))
        end if
        if adoRS("DepNumber") = 99 then
          m_iDE4FAM = formatcurrency(adoRS("Amount"))
        end if
    
		  adoRS.MoveNext
			Loop
      adoRS.Close
      
      set adoCmd = Server.CreateObject("ADODB.Command")
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
      end if
      if Request.Querystring("HWType") <> "" then
	      m_sHWType = Request.Querystring("HWType")
      end if
      if Request.Querystring("Fund") <> "" then
		    m_sFund = Request.Querystring("Fund")
      end if
      if Request.Querystring("Years") <> "" then
		    m_iYears = Request.Querystring("Years")
      else
        m_iYears = Year(Date)
	    end if
      
      adoCmd.CommandText = "pb_GetPrescriptionTracking"
    	adoCmd.CommandType = &H0004
    	Set adoCmd.ActiveConnection = adoConn
      Set adoParam1 = adoCmd.CreateParameter("@P1_SSN", adInteger, &H0001)
    	adoCmd.Parameters.Append adoParam1
      Set adoParam2 = adoCmd.CreateParameter("@P2_DepNumber", adInteger, &H0001)
      adoCmd.Parameters.Append adoParam2
      Set adoParam3 = adoCmd.CreateParameter("@P3_HWType", adChar, &H0001, 1)
      adoCmd.Parameters.Append adoParam3
      Set adoParam4 = adoCmd.CreateParameter("@P4_Fund", adChar, &H0001, 3)
      adoCmd.Parameters.Append adoParam4
      Set adoParam5 = adoCmd.CreateParameter("@P5_Year", adInteger, &H0001)
      adoCmd.Parameters.Append adoParam5
      adoCmd("@P1_SSN") = m_iSSN
      adoCmd("@P2_DepNumber") = m_iDepNum
      adoCmd("@P3_HWType") = m_sHWType
      adoCmd("@P4_Fund") = m_sFund
      adoCmd("@P5_Year") = m_iYears
      
      adoRS.Open adoCmd
      
      Do While Not adoRS.EOF
        if adoRS("DepNumber") <> 99 and adoRS("HWType") = "P" then
          m_iPREYTD = formatcurrency(adoRS("Amount"))
        end if
    
		  adoRS.MoveNext
			Loop
      adoRS.Close
%>
    <center>
    <p><font face="verdana, arial, helvetica" size=2><b>Tracking information for<% if m_iDepNum = 0 then %> Participant <% else %> Dependant <% end if %><% if m_sName <> "" then %><%= m_sName %><% else %> NO NAME FOR THAT DEPENDANT NUMBER<% end if %>
    <p><font face="verdana, arial, helvetica" size=2><b>The year selected is <% if m_iYears <> "" then %><%= m_iYears %><% else %>EMPTY!<% end if %></b></font>
    <div style="height:68%; width:100%;">
		<p>&nbsp;
    <table WIDTH="80%" BORDER="1" bordercolor="white" bgcolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
			<tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
				<td bgcolor="#CCCCCC" align="right"><strong>OOP-IND</strong></td>
        <td align="right" STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><b><% if m_iOOPIND <> "" then %>&nbsp;<%= m_iOOPIND %><% else %>$0.00<% end if %></b></td>
        <td bgcolor="#CCCCCC" align="right"><strong>DED-IND</strong></td>
        <td align="right" STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><b><% if m_iDEDIND <> "" then %>&nbsp;<%= m_iDEDIND %><% else %>$0.00<% end if %></b></td>
        <td bgcolor="#CCCCCC" align="right"><strong>Group Life</strong></td>
        <td align="right" STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><b><% if m_iGrpLife <> "" then %>&nbsp;<%= m_iGrpLife %><% else %>$0.00<% end if %></b></td>
      </tr>
      <tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
        <td bgcolor="#CCCCCC" align="right"><strong>OOP-FAM</strong></td>
        <td align="right" STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><b><% if m_iOOPFAM <> "" then %>&nbsp;<%= m_iOOPFAM %><% else %>$0.00<% end if %></b></td>
        <td bgcolor="#CCCCCC" align="right"><strong>DED-FAM</strong></td>
        <td align="right" STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><b><% if m_iDEDFAM <> "" then %>&nbsp;<%= m_iDEDFAM %><% else %>$0.00<% end if %></b></td>
        <td bgcolor="#CCCCCC" align="right"><strong>Over Paid</strong></td>
        <td align="right" STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><b><% if m_iOVP <> "" then %>&nbsp;<%= m_iOVP %><% else %>$0.00<% end if %></b></td>
      </tr>
      <tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
        <td bgcolor="#CCCCCC" align="right"><strong>PPO-IND</strong></td>
        <td align="right" STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><b><% if m_iPPOIND <> "" then %>&nbsp;<%= m_iPPOIND %><% else %>$0.00<% end if %></b></td>
        <td bgcolor="#CCCCCC" align="right"><strong>4QTR-IND</strong></td>
        <td align="right" STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><b><% if m_iDE4IND <> "" then %>&nbsp;<%= m_iDE4IND %><% else %>$0.00<% end if %></b></td>
        <td bgcolor="#CCCCCC" align="right"><strong>COB(+/-)</strong></td>
        <td align="right" STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><b><% if m_iCOB <> "" then %>&nbsp;<%= m_iCOB %><% else %>$0.00<% end if %></b></td>
      </tr>
      <tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
        <td bgcolor="#CCCCCC" align="right"><strong>PPO-FAM</strong></td>
        <td align="right" STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><b><% if m_iPPOFAM <> "" then %>&nbsp;<%= m_iPPOFAM %><% else %>$0.00<% end if %></b></td>
        <td bgcolor="#CCCCCC" align="right"><strong>4QTR-FAM</strong></td>
        <td align="right" STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><b><% if m_iDE4FAM <> "" then %>&nbsp;<%= m_iDE4FAM %><% else %>$0.00<% end if %></b></td>
        <td bgcolor="#CCCCCC" align="right"><strong>Prescription-YTD</strong></td>
        <td align="right" STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><b><% if m_iPREYTD <> "" then %>&nbsp;<%= m_iPREYTD %><% else %>$0.00<% end if %></b></td>
      </tr>
    </table>
    
    <p>
    <table WIDTH="60%" BORDER="1" bordercolor="white" bgcolor="white" style="font:9pt verdana, arial, helvetica,sans-serif">
      <tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
        <td bgcolor="#CCCCCC" align="center"><strong>Plan</strong></td>
        <td bgcolor="#CCCCCC" align="center"><strong>TC</strong></td>
        <td bgcolor="#CCCCCC" align="center"><strong>Description</strong></td>
        <td bgcolor="#CCCCCC" align="center"><strong>Amount</strong></td>
        <td bgcolor="#CCCCCC" align="center"><strong>Number</strong></td>
      </tr>
<%  
      set adoCmd = Server.CreateObject("ADODB.Command")
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
      end if
      if Request.Querystring("Years") <> "" then
        m_iYears = Request.Querystring("Years")
	    end if
	    if Request.Querystring("HWType") <> "" then
	      m_sHWType = Request.Querystring("HWType")
      end if
      if Request.Querystring("Fund") <> "" then
		    m_sFund = Request.Querystring("Fund")
      end if
      
      adoCmd.CommandText = "pb_GetFinalTracking"
    	adoCmd.CommandType = &H0004
    	Set adoCmd.ActiveConnection = adoConn
      Set adoParam1 = adoCmd.CreateParameter("@P1_SSN", adInteger, &H0001)
    	adoCmd.Parameters.Append adoParam1
      Set adoParam2 = adoCmd.CreateParameter("@P2_Year", adInteger, &H0001)
    	adoCmd.Parameters.Append adoParam2
      Set adoParam3 = adoCmd.CreateParameter("@P3_HWType", adChar, &H0001, 1)
      adoCmd.Parameters.Append adoParam3
      Set adoParam4 = adoCmd.CreateParameter("@P4_Fund", adChar, &H0001, 3)
      adoCmd.Parameters.Append adoParam4
      Set adoParam5 = adoCmd.CreateParameter("@P5_DepNumber", adInteger, &H0001)
      adoCmd.Parameters.Append adoParam5
      adoCmd("@P1_SSN") = m_iSSN
      adoCmd("@P2_Year") = m_iYears
      adoCmd("@P3_HWType") = m_sHWType
      adoCmd("@P4_Fund") = m_sFund
      adoCmd("@P5_DepNumber") = m_iDepNum
      
      adoRS.Open adoCmd
				if not adoRS.EOF then
          Do While Not adoRS.EOF
            
%>
      <tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
        <td align="center" STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><b><%= adoRS("PlanCat") %></b></td>
        <td align="center" STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><b><%= adoRS("TrackCode") %></b></td>
        <td align="center" STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><a href="TrackingDetails.asp?SSN=<%= m_iSSN %>&DepNumber=<%= m_iDepNum %>&TrackCode=<%= adoRS("TrackCode") %>" onClick="WorkingStatus()"><b>&nbsp;<%= adoRS("Description") %></b></a></td>
        <td align="right" STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><b><%= formatcurrency(adoRS("Amount")) %></b></td>
        <td align="center" STYLE='color:<%= Session("EmphColor")%>;' bgcolor="#F0F0F0"><b><%= adoRS("Services") %></b></td>
      </tr>
<%  
         
          adoRS.MoveNext
        Loop
      else

%>
      <tr BorderColorDark="#F0F0F0" BorderColorlight="#999999">
        <td align="center" bgcolor="#F0F0F0" colspan="5"><b>There is no data for the year selected.</b></td>
      </tr>        
    </table>
    </div>
    </center>
<%  
      adoRS.Close
			adoConn.Close
			set adoRS = nothing
			set adoConn = nothing
      end if
      

%>
</body>
<!--#include file="VBFuncs.inc" -->
<script LANGUAGE="VBSCRIPT">
	
	sub ValidSend(iIndex)
	dim bValidData
	dim vbOKCancel, vbOk
  
  if iIndex = -2 then
			WorkingStatus
			ClearCrits.UserCriteria.value="Initial"
			ClearCrits.submit
			exit sub
		end if
  
  
	
		'vbOkCancel = 1
		'vbOk = 1

		Criteria.SSN.Value = trim(Criteria.SSN.Value)
    Criteria.DepNum.Value = trim(Criteria.DepNum.Value)
    Criteria.Years.Value = trim(Criteria.Years.Value)
		Criteria.HWType.Value = trim(Criteria.HWType.Value)
		Criteria.Fund.Value = trim(Criteria.Fund.Value)
    
    if iIndex = 0 and (Criteria.HWType.Value = "" or Criteria.Fund.Value = "") then
			msgbox "You must enter both a HWType and Fund for your query to run!",16,"Missing Data"
      exit sub
		end if

		
		WorkingStatus
		if "<%= m_iSSN%>" <> Criteria.SSN.Value or _
        "<%= m_iDepNum%>" <> Criteria.DepNum.Value or _
        "<%= m_iYears%>" <> Criteria.Years.Value or _
				"<%= m_sHWType%>" <> Criteria.HWType.Value or _
				"<%= m_sFund%>" <> Criteria.Fund.Value then
			Criteria.ActionType.value=""
		else
			Criteria.ActionType.value=iIndex
		end if
		Criteria.submit
	end sub
	
</script>

</html>
