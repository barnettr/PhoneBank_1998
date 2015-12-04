<%@  language="VBSCRIPT" %>
<% Option Explicit %>
<%Response.Expires = 0 %>
<!--#include file="user.inc" -->
<!--#include file="adovbs_mod.inc" -->
<%
Dim adoConn, adoRS, sSQL, sTemp, sCriteria
Dim iPageNo, iRowCount, i, iPageCount
dim  m_iSSN, m_sFrom, m_sThru, m_sACSID, m_sStatus, m_sStatusFrom, m_sStatusThru, m_sFund, m_sHWType, m_sName1, m_sName2
dim bUseSQL, adoParam1, adoParam2, adoParam3, adoParam4, adoParam5, adoParam6, adoParam7, adoParam8, adoParam9, adoParam10
dim adoParam11, adoCmd
dim Manager, Supervisor, Auditor, locktype, Admin

if Request.Querystring("UserCriteria") = "Initial" then
	bUseSQL = false
else
	bUseSQL = true
	
  set adoConn = Server.CreateObject("ADODB.Connection")
  adoConn.ConnectionTimeout = 0
	adoConn.CommandTimeout = 0
  set adoCmd = Server.CreateObject("ADODB.Command")
  adoCmd.CommandTimeout = 0
	set adoRS = Server.CreateObject("ADODB.Recordset")
	adoConn.Open Application("DataConn")

  sSQL = "select * from UserInformation where LogonID='" & Session("User") & "'"
  adoRS.Open sSql, adoConn, adOpenForwardOnly, adLockOptimistic

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
  
  if Request.Querystring("SSN") <> "" then
		m_iSSN = Request.Querystring("SSN")
	end if
  sSQL = "select locktype, SupervisorAccessOnly from Names where SSN=" & m_iSSN & " And DepNumber > 0 and locktype is not null"
  adoRS.Open sSql, adoConn, adOpenForwardOnly, adLockOptimistic
  
  if not adoRS.EOF then
    locktype = adoRS("locktype")
    Admin = adoRS("SupervisorAccessOnly")
  end if
  
  adoRS.Close
	
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
	if Request.Querystring("From") <> "" then
		m_sFrom = Request.Querystring("From")
	end if
	if Request.Querystring("Thru") <> "" then
		m_sThru = Request.Querystring("Thru")
	end if
	if Request.Querystring("SSN") <> "" then
		m_iSSN = Request.Querystring("SSN")
	end if
	if Request.Querystring("ACSID") <> "" then
		m_sACSID = Request.Querystring("ACSID")
	end if
	'if Request.Querystring("Status") <> "" then
		'm_sStatus = Request.Querystring("Status")
	'end if
	if Request.Querystring("StatusFrom") <> "" then
		m_sStatusFrom = Request.Querystring("StatusFrom")
	end if
	if Request.Querystring("StatusThru") <> "" then
		m_sStatusThru = Request.Querystring("StatusThru")
	end if
	if Request.Querystring("Fund") <> "" then
		m_sFund = Request.Querystring("Fund")
	end if
	if Request.Querystring("HWType") <> "" then
		m_sHWType = Request.Querystring("HWType")
	end if
	if Request.Querystring("Name1") <> "" then
		m_sName1 = Request.Querystring("Name1")
	end if
	if Request.Querystring("Name2") <> "" then
		m_sName2 = Request.Querystring("Name2")
	end if
	
	adoCmd.CommandText = "PB_GetLettersSent"
    adoCmd.CommandType = adCmdStoredProc
    Set adoCmd.ActiveConnection = adoConn
    Set adoParam1 = adoCmd.CreateParameter("@P1_DateCreatedFrom", adDBTimeStamp, adParamInput)
    adoCmd.Parameters.Append adoParam1
    Set adoParam2 = adoCmd.CreateParameter("@P2_DateCreatedThru", adDBTimeStamp, adParamInput)
    adoCmd.Parameters.Append adoParam2
    Set adoParam3 = adoCmd.CreateParameter("@P3_StatusDateFrom", adDBTimeStamp, adParamInput)
    adoCmd.Parameters.Append adoParam3
    Set adoParam4 = adoCmd.CreateParameter("@P4_StatusDateThru", adDBTimeStamp, adParamInput)
    adoCmd.Parameters.Append adoParam4
    Set adoParam5 = adoCmd.CreateParameter("@P5_SSN", adInteger, adParamInput)
    adoCmd.Parameters.Append adoParam5
    Set adoParam6 = adoCmd.CreateParameter("@P6_ACSID", adVarChar, adParamInput, 6)
    adoCmd.Parameters.Append adoParam6
    Set adoParam7 = adoCmd.CreateParameter("@P7_Status", adChar, adParamInput, 1)
    adoCmd.Parameters.Append adoParam7
    Set adoParam8 = adoCmd.CreateParameter("@P8_Fund", adChar, adParamInput, 3)
    adoCmd.Parameters.Append adoParam8
    Set adoParam9 = adoCmd.CreateParameter("@P9_HWType", adChar, adParamInput, 1)
    adoCmd.Parameters.Append adoParam9
    Set adoParam10 = adoCmd.CreateParameter("@P10_Name1", adVarChar, adParamInput, 30)
    adoCmd.Parameters.Append adoParam10
    Set adoParam11 = adoCmd.CreateParameter("@P11_Name2", adVarChar, adParamInput, 30)
    adoCmd.Parameters.Append adoParam11
    adoCmd("@P1_DateCreatedFrom") = m_sFrom
    adoCmd("@P2_DateCreatedThru") = m_sThru
    adoCmd("@P3_StatusDateFrom") = m_sStatusFrom
    adoCmd("@P4_StatusDateThru") = m_sStatusThru
    adoCmd("@P5_SSN") = m_iSSN
    adoCmd("@P6_ACSID") = m_sACSID
    adoCmd("@P7_Status") = m_sStatus
    adoCmd("@P8_Fund") = m_sFund
    adoCmd("@P9_HWType") = m_sHWType
    adoCmd("@P10_Name1") = m_sName1
    adoCmd("@P11_Name2") = m_sName2


end if
%>
<html>
<head>
    <title>Details</title>
</head>
<body topmargin="2" leftmargin="2" rightmargin="0" language="VBScript" onload="UpdateScreen(3)">
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
                <font size="+2" face="verdana, arial, helvetica"><strong>ACS Letter Search</strong></font>
            </td>
            <td align="CENTER" width="20%">
                <img src="images/bluebar2.gif" onclick="history.go(-1)">&nbsp;&nbsp;<img src="images/bluebar3.gif"
                    onclick="history.go(+1)" border="0">
            </td>
        </tr>
    </table>
    <%

	SetPage bUseSQL
	
	Sub SetPage(bOpenRS)

		if bOpenRS then
			adoRS.Open adoCmd
      'adoRS.Open sSql, adoConn, adOpenKeyset, adLockOptimistic
			if adoRS.EOF then
				adoRS.Close
				adoConn.Close
				set adoRS = nothing
				set adoConn = nothing
				Response.write ("<ul><p><font size=2 color='red' face='verdana, arial, helvetica'><b>No matches were found -- try a different search.</b></font></ul>")
				exit sub
			end if
			adoRS.PageSize = m_iPageSize ' Number of rows per page
			iPageCount = adoRS.PageCount
			'adoRS.AbsolutePage = iPageNo
		end if
    %>
    <table cols="3" width="100%" cellpadding="0" cellspacing="0" border="0">
        <tr>
            <td width="10%">
                <input type="IMAGE" src="images/query_button.gif" accesskey="S" name="NewQuery" value="Start Query"
                    onclick="ValidSend(0)">
            </td>
            <td width="65%">
                <font face="verdana, arial, helvetica" size="2"><b>Enter your criteria, and then click
                    the Start Query button or key (ALT+S).</b></font>
            </td>
            <td align="RIGHT" name="MorePages" id="MorePages">
                &nbsp;
                <%
				if bUseSQL and iPageCount > 1 then
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
					sTemp = sTemp & "</TD>"
					Response.Write sTemp
				end if
                %>
        </tr>
    </table>
    <div style="height=24%; width: 100%; overflow=auto;">
        <form id="Criteria" method="GET" action="LetterSearch.asp" target="Details">
        <table class="CriteriaTable" cols="6" cellpadding="0" cellspacing="0" border="0">
            <tr>
                <!-- These TH's were required in order to get IE to deal properly with the DIV...among other things, without them and with BORDER=1, things were displayed OK, butwith the desired BORDER=0 IE seemed to 'lose' a number of the elements-->
                <th>
                </th>
                <th>
                </th>
                <th>
                </th>
                <th>
                </th>
                <th>
                </th>
                <th>
                </th>
            </tr>
            <tr>
                <td class="White" align="right">
                    SSN:
                </td>
                <td>
                    <input type="TEXT" id="SSN" name="SSN" size="10" maxlength="10" value="<%= m_iSSN%>">
                </td>
                <td class="White" align="right">
                    Primary Name:
                </td>
                <td>
                    <input type="TEXT" id="Name1" name="Name1" size="15" maxlength="30" value="<%= m_sName1%>">
                </td>
                <td class="White" align="right">
                    Secondary Name:
                </td>
                <td>
                    <input type="TEXT" id="Name2" name="Name2" size="15" maxlength="30" value="<%= m_sName2%>">
                </td>
            </tr>
            <tr>
                <td class="White" align="right">
                    ACS User ID:
                </td>
                <td>
                    <input type="TEXT" id="ACSID" name="ACSID" size="10" maxlength="6" value="<%= m_sACSID%>">
                </td>
                <td class="White" align="right">
                    HWType:
                </td>
                <td>
                    <input type="TEXT" id="HWType" name="HWType" size="3" maxlength="1" value="<%= m_sHWType%>">
                </td>
                <td class="White" align="right">
                    Fund:
                </td>
                <td>
                    <input type="TEXT" id="Fund" name="Fund" size="5" maxlength="3" value="<%= m_sFund%>">
                </td>
            </tr>
            <tr>
                <td class="White">
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
                <td class="White" align="right">
                    Status Date From:
                </td>
                <td>
                    <input type="TEXT" id="StatusFrom" name="StatusFrom" size="10" maxlength="10" value="<%= m_sStatusFrom%>">
                </td>
                <td class="White" align="right">
                    Status Date Thru:
                </td>
                <td>
                    <input type="TEXT" id="StatusThru" name="StatusThru" size="10" maxlength="10" value="<%= m_sStatusThru%>">
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                </td>
                <td class="White" align="right">
                    Creation Date From:
                </td>
                <td>
                    <input type="TEXT" id="From" name="From" size="10" maxlength="10" value="<%= m_sFrom%>">
                </td>
                <td class="White" align="right">
                    Creation Date Thru:
                </td>
                <td>
                    <input type="TEXT" id="Thru" name="Thru" size="10" maxlength="10" value="<%= m_sThru%>">
                </td>
            </tr>
        </table>
        <input type="HIDDEN" id="ActionType" name="ActionType" value>
        </form>
    </div>
    <br clear="LEFT">
    <%
		if bOpenRS then
      if locktype <> "" then
        If Manager = "True" or Supervisor = "True" or Auditor = "True" Then
    %>
    <table width="100%" border="0">
        <tr>
            <td>
                <font size="2" face="verdana, arial, helvetica"><b>
                    <% if locktype <> "" then %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src="images/lock1.gif"
                        width="32" height="32" alt="LOCKED FILE!!"><font size="2" face="verdana, arial, helvetica"
                            color="red"><b>&nbsp;&nbsp;&nbsp;LOCKED "<%= locktype %>" RECORDS!!!</b></font><% end if %>
            </td>
        </tr>
    </table>
    <div style="height=65%; width=100%; overflow=auto">
        <table width="100%" border="1" bordercolor="white" bgcolor="white" style="font: 9pt verdana, arial, helvetica,sans-serif">
            <tr bordercolordark="#F0F0F0" bordercolorlight="#999999">
                <td align="CENTER" nowrap bgcolor="#cccccc">
                    <font color="BLUE"><b>Created</b></font>
                </td>
                <td align="CENTER" nowrap bgcolor="#cccccc">
                    <font color="BLUE"><b>Sequence</b></font>
                </td>
                <td align="CENTER" nowrap bgcolor="#cccccc">
                    <font color="BLUE"><b>Status</b></font>
                </td>
                <td align="CENTER" nowrap bgcolor="#cccccc">
                    <font color="BLUE"><b>Status Date</b></font>
                </td>
                <td align="CENTER" nowrap bgcolor="#cccccc">
                    <font color="BLUE"><b>SSN</b></font>
                </td>
                <td align="CENTER" nowrap bgcolor="#cccccc">
                    <font color="BLUE"><b>Fund</b></font>
                </td>
                <td align="CENTER" nowrap bgcolor="#cccccc">
                    <font color="BLUE"><b>HWType</b></font>
                </td>
                <td align="CENTER" nowrap bgcolor="#cccccc">
                    <font color="BLUE"><b>LetterKey</b></font>
                </td>
                <td align="CENTER" nowrap bgcolor="#cccccc">
                    <font color="BLUE"><b>ACS User ID</b></font>
                </td>
                <td align="CENTER" nowrap bgcolor="#cccccc">
                    <font color="BLUE"><b>Name</b></font>
                </td>
                <td align="CENTER" nowrap bgcolor="#cccccc">
                    <font color="BLUE"><b>2nd Name</b></font>
                </td>
            </tr>
            <% 
					iRowCount = adoRS.PageSize
          Do While Not adoRS.EOF and iRowCount > 0 
            %>
            <tr bordercolordark="#FCFCFC" bordercolorlight="#CCCCCC">
                <td nowrap bgcolor="#F0F0F0">
                    <a href="LetterDetails.asp?DateCreated=<%= adoRS("Created")%>&amp;Seq=<%= adoRS("Seq#")%>"
                        onclick="WorkingStatus()">
                        <%= adoRS("Created") %></a>
                </td>
                <td nowrap bgcolor="#F0F0F0">
                    <%= adoRS("Seq#") %>
                </td>
                <td nowrap bgcolor="#F0F0F0">
                    <%= adoRS("Status") %>
                </td>
                <td nowrap bgcolor="#F0F0F0">
                    <%= adoRS("StatusDate") %>
                </td>
                <td nowrap bgcolor="#F0F0F0">
                    <%= adoRS("SSN") %>
                </td>
                <td nowrap bgcolor="#F0F0F0">
                    <%= adoRS("Fund") %>
                </td>
                <td nowrap bgcolor="#F0F0F0">
                    <%= adoRS("HWType") %>
                </td>
                <td nowrap bgcolor="#F0F0F0">
                    <%= adoRS("LetterKey") %>
                </td>
                <td nowrap bgcolor="#F0F0F0">
                    <%= adoRS("ACSUserID") %>
                </td>
                <td nowrap bgcolor="#F0F0F0">
                    <%= adoRS("Name") %>
                </td>
                <td nowrap bgcolor="#F0F0F0">
                    <%= adoRS("2nd Name") %>
                </td>
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
		    else 'Manager is not true
    %>
    <div style="height=65%; width=100%; overflow=auto">
        <table width="100%" border="1" bordercolor="white" bgcolor="white" style="font: 10pt verdana, arial, helvetica, sans-serif">
            <tr bordercolordark="#FCFCFC" bordercolorlight="#CCCCCC">
                <td align="CENTER" nowrap bgcolor="#F0F0F0">
                    <font color="Red"><b>The Participant or one of the Dependents has a Locked File. Letters
                        are not available. Please see your Supervisor.</b></font>
                </td>
            </tr>
        </table>
    </div>
    <%  
        end if
      else 'locktype is not true
    %>
    <div style="height=65%; width=100%; overflow=auto">
        <table width="100%" border="1" bordercolor="white" bgcolor="white" style="font: 9pt verdana, arial, helvetica,sans-serif">
            <tr bordercolordark="#F0F0F0" bordercolorlight="#999999">
                <td align="CENTER" nowrap bgcolor="#cccccc">
                    <font color="BLUE"><b>Created</b></font>
                </td>
                <td align="CENTER" nowrap bgcolor="#cccccc">
                    <font color="BLUE"><b>Sequence</b></font>
                </td>
                <td align="CENTER" nowrap bgcolor="#cccccc">
                    <font color="BLUE"><b>Status</b></font>
                </td>
                <td align="CENTER" nowrap bgcolor="#cccccc">
                    <font color="BLUE"><b>Status Date</b></font>
                </td>
                <td align="CENTER" nowrap bgcolor="#cccccc">
                    <font color="BLUE"><b>SSN</b></font>
                </td>
                <td align="CENTER" nowrap bgcolor="#cccccc">
                    <font color="BLUE"><b>Fund</b></font>
                </td>
                <td align="CENTER" nowrap bgcolor="#cccccc">
                    <font color="BLUE"><b>HWType</b></font>
                </td>
                <td align="CENTER" nowrap bgcolor="#cccccc">
                    <font color="BLUE"><b>LetterKey</b></font>
                </td>
                <td align="CENTER" nowrap bgcolor="#cccccc">
                    <font color="BLUE"><b>ACS User ID</b></font>
                </td>
                <td align="CENTER" nowrap bgcolor="#cccccc">
                    <font color="BLUE"><b>Name</b></font>
                </td>
                <td align="CENTER" nowrap bgcolor="#cccccc">
                    <font color="BLUE"><b>2nd Name</b></font>
                </td>
            </tr>
            <% 
					iRowCount = adoRS.PageSize
          Do While Not adoRS.EOF and iRowCount > 0 
            %>
            <tr bordercolordark="#FCFCFC" bordercolorlight="#CCCCCC">
                <td nowrap bgcolor="#F0F0F0">
                    <a href="LetterDetails.asp?DateCreated=<%= adoRS("Created")%>&amp;Seq=<%= adoRS("Seq#")%>"
                        onclick="WorkingStatus()">
                        <%= adoRS("Created") %></a>
                </td>
                <td nowrap bgcolor="#F0F0F0">
                    <%= adoRS("Seq#") %>
                </td>
                <td nowrap bgcolor="#F0F0F0">
                    <%= adoRS("Status") %>
                </td>
                <td nowrap bgcolor="#F0F0F0">
                    <%= adoRS("StatusDate") %>
                </td>
                <td nowrap bgcolor="#F0F0F0">
                    <%= adoRS("SSN") %>
                </td>
                <td nowrap bgcolor="#F0F0F0">
                    <%= adoRS("Fund") %>
                </td>
                <td nowrap bgcolor="#F0F0F0">
                    <%= adoRS("HWType") %>
                </td>
                <td nowrap bgcolor="#F0F0F0">
                    <%= adoRS("LetterKey") %>
                </td>
                <td nowrap bgcolor="#F0F0F0">
                    <%= adoRS("ACSUserID") %>
                </td>
                <td nowrap bgcolor="#F0F0F0">
                    <%= adoRS("Name") %>
                </td>
                <td nowrap bgcolor="#F0F0F0">
                    <%= adoRS("2nd Name") %>
                </td>
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
      end if
    end if
	end sub
    %>
</body>
<!--#include file="VBFuncs.inc" -->
<script language="VBSCRIPT">
	
	sub ValidSend(iIndex)
	dim bValidData

		Criteria.SSN.Value = trim(Criteria.SSN.Value)
		Criteria.ACSID.Value = trim(Criteria.ACSID.Value)
		Criteria.Fund.Value = trim(Criteria.Fund.Value)
		Criteria.HWType.Value = trim(Criteria.HWType.Value)
		Criteria.Name1.Value = trim(Criteria.Name1.Value)
		Criteria.Name2.Value = trim(Criteria.Name2.Value)
		Criteria.From.Value = trim(Criteria.From.Value)
		Criteria.Thru.Value = trim(Criteria.Thru.Value)
		'Criteria.Status.Value = trim(Criteria.Status.Value)
		Criteria.StatusFrom.Value = trim(Criteria.StatusFrom.Value)
		Criteria.StatusThru.Value = trim(Criteria.StatusThru.Value)
		
		bValidData = false
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
		if Criteria.ACSID.Value <> "" then
			if ContainsInvalids(Criteria.ACSID.Value) then
				msgbox "Please remove invalid characters from the ACS ID field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.Fund.Value <> "" then
			if ContainsInvalids(Criteria.Fund.Value) then
				msgbox "Please remove invalid characters from the Fund field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.HWType.Value <> "" then
			if ContainsInvalids(Criteria.HWType.Value) then
				msgbox "Please remove invalid characters from the HWType field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.Name1.Value <> "" then
			if ContainsInvalids(Criteria.Name1.Value) then
				msgbox "Please remove invalid characters from the Primary Name field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.Name2.Value <> "" then
			if ContainsInvalids(Criteria.Name2.Value) then
				msgbox "Please remove invalid characters from the Secondary Name field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.From.Value <> "" then
			if not isDate(Criteria.From.Value) then
				msgbox "Please enter a valid date for the Creation From field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.Thru.Value <> "" then
			if not isDate(Criteria.Thru.Value) then
				msgbox "Please enter a valid date for the Creation Thru field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		'if Criteria.Status.Value <> "" then
			'if ContainsInvalids(Criteria.Status.Value) then
				'msgbox "Please remove invalid characters from the Status field.",,"Invalid Data"
				'exit sub
			'end if
			'bValidData = true
		'end if
		if Criteria.StatusFrom.Value <> "" then
			if not isDate(Criteria.StatusFrom.Value) then
				msgbox "Please enter a valid date for the Status From field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.StatusThru.Value <> "" then
			if not isDate(Criteria.StatusThru.Value) then
				msgbox "Please enter a valid date for the Status Thru field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if not bValidData then
			msgbox "You haven't entered anything to search for.",,"No Criteria"
			exit sub
		end if
		
		WorkingStatus
		if "<%= m_iSSN%>" <> Criteria.SSN.Value or _
			"<%= m_sACSID%>" <> Criteria.ACSID.Value or _
			"<%= m_sFund%>" <> Criteria.Fund.Value or _
			"<%= m_sHWType%>" <> Criteria.HWType.Value or _
			"<%= m_sName1%>" <> Criteria.Name1.Value or _
			"<%= m_sName2%>" <> Criteria.Name2.Value or _
			"<%= m_sFrom%>" <> Criteria.From.Value or _
			"<%= m_sThru%>" <> Criteria.Thru.Value or _
			"<%= m_sStatusFrom%>" <> Criteria.StatusFrom.Value or _
			"<%= m_sStatusThru%>" <> Criteria.StatusThru.Value then
			Criteria.ActionType.value=""
		else
			Criteria.ActionType.value=iIndex
		end if
		Criteria.submit
	end sub
	
</script>
</html>
