<%@  language="VBScript" %>
<% Option Explicit %>
<% Response.Expires = 0 %>
<!--#include file="user.inc" -->
<!--#include file="adovbs_mod.inc" -->
<%
Dim i, adoConn, adoRS, sSQL, sTemp
dim m_dLogDate, m_iSequence, m_iDistr
Dim m_sFund

m_dLogDate = Request.QueryString("LogDate")
m_iSequence = Request.QueryString("Sequence")
m_iDistr = Request.QueryString("Distr")

if m_dLogDate = "" or m_iSequence = "" or m_iDistr = "" then
	Response.Write "Log Date, Sequence, or Distr missing; contact your network administrator"
else
	set adoConn = Server.CreateObject("ADODB.Connection")
'**********************************************************************************
' Setting CommandTimeout to 0 waits indefinitely for response; note that per 
' ADO documentation Command object will not inherit the Connection setting (default
' is 30 seconds).  This appeared to help with response issues on seasql02.
'**********************************************************************************
	adoConn.CommandTimeout = 0
	set adoRS = Server.CreateObject("ADODB.Recordset")
	adoConn.Open Application("DataConn")
	
	'sSQL = "select c.LogDate, c.Sequence, c.Distr, c.SSN, c.Fund, c.PlanCode, c.HWType,"
	'sSQL = sSQL & " c.PayCode, c.Status, c.StatusDate, c.PPOSponsor, c.PPOOverride,"
	'sSQL = sSQL & " c.COBIndicator, c.OverpaymentAdjustment, c.OtherAdjustment, c.BillDate,"
	'sSQL = sSQL & " d.DiagnosisCode, h.BillType, h.AdmitType, h.DRG, h.Admitdate,"
	'sSQL = sSQL & " h.DischargeDate, "
	'sSQL = sSQL & " c.DepNumber, n.FirstName + ' ' + n.MiddleName + ' ' "
	'sSQL = sSQL & " + n.LastName Name, p.FullName Provider, c.TaxID, c.Associate"
	'sSQL = sSQL & " from Claims c left join HospClaim h on (c.LogDate="
	'sSQL = sSQL & " h.LogDate AND c.Sequence=h.Sequence AND c.Distr=h.Distr) left join Names n on" 
	'sSQL = sSQL & " (c.SSN=n.SSN and c.DepNumber=n.depnumber) left join Providers P on (c.TaxID=p.TaxID"
	'sSQL = sSQL & " and c.Associate=p.Associate and p.RecordType=0) left join ClaimDiagnosis d on (c.LogDate="
	'sSQL = sSQL & " d.LogDate AND c.Sequence=d.Sequence AND c.Distr=d.Distr)"
	'sSQL = sSQL & " where c.LogDate='" & m_dLogDate & "' AND c.Sequence=" & m_iSequence & " AND c.Distr= " & m_iDistr
  sSQL = "PB_LookUpClaim '" & m_dLogDate & "', " & m_iSequence & ", " & m_iDistr
  'response.write sSQL
%>
<html>
<head>
    <!--#include file="VBFuncs.inc" -->
</head>
<body language="VBScript" onload="UpdateScreen(3)">
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
                <font size="+2" face="verdana, arial, helvetica"><strong>Claim Information</strong></font>
            </td>
            <td align="CENTER" width="20%">
                <img src="images/bluebar2.gif" onclick="history.back">&nbsp;&nbsp;<img src="images/bluebar3.gif"
                    onclick="history.go(+1)" border="0">
            </td>
        </tr>
    </table>
    <%
	Set adoRS = adoConn.execute(sSQL)
  'adoRS.Open sSQL, adoConn, adOpenForwardOnly, adLockOptimistic
	if adoRS.EOF then
		Response.Write ("<p><ul><font color='red' face='verdana, arial, helvetica'>Unable to find all necessary information; please contact your network administrator.</font></ul>")
	else
		m_sFund = adoRS("Fund")
    %>
    <table cols="6" width="100%" border="1" rules="GROUPS" bordercolor="white" bgcolor="white"
        style="font: 9pt verdana, arial, helvetica,sans-serif">
        <colgroup span="2">
            <colgroup span="2">
                <colgroup span="2">
                    <tr bordercolordark="#F0F0F0" bordercolorlight="#999999">
                        <td width="10%" bgcolor="#cccccc">
                            <strong>Log Date:</strong>
                        </td>
                        <td style='color: <%= Session("EmphColor")%>;' width="10%" bgcolor="#cccccc">
                            <%= adoRS("LogDate")%>
                        </td>
                        <td width="13%" bgcolor="#cccccc">
                            <strong>Name:</strong>
                        </td>
                        <%
						sTemp = "<td WIDTH='25%' bgcolor='#cccccc'><A HREF='PersonDetails.asp?SSN=" & adoRS("SSN") & "&DepNo=" & adoRS("DepNumber")
						sTemp = sTemp & "' onClick='javascript:return LogCheck(" & adoRS("SSN") & ")'>"
						sTemp = sTemp & adoRS("Name") & "</A></td>"
						Response.Write sTemp
                        %>
                        <td width="13%" bgcolor="#cccccc">
                            <strong>Provider:</strong>
                        </td>
                        <td style='color: <%= Session("EmphColor")%>;' width="28%" bgcolor="#cccccc">
                            <a href="ProviderDetails.asp?TaxID=<%= adoRS("TaxID")%>&amp;AssocNo=<%= adoRS("Associate")%>"
                                onclick="WorkingStatus()">
                                <%= adoRS("Provider")%></a>
                        </td>
                    </tr>
                    <tr bordercolordark="#FCFCFC" bordercolorlight="#CCCCCC">
                        <td bgcolor="#F0F0F0">
                            <strong>Sequence:</strong>
                        </td>
                        <td style='color: <%= Session("EmphColor")%>;' bgcolor="#F0F0F0">
                            <%= adoRS("Sequence")%>
                        </td>
                        <td bgcolor="#F0F0F0">
                            <strong>SSN:</strong>
                        </td>
                        <td style='color: <%= Session("EmphColor")%>;' bgcolor="#F0F0F0">
                            <%= adoRS("SSN")%>
                        </td>
                        <td bgcolor="#F0F0F0">
                            <strong>TaxID:</strong>
                        </td>
                        <td style='color: <%= Session("EmphColor")%>;' bgcolor="#F0F0F0">
                            <%= adoRS("TaxID")%>
                        </td>
                    </tr>
                    <tr bordercolordark="#F0F0F0" bordercolorlight="#999999">
                        <td bgcolor="#cccccc">
                            <strong>Distr:</strong>
                        </td>
                        <td style='color: <%= Session("EmphColor")%>;' bgcolor="#cccccc">
                            <%= adoRS("Distr")%>
                        </td>
                        <td bgcolor="#cccccc">
                            <strong>Dependent #:</strong>
                        </td>
                        <td style='color: <%= Session("EmphColor")%>;' bgcolor="#cccccc">
                            <%= adoRS("DepNumber")%>
                        </td>
                        <td bgcolor="#cccccc">
                            <strong>Associate #:</strong>
                        </td>
                        <td style='color: <%= Session("EmphColor")%>;' bgcolor="#cccccc">
                            <%= adoRS("Associate")%>
                        </td>
                    </tr>
    </table>
    <br />
    <table width="100%" cols="8" cellpadding="1" cellspacing="1" border="1" bordercolor="white"
        bgcolor="white" style="font: 9pt verdana, arial, helvetica,sans-serif">
        <tr bordercolordark="#F0F0F0" bordercolorlight="#999999">
            <td bgcolor="#cccccc">
                <strong>Group:</strong>
            </td>
            <td style='color: <%= Session("EmphColor")%>;' bgcolor="#cccccc">
                <%= adoRS("Fund")%>
            </td>
            <td bgcolor="#cccccc">
                <strong>Status:</strong>
            </td>
            <td style='color: <%= Session("EmphColor")%>;' bgcolor="#cccccc">
                <%= adoRS("Status")%>
            </td>
            <td bgcolor="#cccccc">
                <strong>PPO Sponsor:</strong>
            </td>
            <td style='color: <%= Session("EmphColor")%>;' bgcolor="#cccccc">
                <%= adoRS("PPOSponsor")%>
            </td>
            <td bgcolor="#cccccc">
                <strong>Pay Code:</strong>
            </td>
            <td style='color: <%= Session("EmphColor")%>;' bgcolor="#cccccc">
                <%= adoRS("PayCode")%>
            </td>
        </tr>
        <tr bordercolordark="#FCFCFC" bordercolorlight="#CCCCCC">
            <td bgcolor="#F0F0F0">
                <strong>Plan:</strong>
            </td>
            <td style='color: <%= Session("EmphColor")%>;' bgcolor="#F0F0F0">
                <%= adoRS("PlanCode")%>
            </td>
            <td bgcolor="#F0F0F0">
                <strong>Status Date:</strong>
            </td>
            <td style='color: <%= Session("EmphColor")%>;' bgcolor="#F0F0F0">
                <%= adoRS("StatusDate")%>
            </td>
            <td bgcolor="#F0F0F0">
                <strong>PPO Override:</strong>
            </td>
            <td style='color: <%= Session("EmphColor")%>;' bgcolor="#F0F0F0">
                <%= adoRS("PPOOverride")%>
            </td>
            <td bgcolor="#F0F0F0">
                <strong>Overpayment Adjustment:</strong>
            </td>
            <%
					if isnull(adoRS("OverpaymentAdjustment")) then
						sTemp = "<td bgcolor='#F0F0F0'></td>"
					else
						sTemp = "<td STYLE='color:" &  Session("EmphColor") & ";' bgcolor='#F0F0F0'>" & formatcurrency(adoRS("OverpaymentAdjustment")) & "</td>"
					end if
					Response.Write sTemp
            %>
        </tr>
        <tr bordercolordark="#F0F0F0" bordercolorlight="#999999">
            <td bgcolor="#cccccc">
                <strong>HW Type:</strong>
            </td>
            <td style='color: <%= Session("EmphColor")%>;' bgcolor="#cccccc">
                <%= adoRS("HWType")%>
            </td>
            <td bgcolor="#cccccc">
                <strong>Diagnosis Code:</strong>
            </td>
            <td style='color: <%= Session("EmphColor")%>;' bgcolor="#cccccc">
                <%= adoRS("DiagnosisCode")%>
            </td>
            <td bgcolor="#cccccc">
                <strong>COB Indicator:</strong>
            </td>
            <td style='color: <%= Session("EmphColor")%>;' bgcolor="#cccccc">
                <%= adoRS("COBIndicator")%>
            </td>
            <td bgcolor="#cccccc">
                <strong>Other Adjustment:</strong>
            </td>
            <%
					if isnull(adoRS("OtherAdjustment")) then
						sTemp = "<td bgcolor='#cccccc';></td>"
					else
						sTemp = "<td STYLE='color:" &  Session("EmphColor") & ";' bgcolor='#cccccc'>" & formatcurrency(adoRS("OtherAdjustment")) & "</td>"
					end if
					Response.Write sTemp
            %>
        </tr>
    </table>
    <br />
    <hr size="3" noshade />
    <%
'
' Hospital Information
'
    %>
    <br />
    
        <font size="3" face="arial, helvetica">
            <center>
                <b>Hospital Information</b></center>
        </font>
        <table cellpadding="2" cellspacing="2" width="100%" border="0" bgcolor="white" bordercolor="white"
            style="font: 9pt verdana, arial, helvetica,sans-serif">
            <tr align="CENTER" bordercolordark="#F0F0F0" bordercolorlight="#999999">
                <td bgcolor="#CCCCCC">
                    <strong>Bill Type:</strong>
                </td>
                <td bgcolor="#CCCCCC">
                    &nbsp;<%= adoRS("BillType")%>
                </td>
                <td bgcolor="#CCCCCC">
                    <strong>Admit Type:</strong>
                </td>
                <td bgcolor="#CCCCCC">
                    &nbsp;<%= adoRS("AdmitType")%>
                </td>
                <td bgcolor="#CCCCCC">
                    <strong>Admit Date:</strong>
                </td>
                <td bgcolor="#CCCCCC">
                    &nbsp;<%= adoRS("AdmitDate")%>
                </td>
                <td bgcolor="#CCCCCC">
                    <strong>Discharge Date:</strong>
                </td>
                <td bgcolor="#CCCCCC">
                    &nbsp;<%= adoRS("DischargeDate")%>
                </td>
                <td bgcolor="#CCCCCC">
                    <strong>DRG</strong>:
                </td>
                <td bgcolor="#CCCCCC">
                    &nbsp;<%= adoRS("DRG")%>
                </td>
            </tr>
        </table>
        <%
			adoRS.Close
'
' Charges
'
        %>
        <br />
        <font size="3" face="arial, helvetica">
            <center>
                <b>Chargelines</b></center>
        </font>
        <table width="100%" border="1" cellpadding="2" cellspacing="2" bgcolor="white" bordercolor="white"
            style="font: 9pt verdana, arial, helvetica,sans-serif">
            <tr align="CENTER" bordercolordark="#F0F0F0" bordercolorlight="#999999">
                <td bgcolor="#CCCCCC">
                    <strong>Line #</strong>
                </td>
                <td bgcolor="#CCCCCC">
                    <strong>Coverage</strong>
                </td>
                <td bgcolor="#CCCCCC">
                    <strong>Proc. Code</strong>
                </td>
                <td bgcolor="#CCCCCC">
                    <strong>Place Of Service</strong>
                </td>
                <td bgcolor="#CCCCCC">
                    <strong>From Date</strong>
                </td>
                <td bgcolor="#CCCCCC">
                    <strong>Thru Date</strong>
                </td>
                <td bgcolor="#CCCCCC">
                    <strong>Charge</strong>
                </td>
                <td bgcolor="#CCCCCC">
                    <strong>Track Code</strong>
                </td>
                <td bgcolor="#CCCCCC">
                    <strong>Revenue Code</strong>
                </td>
                <td bgcolor="#CCCCCC">
                    <strong>Check #</strong>
                </td>
                <td bgcolor="#CCCCCC">
                    <strong>Account</strong>
                </td>
            </tr>
            <%
				'sSQL = "select g.LineNumber, cc.Description Coverage, g.ProcedureCode, p.Description Place, g.ChargeAmount,"
				'sSQL = sSQL & " g.TrackCode, g.RevenueCode, g.CheckNumber, g.CheckAccount, g.FromDate, g.ThruDate"
				'sSQL = sSQL & " FROM Charges g left join POS p ON g.PlaceOfService="
				'sSQL = sSQL & " p.PlaceOfService left join CoverageCodes cc ON"
				'sSQL = sSQL & " (g.CovCode=cc.CovCode and cc.Fund='" & m_sFund & "') WHERE"
				'sSQL = sSQL & " (g.LogDate = '" & m_dLogDate & "' AND"
				'sSQL = sSQL & " g.Sequence=" & m_iSequence & " AND g.Distr=" & m_iDistr			
				'sSQL = sSQL & " ) order by g.LineNumber"
        sSQL = "PB_LookupCharges '" & m_dLogDate & "', " & m_iSequence & ", " & m_iDistr & ", '" & m_sFund & "'"
        'response.write (sSQL)					
				Set adoRS = adoConn.execute(sSQL)
        'adoRS.Open sSQL, adoConn, adOpenForwardOnly, adLockOptimistic
        if not adoRS.EOF then
					Do While Not adoRS.EOF
						if isnull(adoRS("ChargeAmount")) then
							sTemp = ""
						else
							sTemp = formatcurrency(adoRS("ChargeAmount"))
						end if
            %>
            <tr bordercolordark="#FCFCFC" bordercolorlight="#CCCCCC">
                <td bgcolor="#F0F0F0">
                    &nbsp;<%= adoRS("LineNumber")%>
                </td>
                <td bgcolor="#F0F0F0">
                    &nbsp;<%= adoRS("Coverage")%>
                </td>
                <td bgcolor="#F0F0F0">
                    &nbsp;<%= adoRS("ProcedureCode")%>
                </td>
                <td bgcolor="#F0F0F0">
                    &nbsp;<%= adoRS("Place")%>
                </td>
                <td bgcolor="#F0F0F0">
                    &nbsp;<%= adoRS("FromDate")%>
                </td>
                <td bgcolor="#F0F0F0">
                    &nbsp;<%= adoRS("ThruDate")%>
                </td>
                <td bgcolor="#F0F0F0">
                    &nbsp;<%= sTemp%>
                </td>
                <td bgcolor="#F0F0F0">
                    &nbsp;<%= adoRS("TrackCode")%>
                </td>
                <td bgcolor="#F0F0F0">
                    &nbsp;<%= adoRS("RevenueCode")%>
                </td>
                <td bgcolor="#F0F0F0">
                    &nbsp;<%= adoRS("CheckNumber")%>
                </td>
                <td bgcolor="#F0F0F0">
                    &nbsp;<%= adoRS("CheckAccount")%>
                </td>
            </tr>
            <% 
						adoRS.MoveNext
					Loop
				else
            %>
            <tr bordercolordark="#FCFCFC" bordercolorlight="#CCCCCC">
                <td colspan="11" bgcolor="#F0F0F0">
                    <font color="red"><b>No Charges found.</b></font>
                </td>
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
        %>
</body>
</html>
<%
	end if
end if
%>