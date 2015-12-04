<%@  language="VBScript" %>
<% Option Explicit %>
<% Response.Expires = 0 %>
<!--#include file="user.inc" -->
<!--#include file="adovbs_mod.inc" -->
<%
Dim i, adoConn, adoRS, sSQL, sTemp
dim m_iCOBKey, m_sParticName, m_bUseEAS, m_sBirthDate, m_sFName, m_sLName, DepNumber
dim adoCmd, adoParam1, m_iSSN
dim locktype, Admin

m_iCOBKey = Request.QueryString("COBkey")
m_sParticName = Request.QueryString("ParticName")
m_sBirthDate = request.querystring("BirthDay")
m_sFName = request.querystring("dFName")
m_sLName = request.querystring("dLName")
DepNumber = request.querystring("DepNumber")
m_iSSN = request.querystring("SSN")

if Request.Querystring("UseEAS") = "" then
	m_bUseEAS = false
else
	m_bUseEAS = Request.Querystring("UseEAS")
end if

if m_iCOBKey = "" or m_sParticName = "" then
	Response.Write "<ul><font color='red' face='verdana, arial, helvetica' size='2'>COBKey or Participant Name missing; contact your network administrator</font></ul>"
else
  
    set adoConn = Server.CreateObject("ADODB.Connection")
    adoConn.ConnectionTimeout = 0
	adoConn.CommandTimeout = 0
    set adoCmd = Server.CreateObject("ADODB.Command")
    adoCmd.CommandTimeout = 0
	set adoRS = Server.CreateObject("ADODB.Recordset")
	adoConn.Open Application("DataConn")
  
    sSQL = "select locktype, SupervisorAccessOnly from Names where SSN=" & m_iSSN
    adoRS.Open sSQL, adoConn, adOpenKeyset, adLockOptimistic
    if not adoRS.EOF then
        locktype = adoRS("locktype")
        Admin = adoRS("SupervisorAccessOnly")
    end if
    adoRS.Close
  
    adoCmd.CommandText = "PB_ListCOB"
    adoCmd.CommandType = adCmdStoredProc
    Set adoCmd.ActiveConnection = adoConn
    Set adoParam1 = adoCmd.CreateParameter("@P1_CobKey", adInteger, adParamInput)
    adoCmd.Parameters.Append adoParam1
    adoCmd("@P1_CobKey") = m_iCOBKey
%>
<html>
<head>
    <!--#include file="VBFuncs.inc" -->
</head>
<body language="VBScript" onload="UpdateStatus()">
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
                <font size="+2" face="verdana, arial, helvetica"><strong>COB Details</strong></font>
            </td>
            <td align="CENTER" width="20%">
                <img src="images/bluebar2.gif" onclick="history.go(-1)">&nbsp;&nbsp;<img src="images/bluebar3.gif"
                    onclick="history.go(+1)" border="0">
            </td>
        </tr>
    </table>
    <%
	adoRS.Open adoCmd
	if adoRS.EOF then
		Response.Write ("<p><ul><font color='red' face='verdana, arial, helvetica' size='2'><b>Unable to find all necessary information; please contact your network administrator.</b></font></ul>")
	else
    %>
    <p>
        <center>
            <font face="verdana, arial, helvetica" size="2"><b>Participant</b></font></center>
        <table width="100%" border="2" bgcolor="white" bordercolor="white" style="font: 9pt verdana, arial, helvetica,sans-serif">
            <tr bordercolordark="#F0F0F0" bordercolorlight="#999999">
                <td width="33%" bgcolor="#cccccc">
                    <strong>Participant Name:</strong>
                    <%
						sTemp = "<A HREF='PersonDetails.asp?SSN=" & adoRS("SSN") & "&DepNum=0&locktype=" & locktype & "&Admin=" & Admin
						sTemp = sTemp & "' onClick='javascript:return LogCheck(" & adoRS("SSN") & ")'>"
						sTemp = sTemp &"<b>" & m_sParticName & "</b></A>"
						Response.Write sTemp
                    %>
                </td>
                <td width="33%" bgcolor="#cccccc">
                    <strong>Participant SSN:</strong> <font color="<%= Session("EmphColor")%>"><b>
                        <%= adoRS("SSN")%></b></font>
                </td>
                <td width="33%" bgcolor="#cccccc">
                    <strong>Participant Birthday:</strong> <font color="<%= Session("EmphColor")%>"><b>
                        <%= m_sBirthDate %></b></font>
                </td>
            </tr>
        </table>
        <p>
            &nbsp;
            <p>
                <center>
                    <font face="verdana, arial, helvetica" size="2"><b>Dependent</b></font></center>
                <table width="100%" align="center" border="1" cellpadding="1" cellspacing="1" bgcolor="white"
                    bordercolor="white" style="font: 9pt verdana, arial, helvetica,sans-serif">
                    <tr bordercolordark="#F0F0F0" bordercolorlight="#999999">
                        <td bgcolor="#cccccc">
                            <strong>Dependent Name:</strong> <font color="<%= Session("EmphColor")%>"><b>
                                <%= m_sFName %>&nbsp;<%= m_sLName %></b></font>
                        </td>
                        <td bgcolor="#cccccc">
                            <strong>Dependent Number:</strong> <font color="<%= Session("EmphColor")%>"><b>
                                <%= DepNumber %></b></font>
                        </td>
                        <td bgcolor="#cccccc">
                            <strong>Employer:</strong> <font color="<%= Session("EmphColor")%>"><b>
                                <%= adoRS("EmployerName")%></b></font>
                        </td>
                    </tr>
                    <%
			if adoRS("SameAddress") then
                    %>
                    <tr bordercolordark="#FCFCFC" bordercolorlight="#CCCCCC">
                        <td bgcolor="#f0f0f0" colspan="3">
                            <strong>Address:</strong>&nbsp;<font color="<%= Session("EmphColor")%>"><b>Same as Participant</b></font>
                        </td>
                    </tr>
                    <%
			else
                    %>
                    <tr bordercolordark="#FCFCFC" bordercolorlight="#CCCCCC">
                        <td bgcolor="#f0f0f0" colspan="3">
                            <strong>Address:</strong> <font color="<%= Session("EmphColor")%>"><b>
                                <%= adoRS("ResName")%></b></font>
                        </td>
                    </tr>
                    <tr bordercolordark="#f0f0f0" bordercolorlight="#999999">
                        <td bgcolor="#cccccc" colspan="3">
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font
                                color="<%= Session("EmphColor")%>"><b><%= adoRS("ResStreet")%></b></font>
                        </td>
                    </tr>
                    <tr bordercolordark="#FCFCFC" bordercolorlight="#CCCCCC">
                        <td bgcolor="#f0f0f0" colspan="3">
                            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font
                                color="<%= Session("EmphColor")%>"><b><%= adoRS("ResCity") & ", " & adoRS("ResState") & " " & adoRS("ResPostalCode")%></b></font>
                        </td>
                    </tr>
                    <%
			end if
                    %>
                </table>
                <p>
                    &nbsp;
                    <p>
                        <center>
                            <font face="verdana, arial, helvetica" size="2"><b>OIC Participant</b></font></center>
                        <table width="100%" align="CENTER" border="1" cellpadding="1" cellspacing="1" bgcolor="white"
                            bordercolor="white" style="font: 9pt verdana, arial, helvetica,sans-serif">
                            <tr bordercolordark="#F0F0F0" bordercolorlight="#999999">
                                <td bgcolor="#cccccc">
                                    <strong>Other Insured Name:</strong> <font color="<%= Session("EmphColor")%>"><b>
                                        <%= adoRS("OtherName")%></b></font>
                                </td>
                                <td bgcolor="#cccccc">
                                    <strong>SSN:</strong> <font color="<%= Session("EmphColor")%>"><b>
                                        <%= adoRS("OtherSSN")%></b></font>
                                </td>
                                <td bgcolor="#cccccc">
                                    <strong>Birthdate:</strong> <font color="<%= Session("EmphColor")%>"><b>
                                        <%= adoRS("BirthDate")%></b></font>
                                </td>
                            </tr>
                            <tr bordercolordark="#FCFCFC" bordercolorlight="#CCCCCC">
                                <td bgcolor="#f0f0f0">
                                    <strong>Relationship To Dependent:</strong> <font color="<%= Session("EmphColor")%>">
                                        <b>
                                            <%= adoRS("DescOfRelation")%></b></font>
                                </td>
                                <td bgcolor="#F0F0F0">
                                    <strong>Carrier:</strong> <font color="<%= Session("EmphColor")%>"><b>
                                        <%= adoRS("Carrier")%></b></font>
                                </td>
                                <td bgcolor="#F0F0F0">
                                    <strong>Carrier Code:</strong> <font color="<%= Session("EmphColor")%>"><b>
                                        <%= adoRS("CarrierCode") %></b></font>
                                </td>
                            </tr>
                            <tr bordercolordark="#F0F0F0" bordercolorlight="#999999">
                                <td bgcolor="#cccccc">
                                    <strong>Group or Individual:</strong> <font color="<%= Session("EmphColor")%>"><b>
                                        <%= adoRS("GroupOrIndividual")%></b></font>&nbsp;&nbsp;&nbsp;&nbsp;<strong>Group Number:</strong>
                                    <font color="<%= Session("EmphColor")%>"><b>
                                        <%= adoRS("GroupNumber") %></b></font>&nbsp;&nbsp;&nbsp;&nbsp;<strong>Employer:</strong>
                                    <font color="<%= Session("EmphColor")%>"><b>
                                        <%= adoRS("EmployerName") %></b></font>
                                </td>
                                <td bgcolor="#cccccc">
                                    <strong>Effective Date:</strong> <font color="<%= Session("EmphColor")%>"><b>
                                        <%= adoRS("EffectiveDate") %></b></font>
                                </td>
                                <td bgcolor="#cccccc">
                                    <strong>Termination Date:</strong> <font color="<%= Session("EmphColor")%>"><b>
                                        <%= adoRS("TermDate") %></b></font>
                                </td>
                            </tr>
                        </table>
                        <p>
                            &nbsp;
                            <p>
                                <center>
                                    <font face="verdana, arial, helvetica" size="2"><b>Coverage Types</b></font></center>
                                <table width="100%" align="CENTER" border="1" cellpadding="1" cellspacing="1" bgcolor="white"
                                    bordercolor="white" style="font: 9pt verdana, arial, helvetica,sans-serif">
                                    <tr bordercolordark="#F0F0F0" bordercolorlight="#999999">
                                        <td bgcolor="#cccccc">
                                            <strong>Medical:</strong> <font color="<%= Session("EmphColor")%>">
                                                <% if adoRS("Coveragemedical") = "True" then %><b>Yes</b><% else %><b>No</b><% end if %></font>
                                        </td>
                                        <td bgcolor="#cccccc">
                                            <strong>Dental:</strong> <font color="<%= Session("EmphColor")%>">
                                                <% if adoRS("Coveragedental") = "True" then %><b>Yes</b><% else %><b>No</b><% end if %></font>
                                        </td>
                                        <td bgcolor="#cccccc">
                                            <strong>Orthodontic:</strong> <font color="<%= Session("EmphColor")%>">
                                                <% if adoRS("CoverageOrtho") = "True" then %><b>Yes</b><% else %><b>No</b><% end if %></font>
                                        </td>
                                        <td bgcolor="#cccccc">
                                            <strong>Vision:</strong> <font color="<%= Session("EmphColor")%>">
                                                <% if adoRS("CoverageVision") = "True" then %><b>Yes</b><% else %><b>No</b><% end if %></font>
                                        </td>
                                        <td bgcolor="#cccccc">
                                            <strong>Prescription:</strong> <font color="<%= Session("EmphColor")%>">
                                                <% if adoRS("CoveragePrescription") = "True" then %><b>Yes</b><% else %><b>No</b><% end if %></font>
                                        </td>
                                        <td bgcolor="#cccccc">
                                            <strong>Chiropractic:</strong> <font color="<%= Session("EmphColor")%>">
                                                <% if adoRS("Coveragechiro") = "True" then %><b>Yes</b><% else %><b>No</b><% end if %></font>
                                        </td>
                                        <td bgcolor="#cccccc">
                                            <strong>Other:</strong> <font color="<%= Session("EmphColor")%>">
                                                <% if adoRS("CoverageOther") = "True" then %><b>Yes</b><% else %><b>No</b><% end if %></font>
                                        </td>
                                    </tr>
                                </table>
                                <% if adoRS("NWAOrder") <> "1" then %>
                                <p>
                                    &nbsp;
                                    <p>
                                        <center>
                                            <font face="verdana, arial, helvetica" size="2"><b>OIC Becomes Primary Due To The Following
                                                Rule</b></font></center>
                                        <table width="100%" align="CENTER" border="1" cellpadding="1" cellspacing="1" bgcolor="white"
                                            bordercolor="white" style="font: 9pt verdana, arial, helvetica,sans-serif">
                                            <tr bordercolordark="#F0F0F0" bordercolorlight="#999999">
                                                <% if adoRS("Primary_Birth") = "True" then %>
                                                <td bgcolor="#cccccc">
                                                    <strong>Birthday:</strong> <font color="<%= Session("EmphColor")%>"><b>Yes</b></font>
                                                </td>
                                                <% end if %>
                                                <% if adoRS("Custody") = "True" then %>
                                                <td bgcolor="#cccccc">
                                                    <strong>Custodial:</strong> <font color="<%= Session("EmphColor")%>"><b>Yes</b></font>
                                                </td>
                                                <% end if %>
                                                <% if adoRS("FinancialResp") = "True" then %>
                                                <td bgcolor="#cccccc">
                                                    <strong>Financial:</strong> <font color="<%= Session("EmphColor")%>"><b>Yes</b></font>
                                                </td>
                                                <% end if %>
                                                <% if adoRS("DivDecree") = "True" then %>
                                                <td bgcolor="#cccccc">
                                                    <strong>Divorce Decree:</strong> <font color="<%= Session("EmphColor")%>"><b>Yes</b></font>
                                                </td>
                                                <% end if %>
                                                <% if adoRS("Primary_Employment") = "True" then %>
                                                <td bgcolor="#cccccc">
                                                    <strong>Employment:</strong> <font color="<%= Session("EmphColor")%>"><b>Yes</b></font>
                                                </td>
                                                <% end if %>
                                                <% if adoRS("Primary_Other") = "True" then %>
                                                <td bgcolor="#cccccc">
                                                    <strong>Other:</strong> <font color="<%= Session("EmphColor")%>"><b>Yes</b></font>
                                                </td>
                                                <% end if %>
                                            </tr>
                                        </table>
                                        <% end if %>
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