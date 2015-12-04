<%@  language="VBSCRIPT" %>
<% Option Explicit %>
<% Response.Expires = 0 %>
<% Response.Buffer = true %>
<!--#include file="user.inc" -->
<!--#include file="adovbs_mod.inc" -->
<%
Dim adoConn, adoRS, sSQL, sTemp, sCriteria
Dim sPageData, iPageNo, iRowCount, i, iPageCount
Dim m_iSSN, m_sLastName, m_sAddress1, m_sCity, m_sState, m_sZip, m_sFirstName
dim m_bUseDepSSN, m_bUseEAS
dim bUseSQL, bContinueProcessing, Manager, Supervisor, Auditor, SupervisorAccessOnly, locktype, OtherArea
dim DepNum
dim User, Admin

set adoConn = Server.CreateObject("ADODB.Connection")
adoConn.CommandTimeout = 0
set adoRS = Server.CreateObject("ADODB.Recordset")
adoConn.Open Application("DataConn")
sSQL = "select * from UserInformation where LogonID='" & Session("User") & "'"
adoRS.Open sSQL, adoConn, adOpenKeyset, adLockOptimistic

if not adoRS.EOF then
  Manager = adoRS("IsManager")
  Supervisor = adoRS("IsSupervisor")
  Auditor = adoRS("IsAuditor")
else
  response.write "<p><font color='red' size='2' face='verdana, arial, helvetica'><b>You have not been added to the UserInformation table as a verified user of PhoneBank.</b></font>"
  adoRS.Close
end if

bContinueProcessing = true

	function PSearchFailed()
	dim sRedir
		sRedir = "PSearchFailed.asp?"
		sRedir = sRedir & "SSN=" & m_iSSN
		sRedir = sRedir & "&Address1=" & m_sAddress1
		sRedir = sRedir & "&City=" & m_sCity
		sRedir = sRedir & "&LastName=" & m_sLastName
    sRedir = sRedir & "&FirstName=" & m_sFirstName
		sRedir = sRedir & "&State=" & m_sState
		sRedir = sRedir & "&ZIP=" & m_sZip
		sRedir = sRedir & "&UseDepSSN=" & m_bUseDepSSN
		sRedir = sRedir & "&UseEAS=" & m_bUseEAS
    sRedir = sRedir & "&OtherArea=" & OtherArea
  	Response.Redirect(sRedir)
		Response.End   
  end function
  

if Request.Querystring("UserCriteria") = "Initial" then
	bUseSQL = false
else
	bUseSQL = true
	if Request.Querystring("UseEAS") = "" then
		m_bUseEAS = false
	else
		m_bUseEAS = Request.Querystring("UseEAS")
	end if
  if Request.Querystring("OtherArea") = "" then
		OtherArea = false
	else
		OtherArea = Request.Querystring("OtherArea")
	end if
	if not m_bUseEAS then
		sSQL = "select rtrim(n.LastName) + ', ' + rtrim(n.FirstName) + ' ' + rtrim(n.MiddleName) 'Name'," 
		sSQL = sSQL & " n.ssn 'SSN', n.depNumber 'Dep. #', n.depSSN 'Dep. SSN', n.SupervisorAccessOnly, n.locktype, n.LockedDate, "
		sSQL = sSQL & " h.StreetAddress 'Street Address', "
	else
		sSQL = "select rtrim(n.LastName) + ', ' + rtrim(n.FirstName) + ' ' + rtrim(n.MiddleName) 'Name [EAS Data]'," 
		sSQL = sSQL & " n.ssn 'SSN', n.depNumber 'Dep. #', n.depSSN 'Dep. SSN',"
		sSQL = sSQL & " h.Addr1 'Street Address', "
  end if
  if OtherArea = "True" then
    sSQL = "select rtrim(n.LastName) + ', ' + rtrim(n.FirstName) + ' ' + rtrim(n.MiddleName) 'Name [AAO Data]'," 
	sSQL = sSQL & " n.ssn 'SSN', n.depNumber 'Dep. #', n.depSSN 'Dep. SSN',"
	sSQL = sSQL & " h.StreetAddress 'Street Address', "
  end if
	sSQL = sSQL & " h.City 'City', h.State 'State', h.PostalCode 'ZIP Code' from"
	if not m_bUseEAS then
    if OtherArea = "True" then
      sSQL = sSQL & " OtherAreaNames n left join OtherAreaHomes h on n.ssn=h.ssn where"
    else
		  sSQL = sSQL & " names n left join homes h on n.ssn=h.ssn where"
    end if
	else 
		sSQL = sSQL & " EASnames n left join eashomes h on n.ssn=h.ssn where"
  end if
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
	if Request.Querystring("UseDepSSN") = "" then
		m_bUseDepSSN = false
	else
		m_bUseDepSSN = Request.Querystring("UseDepSSN")
	end if
	if Request.Querystring("SSN") <> "" then
    if len(Request.Querystring("SSN")) = 7 then
      m_iSSN = "00" & Request.Querystring("SSN")
    else
		  m_iSSN = Request.Querystring("SSN")
    end if
		if m_bUseDepSSN then
			sSQL = sSQL & " (n.ssn=" & m_iSSN & " or n.depssn=" & m_iSSN & ")"
		else 
			sSQL = sSQL & " n.ssn=" & m_iSSN 
		end if
		sSQL = sSQL & " AND"
	end if
	if Request.Querystring("LastName") <> "" then
		m_sLastName = Request.Querystring("LastName")
		sSQL = sSQL & " n.lastname like '" & m_sLastName & "%' AND" 
	end if
  if Request.Querystring("FirstName") <> "" then
		m_sFirstName = Request.Querystring("FirstName")
		sSQL = sSQL & " n.firstname like '" & m_sFirstName & "%' AND" 
	end if
	if Request.Querystring("Address1") <> "" then
		m_sAddress1 = Request.Querystring("Address1")
		sSQL = sSQL & " h.StreetAddress like '" & m_sAddress1 & "%' AND" 
	end if
	if Request.Querystring("City") <> "" then
		m_sCity = Request.Querystring("City")
		sSQL = sSQL & " h.City like '" & m_sCity & "%' AND" 
	end if
	if Request.Querystring("State") <> "" then
		m_sState = Request.Querystring("State")
		sSQL = sSQL & " h.State like '" & m_sState & "%' AND" 
	end if
	if Request.Querystring("ZIP") <> "" then
		m_sZip = Request.Querystring("ZIP")
		sSQL = sSQL & " h.PostalCode like '" & m_sZip & "%' AND"
	end if
	
	if not m_bUseEAS then
    if OtherArea = "True" then
      sSQL = left(sSQL,len(sSQL)-4)
		  sSQL = sSQL & " order by 'Name [AAO Data]', n.ssn"
    else
		  sSQL = sSQL & " h.DateCreated=(select max(datecreated) from homes where ssn=n.ssn)"
		  sSQL = sSQL & " order by Name, n.ssn"
    end if
	else
		sSQL = left(sSQL,len(sSQL)-4)
		sSQL = sSQL & " order by 'Name [EAS Data]', n.ssn"
	end if
end if
'response.write sSQL
'response.write "<br>" & 

%>
<html>
<head>
    <title>Details</title>
    <script language="javascript" src="function.js"></script>
</head>
<body topmargin="2" leftmargin="2" rightmargin="0" language="VBScript" onload="TwoFunctions()">
    <link rel="STYLESHEET" href="styles/CritTable.css">
    <table width="100%" cols="3" border="0">
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
                <font size="+2" face="verdana, arial, helvetica"><strong>Participant/Dependent Search</strong></font>
            </td>
            <td align="CENTER" width="20%">
                <img src="images/bluebar2.gif" onclick="history.go(-1)" border="0">&nbsp;&nbsp;<img
                    src="images/bluebar3.gif" onclick="history.go(+1)" border="0">
            </td>
        </tr>
    </table>
    <%
  'response.write "Supervisor = " & Supervisor & "<br>"
  'response.write "Manager = " & Manager & "<br>"
  'response.write "Auditor = " & Auditor & "<br>"
  if bUseSQL then
		set adoConn = Server.CreateObject("ADODB.Connection")
		adoConn.CommandTimeout = 0
		set adoRS = Server.CreateObject("ADODB.Recordset")
		adoConn.Open Application("DataConn")
    adoRS.MaxRecords = 200
    adoRS.Open sSQL, adoConn, adOpenKeyset, adLockOptimistic
    
    if adoRS.EOF then
        bContinueProcessing = false
        adoRS.Close
        adoConn.Close
        set adoRS = nothing
        set adoConn = nothing
        PSearchFailed()
    end if
        
    if not m_bUseEAS then
      if OtherArea = "False" then
        locktype = adoRS("locktype")
        Admin = adoRS("SupervisorAccessOnly")
        DepNum = adoRS("Dep. #")
        sTemp = "0"
        if adoRS("locktype") <> "" then 'or adoRS("SupervisorAccessOnly") then
          if Manager = "True" or Supervisor = "True" or Auditor = "True" then
            bContinueProcessing = true
          else
            if m_sLastName <> "" then
              bContinueProcessing = true
            else
      	      bContinueProcessing = false
  			      adoRS.Close
  			      adoConn.Close
  			      set adoRS = nothing
  			      set adoConn = nothing
              response.redirect "PDDirect.asp?SSN=" & m_iSSN & "&LastName=" & m_sLastName & "&locktype=" & locktype & "&DepNum=" & sTemp & "&Admin=" & Admin
            end if
          end if
        end if
        if adoRS("SupervisorAccessOnly") then
          if Manager = "True" or Supervisor = "True" or Auditor = "True" then
            bContinueProcessing = true
          else
            if m_sLastName <> "" then
              bContinueProcessing = true
            else
      	      bContinueProcessing = false
  			      adoRS.Close
  			      adoConn.Close
  			      set adoRS = nothing
  			      set adoConn = nothing
              response.redirect "LockedFile.asp?SSN=" & m_iSSN & "&LastName=" & m_sLastName & "&locktype=" & locktype & "&DepNum=" & sTemp & "&Admin=" & Admin
            end if
          end if
        end if
      end if
    end if
    
    
    
    'if adoRS.EOF then
			'bContinueProcessing = false
			'adoRS.Close
			'adoConn.Close
			'set adoRS = nothing
			'set adoConn = nothing
			'PSearchFailed()
		'else
			if m_iSSN <> "" then
'				if not m_bUseDepSSN then
					sTemp = "0"
'				else
'					sTemp =  adoRS("Dep. #")
'				end if
				if (not Session("IsClerk")) or (trim(Session("CurrSSN")) = trim(adoRS("SSN"))) then
       
					Response.Redirect "PersonDetails.asp?SSN=" & adoRS("SSN") & "&DepNum=" & sTemp & "&UseEAS=" & m_bUseEAS & "&locktype=" & locktype & "&OtherArea=" & OtherArea & "&Admin=" & Admin
				else
          if not m_bUseEAS then
					  Response.Redirect "PDDirect.asp?SSN=" & adoRS("SSN") & "&DepNum=" & sTemp & "&UseEAS=" & m_bUseEAS & "&locktype=" & locktype & "&OtherArea=" & OtherArea & "&Admin=" & Admin
          else
            Response.Redirect "PDDirect.asp?SSN=" & adoRS("SSN") & "&DepNum=" & sTemp & "&UseEAS=" & m_bUseEAS & "OtherArea=True"
          end if
				end if
				Response.End
			end if
		end if
	'end if
	if bContinueProcessing then
		if bUseSQL then
      Admin = adoRS("SupervisorAccessOnly")
			adoRS.PageSize = m_iPageSize ' Number of rows per page
			iPageCount = adoRS.PageCount
			adoRS.AbsolutePage = iPageNo
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
    <div style="height=13%; width: 100%; overflow=auto;">
        <form id="Criteria" method="GET" action="PersonSearch.asp" target="Details">
        <table class="CriteriaTable" width="100%" cellpadding="0" cellspacing="0" border="0"
            cols="3">
            <tr>
                <!-- These TH's were required in order to get IE to deal properly with the DIV...among other things, without them and with BORDER=1, things were displayed OK, butwith the desired BORDER=0 IE seemed to 'lose' a number of the elements-->
                <th>
                </th>
                <th>
                </th>
                <th>
                </th>
            </tr>
            <tr>
                <td width="33%" class="White">
                    Social Security Number:&nbsp;&nbsp;<input type="TEXT" id="SSN" name="SSN" size="10"
                        maxlength="10" value="<%= m_iSSN%>">
                </td>
                <td width="33%" class="White">
                    Last Name:&nbsp;&nbsp;<input type="TEXT" id="LastName" name="LastName" size="20"
                        maxlength="20" value="<%= m_sLastName%>">
                </td>
                <td width="34%" class="White">
                    First Name:&nbsp;&nbsp;<input type="TEXT" id="FirstName" name="FirstName" size="20"
                        maxlength="20" value="<%= m_sFirstName%>">
                </td>
            </tr>
            <tr>
                <td colspan="2" class="White">
                    <input type="CHECKBOX" id="IncludeDepSSN" name="UseDepSSN" value="TRUE" <%
							if bUseSQL then
								if m_bUseDepSSN then
									Response.Write " CHECKED "
								end if
							end if
%>>Check here to also search for Dependent SSN (slows search)
                </td>
                <td>
                    &nbsp;
                </td>
                <td>
                    &nbsp;
                </td>
            </tr>
            <tr>
                <td class="White">
                    Address:&nbsp;&nbsp;<input type="TEXT" id="Address1" name="Address1" size="20" maxlength="30"
                        value="<%= m_sAddress1%>">
                </td>
                <td class="White">
                    City:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input
                        type="TEXT" id="City" name="City" size="20" maxlength="28" value="<%= m_sCity%>">
                </td>
                <td class="White">
                    State:&nbsp;&nbsp;<input type="TEXT" id="State" name="State" size="3" maxlength="2"
                        value="<%= m_sState%>">&nbsp;&nbsp;&nbsp;&nbsp;Zip Code:&nbsp;&nbsp;<input type="TEXT"
                            id="ZIP" name="ZIP" size="10" maxlength="12" value="<%= m_sZip%>">
                </td>
            </tr>
            <tr>
                <td colspan="3">
                    &nbsp;
                </td>
            </tr>
        </table>
        <input type="HIDDEN" id="ActionType" name="ActionType" value>
        </form>
    </div>
    <!--- <br CLEAR="LEFT"> --->
    <%
		if bUseSQL then
    
    
    %>
    <div style="height=68%; width=100%; overflow=auto">
        <table width="100%" border="1" bordercolor="white" bgcolor="white" style="font: 9pt verdana, arial, helvetica,sans-serif">
            <tr bordercolordark="#F0F0F0" bordercolorlight="#999999">
                <td align="center" nowrap bgcolor="#cccccc">
                    <font color="blue"><b>Name</b></font>
                </td>
                <td align="center" nowrap bgcolor="#cccccc">
                    <font color="blue"><b>SSN</b></font>
                </td>
                <td align="center" nowrap bgcolor="#cccccc">
                    <font color="blue"><b>Dep #</b></font>
                </td>
                <td align="center" nowrap bgcolor="#cccccc">
                    <font color="blue"><b>Dep SSN</b></font>
                </td>
                <td align="center" nowrap bgcolor="#cccccc">
                    <font color="blue"><b>Street Address</b></font>
                </td>
                <td align="center" nowrap bgcolor="#cccccc">
                    <font color="blue"><b>City</b></font>
                </td>
                <td align="center" nowrap bgcolor="#cccccc">
                    <font color="blue"><b>State</b></font>
                </td>
                <td align="center" nowrap bgcolor="#cccccc">
                    <font color="blue"><b>Zip Code</b></font>
                </td>
            </tr>
            <% 
					iRowCount = adoRS.PageSize
					Do While Not adoRS.EOF and iRowCount > 0 
            %>
            <tr bordercolordark="#FCFCFC" bordercolorlight="#CCCCCC">
                <td nowrap bgcolor="F0F0F0">
                    <a href="PersonDetails.asp?SSN=<%= adoRS("SSN") %>&DepNum=<%= adoRS("Dep. #") %>&UseEAS=<%= m_bUseEAS %>&locktype=<%= adoRS("locktype") %>&Admin=<%= adoRS("SupervisorAccessOnly") %>"
                        onclick='LogCheck(<%= adoRS("SSN")%>)'>
                        <%= adoRS("Name") %></a>
                </td>
                <td nowrap bgcolor="F0F0F0">
                    <% if len(adoRS("SSN")) = 7 then %>00<%= adoRS("SSN") %><% else %><%= adoRS("SSN") %><% end if %>
                </td>
                <td nowrap bgcolor="F0F0F0">
                    <%= adoRS("Dep. #") %>
                </td>
                <td nowrap bgcolor="F0F0F0">
                    <%= adoRS("Dep. SSN") %>
                </td>
                <td nowrap bgcolor="F0F0F0">
                    <%= adoRS("Street Address") %>
                </td>
                <td nowrap bgcolor="F0F0F0">
                    <%= adoRS("City") %>
                </td>
                <td nowrap bgcolor="F0F0F0">
                    <%= adoRS("State") %>
                </td>
                <td nowrap bgcolor="F0F0F0">
                    <%= adoRS("ZIP Code") %>
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
    %>
</body>
<!--#include file="VBFuncs.inc" -->
<script language="VBSCRIPT">

	sub ValidSend(iIndex)
	dim bValidData, iTemp

		Criteria.SSN.Value = trim(Criteria.SSN.Value)
		Criteria.LastName.Value = trim(Criteria.LastName.Value)
        Criteria.FirstName.Value = trim(Criteria.FirstName.Value)
		Criteria.Address1.Value = trim(Criteria.Address1.Value)
		Criteria.City.Value = trim(Criteria.City.Value)
		Criteria.State.Value = trim(Criteria.State.Value)
		Criteria.Zip.Value = trim(Criteria.Zip.Value)
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
		if Criteria.IncludeDepSSN.Checked and Criteria.SSN.Value = "" then
			msgbox "To search by Dependent SSN, you have to enter a valid SSN.",,"Invalid Data"
			exit sub
		end if
		if Criteria.LastName.Value <> "" then
			if ContainsInvalids(Criteria.LastName.Value) then
				msgbox "Please remove invalid characters from the Last Name field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
    if Criteria.FirstName.Value <> "" then
			if ContainsInvalids(Criteria.FirstName.Value) then
				msgbox "Please remove invalid characters from the Last Name field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.Address1.Value <> "" then
			if ContainsInvalids(Criteria.Address1.Value) then
				msgbox "Please remove invalid characters from the Address field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.City.Value <> "" then
			if ContainsInvalids(Criteria.City.Value) then
				msgbox "Please remove invalid characters from the City field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.State.Value <> "" then
			if ContainsInvalids(Criteria.State.Value) then
				msgbox "Please remove invalid characters from the State field.",,"Invalid Data"
				exit sub
			end if
			bValidData = true
		end if
		if Criteria.Zip.Value <> "" then
			iTemp = replace(Criteria.Zip.Value,"-","")
			if not isNumeric(iTemp) then
				msgbox "Please enter a valid Zip Code.",,"Invalid Data"
				exit sub
			else
				if iTemp > <%= Application("IntMax")%> or _
						iTemp < <%= Application("IntMin")%> then
					msgbox "Please enter a valid Zip Code.",,"Invalid Data"
					exit sub
				end if
			end if
			bValidData = true
		end if
		if not bValidData then
			msgbox "You haven't entered anything to search for.",48,"No Criteria"
			exit sub
		end if
		
		WorkingStatus
		if "<%= m_iSSN%>" <> Criteria.SSN.Value or _
				"<%= m_sLastName%>" <> Criteria.LastName.Value or _
        "<%= m_sFirstName%>" <> Criteria.FirstName.Value or _
				"<%= m_sAddress1%>" <> Criteria.Address1.Value or _
				"<%= m_sCity%>" <> Criteria.City.Value or _
				"<%= m_sState%>" <> Criteria.State.Value or _
				"<%= m_sZip%>" <> Criteria.Zip.Value then
			Criteria.ActionType.value=""
		else
			Criteria.ActionType.value=iIndex
		end if
		Criteria.submit
	end sub
		
</script>
</html>
