<%@ LANGUAGE="VBSCRIPT" %>
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
dim bUseSQL, bContinueProcessing

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
	if not m_bUseEAS then
		sSQL = "select rtrim(n.LastName) + ', ' + rtrim(n.FirstName) + ' ' + rtrim(n.MiddleName) 'Name'," 
		sSQL = sSQL & " n.ssn 'SSN', n.depNumber 'Dep. #', n.depSSN 'Dep. SSN',"
		sSQL = sSQL & " h.StreetAddress 'Street Address', "
	else
		sSQL = "select rtrim(n.LastName) + ', ' + rtrim(n.FirstName) + ' ' + rtrim(n.MiddleName) 'Name [EAS Data]'," 
		sSQL = sSQL & " n.ssn 'SSN', n.depNumber 'Dep. #', n.depSSN 'Dep. SSN',"
		sSQL = sSQL & " h.Addr1 'Street Address', "
	end if
	sSQL = sSQL & " h.City 'City', h.State 'State', h.PostalCode 'ZIP Code' from"
	if not m_bUseEAS then
		sSQL = sSQL & " names n left join homes h on n.ssn=h.ssn where"
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
		m_iSSN = Request.Querystring("SSN")
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
		sSQL = sSQL & " h.DateCreated=(select max(datecreated) from homes where ssn=n.ssn)"
		sSQL = sSQL & " order by Name, n.ssn"
	else
		sSQL = left(sSQL,len(sSQL)-4)
		sSQL = sSQL & " order by 'Name [EAS Data]', n.ssn"
	end if
end if


	if bUseSQL then
		set adoConn = Server.CreateObject("ADODB.Connection")
		'************************************************************************************
		' Setting CommandTimeout to 0 waits indefinitely for response; note that per 
		' ADO documentation Command object will not inherit the Connection setting (default
		' is 30 seconds).  This appeared to help with response issues on seasql02.
		'************************************************************************************
		adoConn.CommandTimeout = 0
		set adoRS = Server.CreateObject("ADODB.Recordset")
		adoConn.Open Application("DataConn")
		adoRS.Open sSql, adoConn, adOpenKeyset, adLockOptimistic
			
		if adoRS.EOF then
			bContinueProcessing = false
			adoRS.Close
			adoConn.Close
			set adoRS = nothing
			set adoConn = nothing
			PSearchFailed()
		else
			if m_iSSN <> "" then
'				if not m_bUseDepSSN then
					sTemp = "0"
'				else
'					sTemp =  adoRS("Dep. #")
'				end if
				if (not Session("IsClerk")) or (trim(Session("CurrSSN")) = trim(adoRS("SSN"))) then
					Response.Redirect "PersonDetails.asp?SSN=" & adoRS("SSN") & "&DepNo=" & sTemp & "&UseEAS=" & m_bUseEAS
				else
					Response.Redirect "PDDirect.asp?SSN=" & adoRS("SSN") & "&DepNo=" & sTemp & "&UseEAS=" & m_bUseEAS
				end if
				Response.End
			end if
		end if
	end if
	if bContinueProcessing then
		if bUseSQL then
			adoRS.PageSize = m_iPageSize ' Number of rows per page
			iPageCount = adoRS.PageCount
			adoRS.AbsolutePage = iPageNo
		end if
%>
<html>
<body>

<%
		if bUseSQL then
%>
			<div STYLE="height=56%; width=100%; overflow=auto">
				<table WIDTH="100%" BORDER="1">
					<tr>
<%
						For i = 0 to adoRS.Fields.Count - 1 
%>
							<td ALIGN="CENTER" NOWRAP>
								<font COLOR="BLUE"><b><%=adoRS(i).Name %></b></font>
							</td>
<%
						Next
%>
					</tr>
<% 
					iRowCount = adoRS.PageSize
					Do While Not adoRS.EOF and iRowCount > 0 
%>
						<tr>
<%
							sTemp = "<TD NOWRAP><A HREF='PersonDetails.asp?SSN=" & adoRS("SSN") & "&DepNo=" & adoRS("Dep. #") & "&UseEAS=" & m_bUseEAS
							sTemp = sTemp & "' onClick='javascript:return LogCheck(" & adoRS("SSN") & ")'>"
							sTemp = sTemp & adoRS(0) & "</A></TD>"
							Response.Write sTemp
								
							For i = 1 to adoRS.Fields.Count - 1 
								if adoRS(i).type = adCurrency then
									Response.Write "<TD ALIGN=RIGHT NOWRAP>" & formatcurrency(adoRS(i)) & "</TD>"
								else 
									Response.Write "<TD NOWRAP>" & adoRS(i) & "</TD>"
								end if
							Next
%>
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
</html>