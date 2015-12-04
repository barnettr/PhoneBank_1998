<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Untitled</title>
</head>

<body>
<!--- <td ALIGN="CENTER" NOWRAP bgcolor="#cccccc">
								<font COLOR="BLUE"><b><%=adoRS(i).Name %></b></font>
							</td> --->

<% 
    if adoRS("locktype") <> "" then
      if Manager = "True" then
        bContinueProcessing = true
      elseif Supervisor = "True" then
        bContinueProcessing = true
      elseif Auditor = "True" then
        bContinueProcessing = true
      else
    	  bContinueProcessing = false
			  adoRS.Close
			  adoConn.Close
			  set adoRS = nothing
			  set adoConn = nothing
        response.redirect "PersonLockDetail.asp?Lock=1&SSN=" & m_iSSN & "&DepNo=0"
      end if
    end if 
    
    adoRS("SupervisorAccessOnly") = "True" or
%>
</body>
</html>
