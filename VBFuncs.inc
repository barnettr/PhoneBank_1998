<SCRIPT LANGUAGE="VBSCRIPT">

	sub TwoFunctions()
    UpdateScreen(0)
    StartForm()
  end sub
  
  sub TwoFunctions2()
    UpdateScreen(3)
    StartForm()
  end sub
  
  sub TwoFunctions3()
    UpdateScreen(1)
    StartForm2()
  end sub
  
  sub UpdateScreen(iIndex)
		Parent.NavFrame.ChangeCritTabVisual(iIndex)
		UpdateStatus()	
	end sub
	
	sub UpdateStatus()
		top.AppStatus.location.replace("Status.asp")
	end sub
	
	sub WorkingStatus()
		top.AppStatus.location.replace("Working.htm")
	end sub

	function LogCheck(iSSN)
		if <%= Session("IsClerk")%> and (top.NavFrame.iCurrSSN <> iSSN) then
			LogCheck = PDetailsDialog() 
		else
			WorkingStatus()
			LogCheck = true 
		end if
	end function
		
	function PDetailsDialog()
	dim sResult

		sResult=showModalDialog("LogDialog.asp", "LogDialog", "dialogWidth:500px; dialogHeight:295px; help:no;")
		select case sResult
			case "normal", "override"
				WorkingStatus
				PDetailsDialog = true
				window.event.srcElement.href = "WriteLog.asp" & mid(window.event.srcElement.href,instr(window.event.srcElement.href,"?"), len(window.event.srcElement.href)) & "&LogType=" & sResult
			case "cancel", ""
				PDetailsDialog = false
		end select
	end function
	
	function LogCall()
	dim sResult

		sResult=showModalDialog("NonClerkLogDialog.asp", "NonClerkLogDialog", "dialogWidth:500px; dialogHeight:295px; help:no;")
		if not(sResult = "cancel" or  sResult = "") then
			WorkingStatus
			location.href = "WriteLog.asp?SSN=" & sResult & "&LogType=normal&OpenPhoneSearch=true"
		end if
	end function
	
	function ContainsInvalids(sInput)
	dim Invalids(2), i
	dim bFound
	
		bFound = false
		Invalids(0) = "&"
		Invalids(1) = chr(34) 'double quote
		Invalids(2) = "'"
		if ubound(Invalids) = 0 then
			sInvals = "character (" & Invalids(0)
		else 
			sInvals = "characters (" & Invalids(0)
			for i = 1 to ubound(Invalids)
				sInvals = sInvals & " or " & Invalids(i)
			next
		end if
		sInvals = sInvals & ")"
		for i = 0 to ubound(Invalids)
			if instr(sInput,Invalids(i)) > 0 then
				msgbox "Use of reserved " & sInvals & " in criteria.",,"Reserved Characters"
				ContainsInvalids = true
				exit function
			end if
		next
		
		ContainsInvalids = false
	
	end function

	function FixQuote(sInput)
		FixQuote = replace(sInput, "'", "''")
	end function
			
</SCRIPT>
