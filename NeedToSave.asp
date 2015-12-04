<% 
'****************************************************************************************
'* Immediately below are the fields that were removed form the conditional statement on the
'* ClaimSearch.asp page. If any of these fields are added back into the search function the
'* corresponding statement needs to be added at the bottom of the page in the Select Case.
'*
'* Below the fields is the origional conditional statement Just in case.
'*
'* Rob Barnett
'* September 20, 1999
'****************************************************************************************
 %>

"<%= m_sLastName%>" <> Criteria.LName.Value or _
"<%= m_sFund%>" <> Criteria.Fund.Value or _
"<%= m_sPlan%>" <> Criteria.Plan.Value or _
"<%= m_iBenPeriod%>" <> Criteria.BenPeriod.Value or _
"<%= m_dBenFrom%>" <> Criteria.BenFrom.Value or _
"<%= m_dBenThru%>" <> Criteria.BenThru.Value or _
"<%= m_dAdmitFrom%>" <> Criteria.AdmitFrom.Value or _
"<%= m_dAdmitThru%>" <> Criteria.AdmitThru.Value or _
"<%= m_dDisFrom%>" <> Criteria.DisFrom.Value or _
"<%= m_dDisThru%>" <> Criteria.DisThru.Value or _
"<%= m_sAdmitType%>" <> Criteria.AdmitType.Value or _
"<%= m_iBillType%>" <> Criteria.BillType.Value or _
"<%= m_dBillFrom%>" <> Criteria.BillFrom.Value or _
"<%= m_dBillThru%>" <> Criteria.BillThru.Value or _
"<%= m_iConfine%>" <> Criteria.Confine.Value or _
"<%= m_iICD9_Proc%>" <> Criteria.ICD9_Proc.Value or _
"<%= m_iDRG%>" <> Criteria.DRG.Value or _
"<%= m_iRevCode%>" <> Criteria.RevCode.Value or _
"<%= m_sModifier%>" <> Criteria.Modifier.Value or _
"<%= m_sCovCode%>" <> Criteria.CovCode.Value or _
"<%= m_iPOSCode%>" <> Criteria.POSCode.Value or _
"<%= m_sPayCode%>" <> Criteria.PayCode.Value or _
"<%= m_sPPO%>" <> Criteria.PPO.Value or _
"<%= m_bPPOCheck%>" <> Criteria.PPOOverride.Checked or _




if  "<%= m_dFromDate%>" <> Criteria.FromDate.Value or _
						"<%= m_dThruDate%>" <> Criteria.ThruDate.Value or _
						"<%= m_iSequence%>" <> Criteria.Sequence.Value or _
						"<%= m_iDistr%>" <> Criteria.Distr.Value or _
						"<%= m_iSSN%>" <> Criteria.SSN.Value or _
            "<%= m_sLastName%>" <> Criteria.LName.Value or _
						"<%= m_iDepNum%>" <> Criteria.DepNum.Value or _
						"<%= m_iTaxID%>" <> Criteria.TaxID.Value or _
						"<%= m_sProvName%>" <> Criteria.ProvName.Value or _
						"<%= m_iAssocNum%>" <> Criteria.AssocNum.Value or _
						"<%= m_sFund%>" <> Criteria.Fund.Value or _
						"<%= m_sStatus1%>" <> Criteria.Status1.Value or _
						"<%= m_sStatus2%>" <> Criteria.Status2.Value or _
						"<%= m_sStatus3%>" <> Criteria.Status3.Value or _
						"<%= m_dStatFrom%>" <> Criteria.StatFrom.Value or _
						"<%= m_dStatThru%>" <> Criteria.StatThru.Value or _
						"<%= m_sPlan%>" <> Criteria.Plan.Value or _
						"<%= m_sRSN%>" <> Criteria.RSN.Value or _
						"<%= m_dServiceFrom%>" <> Criteria.ServiceFrom.Value or _
						"<%= m_dServiceThru%>" <> Criteria.ServiceThru.Value or _
						"<%= m_sHWType%>" <> Criteria.HWType.Value or _
						"<%= m_iBenPeriod%>" <> Criteria.BenPeriod.Value or _
						"<%= m_dBenFrom%>" <> Criteria.BenFrom.Value or _
						"<%= m_dBenThru%>" <> Criteria.BenThru.Value or _
						"<%= m_dAdmitFrom%>" <> Criteria.AdmitFrom.Value or _
						"<%= m_dAdmitThru%>" <> Criteria.AdmitThru.Value or _
						"<%= m_dDisFrom%>" <> Criteria.DisFrom.Value or _
						"<%= m_dDisThru%>" <> Criteria.DisThru.Value or _
						"<%= m_sAdmitType%>" <> Criteria.AdmitType.Value or _
						"<%= m_iBillType%>" <> Criteria.BillType.Value or _
						"<%= m_dBillFrom%>" <> Criteria.BillFrom.Value or _
						"<%= m_dBillThru%>" <> Criteria.BillThru.Value or _
						"<%= m_iConfine%>" <> Criteria.Confine.Value or _
						"<%= m_iICD9_Proc%>" <> Criteria.ICD9_Proc.Value or _
						"<%= m_iDRG%>" <> Criteria.DRG.Value or _
						"<%= m_iRevCode%>" <> Criteria.RevCode.Value or _
						"<%= m_sDiagCode%>" <> Criteria.DiagCode.Value or _
						"<%= m_sProcCode%>" <> Criteria.ProcCode.Value or _
						"<%= m_sModifier%>" <> Criteria.Modifier.Value or _
						"<%= m_sTrackCode%>" <> Criteria.TrackCode.Value or _
						"<%= m_sCovCode%>" <> Criteria.CovCode.Value or _
						"<%= m_sPreCalCode%>" <> Criteria.PreCalCode.Value or _
						"<%= m_iPOSCode%>" <> Criteria.POSCode.Value or _
						"<%= m_sPayCode%>" <> Criteria.PayCode.Value or _
						"<%= m_cCharge%>" <> Criteria.Charge.Value or _
						"<%= m_sCheckAcct%>" <> Criteria.CheckAcct.Value or _
						"<%= m_iCheckNum%>" <> Criteria.CheckNum.Value or _
						"<%= m_sPPO%>" <> Criteria.PPO.Value or _
						"<%= m_bPPOCheck%>" <> Criteria.PPOOverride.Checked or _
						"<%= m_bCOBCheck%>" <> Criteria.COB.Checked or _
						"<%= m_bOverPayCheck%>" <> Criteria.OverPay.Checked or _
						"<%= m_bOtherAdjustCheck%>" <> Criteria.NonOverPay.Checked or _
						"<%= m_bCOBOverCheck%>" <> Criteria.COBOver.Checked then

            
            
            
From Form Letters and Letters: 

"<%= m_sStatus%>" <> Criteria.Status.Value or _
"<%= m_sStatus%>" <> Criteria.Status.Value or _