<%
' ========================================================================================
' Include file Group functions
' ========================================================================================
'This function returns a recordset of divisions the input user group(s) are in
'sGroupIDs: a list of comma delimited user group IDs
function GetUserGroupDivisionsRs(byval sGroupIDs, byref oRs)
    dim oErr, oSvr
    set oSvr = Server.CreateObject("JF_S_Permission.ISTestGroup")
    set oErr = oSvr.GetDivisions(Application("REPOSITORY"), sGroupIDs, oRs)
    set oSvr = nothing
    set GetUserGroupDivisionsRs = oErr
end function

' ========================================================================================
'This function returns a comma delimited list of division IDs the input 
'user group(s) are in
'sGroupIDs: a list of comma delimited user group IDs
function GetUserGroupDivisionsList(byval sGroupIDs, byref sDivisionIDs)
    dim oRs, oErr
    sDivisionIDs = ""
    set oErr = GetUserGroupDivisionsRs(sGroupIDs, oRs)
    if (oErr is nothing) then
        if (not oRs is nothing) then
            while (not oRs.EOF)
                if (sDivisionIDs <> "") then
                    sDivisionIDs = sDivisionIDs & ","
                end if
                sDivisionIDs = sDivisionIDs & oRs.Fields("DivisionID").Value
                oRs.MoveNext
            wend
        end if
    end if
    set GetUserGroupDivisionsList = oErr
end function

' ========================================================================================
'This function returns a comma delimited list of division IDs the input 
'user group(s) are in
'NOTE: SAME as above but unique DivisionIDs
'sGroupIDs: a list of comma delimited user group IDs
function GetUserGroupDivisionsList2(byval sGroupIDs, byref sDivisionIDs)
    dim oRs, oErr
    sDivisionIDs = ""
    set oErr = GetUserGroupDivisionsRs(sGroupIDs, oRs)
    if (oErr is nothing) then
        if (not oRs is nothing) then
            while (not oRs.EOF)
				if instr(sDivisionIDs, cstr(oRs.Fields("DivisionID").Value)) = 0 then
					'Not already there so add it
					if (sDivisionIDs <> "") then
						sDivisionIDs = sDivisionIDs & ","
					end if
					sDivisionIDs = sDivisionIDs & oRs.Fields("DivisionID").Value
				end if
                oRs.MoveNext
            wend
        end if
    end if
    set GetUserGroupDivisionsList2 = oErr
end function

%>