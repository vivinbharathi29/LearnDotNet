<!--NO LONGER USING REMOTE SCRIPTING--> 
<!--#INCLUDE FILE="_SCRIPTLIBRARY/RS.ASP"--> 
<% RSDispatch %> 
<!-- #include file="../library/includes/ErrHandler.inc" -->
<!-- #include file="../library/includes/general.inc" -->
<!-- #include file="../includes/AMO.inc" -->
<script language="JavaScript" runat="server"> 

var public_description = new MainMethod(); 

function MainMethod() { 
//Make sure that the call for setFieldValue does an ESCAPE on the NewValue parameter
this.setFieldValue = Function('sRepository', 'ModuleID', 'Field', 'NewValue', 'strPerson', 'return SaveFieldValue(sRepository, ModuleID, Field, unescape(NewValue), strPerson)');
this.setPlatform = Function('sRepository', 'ModuleID', 'PlatformID', 'SetStatus', 'NewValue', 'strPerson', 'return SavePlatformStatus(sRepository, ModuleID, PlatformID, SetStatus, NewValue, strPerson)');
this.setComparability = Function('sRepository', 'ModuleID', 'DivisionID', 'SetStatus', 'NewValue', 'strPerson', 'return SaveComparabilityStatus(sRepository, ModuleID, DivisionID, SetStatus, NewValue, strPerson)');
this.setRegion = Function('sRepository', 'ModuleID', 'RegionID', 'SetStatus', 'strPerson', 'return SaveRegionStatus(sRepository, ModuleID, RegionID, SetStatus, strPerson)');
this.setGEODate = Function('sRepository', 'ModuleID', 'GEOID', 'NewValue', 'strPerson', 'return SaveGEODate(sRepository, ModuleID, GEOID, NewValue, strPerson)');
this.setBulkStatusValue = Function('sRepository', 'intStatus', 'strModuleIDs', 'strRegionIDs', 'strPerson', 'return updateBulkStatus(sRepository, intStatus, strModuleIDs, strPerson)');
this.setDateFieldValue = Function('sRepository', 'ModuleID', 'RegionID', 'Field', 'NewValue', 'strPerson', 'return SaveDateFieldValue(sRepository, ModuleID, RegionID, Field, unescape(NewValue), strPerson)');
this.setBulkDateValue = Function('sRepository', 'strModuleRegionIDs', 'strCPLBlindDate', 'strBOMRevADate', 'strRasDisconDate', 'strObsoleteDate', 'strGlobalseriesdate', 'strPerson', 'return updateBulkDate(sRepository, strModuleRegionIDs, strCPLBlindDate, strBOMRevADate, strRasDisconDate, strObsoleteDate, strGlobalseriesdate, strPerson)');
this.setBulkStatusEmailValue = Function('strModuleIDs', 'strRegionIDs', 'intStatusID', 'return SendBulkStatusChangeEmail(strModuleIDs, strRegionIDs, intStatusID)');
} 
</script> 

<script language="VBScript" runat="server"> 
'save a single field's value
function SaveFieldValue(sRepository, ModuleID, Field, NewValue, strPerson)
	dim oSvr, oErr
   
	SaveFieldValue = ""
	On Error Resume Next
	set oSvr = Server.CreateObject("JF_S_AMO.ISAMO")
    
	if Err.Number <> 0 then
		SaveFieldValue = Err.Description 
	else
	
		'fixed problem with passing in + characters
		NewValue = replace(NewValue,"!%%%!","+")
		
		set oErr = oSvr.SaveFieldValue(cstr(sRepository), ModuleID, Field, cstr(NewValue), strPerson)
		if Err.number <> 0 then 
			SaveFieldValue = Err.Description 
		end if 
		if not oErr is nothing then
		  SaveFieldValue = Errors_GetHTML(oErr)
		end if
	end if 
	set oErr = nothing
	set oSvr = nothing
end function

function SaveDateFieldValue(sRepository, ModuleID, RegionID, Field, NewValue, strPerson)
	dim oSvr, oErr

	SaveDateFieldValue = ""
	On Error Resume Next
	set oSvr = Server.CreateObject("JF_S_AMO.ISAMO")
	if Err.Number <> 0 then
		SaveDateFieldValue = Err.Description 
	else
	
		'fixed problem with passing in + characters
		NewValue = replace(NewValue,"!%%%!","+")
		
		set oErr = oSvr.SaveDateFieldValue(cstr(sRepository), ModuleID, RegionID, Field, cstr(NewValue), strPerson)
		if Err.number <> 0 then 
			SaveDateFieldValue = Err.Description 
		end if 
		if not oErr is nothing then
		  SaveDateFieldValue = Errors_GetHTML(oErr)
		end if
	end if 
	set oErr = nothing
	set oSvr = nothing
end function

'save a single field's value
function updateBulkStatus(sRepository, intStatus, strModuleIDs, strPerson)
	dim oSvr, oErr

	updateBulkStatus = ""
	On Error Resume Next
	set oSvr = Server.CreateObject("JF_S_AMO.ISAMO")
	if Err.Number <> 0 then
		updateBulkStatus = Err.Description 
	else
	
		set oErr =  oSvr.UpdateBulkStatus(cstr(sRepository), int(intStatus), cstr(strModuleIDs), strPerson)
		if Err.number <> 0 then 
			updateBulkStatus = Err.Description 
		end if 
		if not oErr is nothing then
			updateBulkStatus = Errors_GetHTML(oErr)
		end if
	end if 
	set oErr = nothing
	set oSvr = nothing
end function

function SendBulkStatusChangeEmail(strModuleIDs, strRegionIDs, intStatusID)
	call SendStatusChangeEmail (True,False,intStatusID,strModuleIDs,strRegionIDs)
end function

function updateBulkDate(sRepository, strModuleRegionIDs, strCPLBlindDate, strBOMRevADate, strRasDisconDate, strObsoleteDate, strGlobalseriesdate, strPerson)
	dim oSvr, oErr

	updateBulkDate = ""
	On Error Resume Next
	set oSvr = Server.CreateObject("JF_S_AMO.ISAMO")
	if Err.Number <> 0 then
		updateBulkDate = Err.Description 
	else
	
		set oErr = oSvr.UpdateBulkDate(sRepository, strModuleRegionIDs, strCPLBlindDate, strBOMRevADate, strRasDisconDate, strObsoleteDate, strGlobalseriesdate, strPerson)
		if Err.number <> 0 then 
			updateBulkDate = Err.Description 
		end if 
		if not oErr is nothing then
		  updateBulkDate = Errors_GetHTML(oErr)
		end if
	end if 
	set oErr = nothing
	set oSvr = nothing

end function

'save the Platform checkmark
function SavePlatformStatus(sRepository, ModuleID, PlatformID, SetStatus, NewValue, strPerson)
	dim oSvr, oErr

	SavePlatformStatus = ""
	On Error Resume Next
	set oSvr = Server.CreateObject("JF_S_AMO.ISAMO")
	if Err.Number <> 0 then
		SavePlatformStatus = Err.Description 
	else
		set oErr = oSvr.SavePlatformStatus(cstr(sRepository), ModuleID, PlatformID, SetStatus, cstr(NewValue), strPerson)
		if Err.number <> 0 then 
			SavePlatformStatus = Err.Description 
		end if 
		if not oErr is nothing then
		  SavePlatformStatus = Errors_GetHTML(oErr)
		end if
	end if 
	set oErr = nothing
	set oSvr = nothing
end function


'save the Comparability checkmark
function SaveComparabilityStatus(sRepository, ModuleID, DivisionID, SetStatus, NewValue, strPerson)
	dim oSvr, oErr

	SaveComparabilityStatus = ""
	On Error Resume Next
	set oSvr = Server.CreateObject("JF_S_AMO.ISAMO")
	if Err.Number <> 0 then
		SaveComparabilityStatus = Err.Description 
	else
		set oErr = oSvr.SaveComparabilityStatus(cstr(sRepository), ModuleID, DivisionID, SetStatus, strPerson)
		if Err.number <> 0 then 
			SaveComparabilityStatus = Err.Description 
		end if 
		if not oErr is nothing then
		  SaveComparabilityStatus = Errors_GetHTML(oErr)
		end if
	end if 
	set oErr = nothing
	set oSvr = nothing
end function

'save the Region checkmark
function SaveRegionStatus(sRepository, ModuleID, RegionID, SetStatus, strPerson)
	dim oSvr, oErr

	SaveRegionStatus = ""
	On Error Resume Next
	set oSvr = Server.CreateObject("JF_S_AMO.ISAMO")
	if Err.Number <> 0 then
		SaveRegionStatus = Err.Description 
	else
		set oErr = oSvr.SaveRegionStatus(cstr(sRepository), ModuleID, RegionID, SetStatus, strPerson)
		if Err.number <> 0 then 
			SaveRegionStatus = Err.Description 
		end if 
		if not oErr is nothing then
		  SaveRegionStatus = Errors_GetHTML(oErr)
		end if
	end if 
	set oErr = nothing
	set oSvr = nothing
end function

'save GEO Date
function SaveGEODate(sRepository, ModuleID, GEOID, NewValue, strPerson)
	dim oSvr, oErr

	SaveGEODate = ""
	On Error Resume Next
	set oSvr = Server.CreateObject("JF_S_AMO.ISAMO")
	if Err.Number <> 0 then
		SaveGEODate = Err.Description 
	else
		set oErr = oSvr.SaveGEODate(cstr(sRepository), ModuleID, GEOID, cstr(NewValue), strPerson)
		if Err.number <> 0 then 
			SaveGEODate = Err.Description 
		end if 
		if not oErr is nothing then
		  SaveGEODate = Errors_GetHTML(oErr)
		end if
	end if 
	set oErr = nothing
	set oSvr = nothing
end function

</script> 
