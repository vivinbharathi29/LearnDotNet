﻿<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #include file="../Fusion/lib_debug.inc" -->
<!-- #include file="../../includes/DataWrapper.asp" -->
<HTML>
<HEAD>
<script src="/excalibur/includes/client/jquery.min.js" type="text/javascript"></script>
<script src="/excalibur/includes/client/json2.js"></script>
<script src="/Pulsar/Scripts/spin/spin.js"></script>
<script src="/Pulsar/Scripts/spin/jquery.spin.js"></script>
<SCRIPT>
$(document).ready(function() {
	
		$("#CompUpdLink").click(function(e) {
		GetComponentUpdateList();
	});
});


 function GetComponentUpdateList(){
	
	var strComp = "<%=request("ChangeDetails")%>"
	if (!strComp || strComp =="" )
	{
		alert("You must select at least one component.");
		return;
	}

	$("#CompUpdDiv").spin("medium", "#0096D6");	
    var CompList = strComp.split(",");
	var CompDetail ;
	var item;
	var Comp = [];
	for (i = 0; i < CompList.length; i++) {
			
		CompDetail = CompList[i].split("|");
		
		if (document.getElementById("ErrSpan" + "|" + CompDetail[2] + "|" + CompDetail[3] + "|")){
			continue;
		}
		
		if (CompDetail[4].indexOf("rep") >= 0)
			CompDetail[4] = "replace"
		else if (CompDetail[4].indexOf("rem") >= 0)
			CompDetail[4] = "remove"
			
		item = {
			ProductName: CompDetail[0].replace(/^[ ]+|[ ]+$/g,""),
			ProductDropName: CompDetail[1],
			Action: CompDetail[4],
			NewCompPartNo: CompDetail[5].replace(/^[ ]+|[ ]+$/g,"").toUpperCase(),
			OldCompPartNo: CompDetail[11].replace(/^[ ]+|[ ]+$/g,"").toUpperCase() 
		}
		Comp.push(item);		
    } 

    var jsonObj = { "comp":Comp, "isIRS": 0 }
	
	 
	 if (Comp.length > 0){
		 $.ajax({
            type: "POST",
            contentType: "application/json; charset=utf-8",
            url: "/Pulsar/Common/ExportComponentListToExcel",
            data: JSON.stringify(jsonObj),
            dataType: "json",
            async: true,
			beforeSend: function(){
				$("#CompUpdDiv").children().addClass("disabledLink");
				$("#CompUpdLink").unbind("click");

			},
			complete: function () {
				$("#CompUpdDiv").spin(false);
				$("#CompUpdDiv").children().removeClass("disabledLink");
				$("#CompUpdLink").bind("click", function(){
					GetComponentUpdateList();
				});
				
			},
			success: function (data) {
				if (data.ReturnValue ==1){
					alert(data.Message);
					return;
				}
				else{
					window.location ="/Pulsar/Common/DownloadFile?fileName=" + data.FileName;
		
				}
            },
            error: function (msg, status, error) {
				$("#CompUpdDiv").spin(false);
                var errMsg = $.parseJSON(msg.responseText);
                alert("Please try again later: " + errMsg);
            }   
        });
	 } else{
        $("#CompUpdDiv").spin(false);
		alert("component update list is empty");
	 }
	 
	
 }
</SCRIPT>
<style>
	body {
		 font-family: Verdana;
		 font-size: xx-small;
		 color:black;
	}

      .disabledLink {
        color: #333;
        text-decoration: none;
        cursor: default;
    }
</style>
</HEAD>
<BODY>
<%
'printrequest
'response.end
	Dim cn, cm, cn2, dw, dw2, cmd2
	dim strSuccess, strDeliveryTypeIDList, strRetMsg, strSpaces, strUserGroupIDs, strProductDropReplace, strRequest, strSORequest, strDestRequest, strLocRequest 
	dim strSpanDONE, strSpanERROR, strXML
	dim p_strProduct, p_strProductDropName, p_strProductDropID, p_strDeliverableName, p_strAction, p_strNewPartNumber, p_strNewIRSComponentPassID, p_strPartNumber
	dim p_strSRP, p_strNotInRCD, p_strInAppRecovery, p_strDIB, p_strOldPartNumber, p_InstallOptionRequest, p_SortOrderOverride, p_DestOverride, p_LocaleOverride
	dim intError
	dim ProcessArray, ParameterArray, arrRS
	dim i, p, x
	dim rs, rs2
	dim strChangeDetails
	dim CurrentDomain, Currentuser, CurrentUserID, CurrentUserEmail, CurrentUserName
	dim objErr
    dim hasPermission

	strSuccess = "1"
	strSpaces = " &nbsp; "
	strSpanDONE = "<span style='color:blue;font-weight:bold;'>"
	strSpanERROR = "<span style='color:red;font-weight:bold;'>"

	strChangeDetails = request("ChangeDetails")
	
'strChangeDetails="Astro 2.0|13WWASAA3##|8663|HP PageLift|add|P008N1-B2M|302279|1|1|0|0|P008N1-B2M,Astro 2.0|13WWASAA3##|8663|HP PageLift|rep|P008N1-B2N|304855|1|1|0|0|P008N1-B2M,Astro 2.0|13WWASAA3##|8663|HP PageLift|rem|P008N1-B2N|304855|1|1|0|0|P008N1-B2M"

	
	
	if trim(strChangeDetails) <> "" then
		'Get User
		CurrentUser = lcase(Session("LoggedInUser"))
		if instr(currentuser,"\") > 0 then
			CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
			Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
		end if

         set cn = server.CreateObject("ADODB.Connection")
		 cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	     cn.Open
	     set rs = server.CreateObject("ADODB.recordset")
	     set cm = server.CreateObject("ADODB.Command")
		 Set cm.ActiveConnection = cn
		 cm.CommandType = 4
		 cm.CommandText = "spGetUserInfo"
		
		 Set p = cm.CreateParameter("@UserName", 200, &H0001, 80)
		 p.Value = Currentuser
		 cm.Parameters.Append p
	
		 Set p = cm.CreateParameter("@Domain", 200, &H0001, 30)
		 p.Value = CurrentDomain
		 cm.Parameters.Append p
	
	     rs.CursorType = adOpenForwardOnly
	     rs.LockType=AdLockReadOnly
	     Set rs = cm.Execute 

		CurrentUserID = 0
		if not (rs.EOF and rs.BOF) then
			CurrentUserID = rs("ID")
            CurrentUserName = rs("Name")
		end if
		rs.Close
        cn.Close
        set cm = nothing
        set cn=nothing

        set dw2 = new DataWrapper
		set cn2 = dw2.CreateConnection("PDPIMS_ConnectionString")
    	'Get IRS UserGroupIDs that the User is part of
		Set cmd2 = dw2.CreateCommandSP(cn2, "usp_USR_ViewGroupsByUser")
		dw2.CreateParameter cmd2, "@p_intID", adInteger, adParamInput, 30, CurrentUserID
		Set rs = dw2.ExecuteCommandReturnRS(cmd2)

		if rs.eof and rs.bof then
			strSuccess = "0"
			response.write "Could not get User Groups</br>" & vbCrLf
			response.End
		end if
		
		strUserGroupIDs = ""
		Do while not rs.eof
			if strUserGroupIDs <> "" then strUserGroupIDs = strUserGroupIDs & ","
			strUserGroupIDs = strUserGroupIDs & rs("GroupID")
			rs.MoveNext
		loop
		rs.Close
		set dw2 = nothing
		set cn2 = nothing
		set cmd2 = nothing


		ProcessArray = split(strChangeDetails,",")
	
		'cn.BeginTrans
		for i = lbound(ProcessArray) to ubound(ProcessArray)
'			
			ParameterArray = split(ProcessArray(i),"|")
			p_strProduct = ParameterArray(0)								'Product
			p_strProductDropName = ParameterArray(1)						'ProductDrop name
			p_strProductDropID = ParameterArray(2)							'ProductDropID
			p_strDeliverableName = replace(ParameterArray(3),"^",",")		'Deliverable Name
			p_strAction = ParameterArray(4)									'action, i.e. add, rep, rem
			p_strNewPartNumber = ucase(ParameterArray(5))					'New Part Number
			p_strNewIRSComponentPassID = ParameterArray(6)					'IRS ComponentPassID for new Part Number
			p_strSRP = ParameterArray(7)									'SRP
			p_strNotInRCD = ParameterArray(8)								'NotInRCD
			p_strInAppRecovery = ParameterArray(9)							'InAppRecovery
			p_strDIB = ParameterArray(10)									'DIB
			p_strOldPartNumber = ucase(ParameterArray(11))					'Replace (Old) Part Number

            strRequest = "IO" & p_strProductDropName & "|" &p_strNewPartNumber
			If Not(isnull(request(strRequest))) and  request(strRequest) <> "" then 
				p_InstallOptionRequest = CInt(trim(request(strRequest)))
			else
				p_InstallOptionRequest = 1
			end if	

            strSORequest = "SO" & p_strProductDropName & "|" & p_strNewPartNumber 
			If Not(isnull(request(strSORequest))) and  request(strSORequest) <> "" then 
				p_SortOrderOverride = CInt(trim(request(strSORequest)))
			else
				p_SortOrderOverride = 0
			end if

			strDestRequest = "DC" & p_strProductDropName & "|" & p_strNewPartNumber 
            If Not(isnull(request(strDestRequest))) and  request(strDestRequest) <> "" then 
				p_DestOverride = CInt(trim(request(strDestRequest)))
			else
				p_DestOverride = 0
			end if
            
            strLocRequest = "IL" & p_strProductDropName & "|" & p_strNewPartNumber 
            If Not(isnull(request(strLocRequest ))) and  request(strLocRequest) <> "" then 
				p_LocaleOverride = CInt(trim(request(strLocRequest)))
			else
				p_LocaleOverride = 0
			end if

            Set dw2 = New DataWrapper
			Set cn2 = dw2.CreateConnection("PDPIMS_ConnectionString")
    		Set cmd2 = dw2.CreateCommandSP(cn2, "usp_ProductExplorer_CheckStateAndApprover")  
			dw2.CreateParameter cmd2, "@p_intPDID", adInteger, adParamInput, 30, cint(p_strProductDropID)
            dw2.CreateParameter cmd2, "@p_intUserID", adInteger, adParamInput, 30, CurrentUserID
            dw2.CreateParameter cmd2, "@p_hasPermission", adInteger, adParamOutput, 8, 0
    
			dw2.ExecuteNonQuery(cmd2)
			hasPermission = cmd2("@p_hasPermission")
            set dw2 = nothing
			set cn2 = nothing
			set cmd2 = nothing

            if (hasPermission = 1) Then

					'Call IRS stored procedures for whatever action is needed
					select case p_strAction
						case "add"
							response.write "<b>Add:</b> Part Number " & p_strNewPartNumber & " (" & p_strDeliverableName & ") to Product " & p_strProduct & ", Product Drop (" & p_strProductDropName & ", ID=" & p_strProductDropID & ").</br>"

							' Need Delivery Type List
							'	DIB (DeliveryTypeID = 16) = p_strDIB
							'	In App Rcvry (DeliveryTypeID = 8) = p_strInAppRecovery
							'	Not in RCD (DeliveryTypeID = 32) = p_strNotInRCD
							'	SRP (DeliveryTypeID = 1) = p_strSRP
							' Updater = Need 18 characters of Excalibur user. Using left 18 characters of the name.

							strDeliveryTypeIDList = ""
							if p_strSRP = "1" then	'SRP, DeliverableTypeID = 1
								strDeliveryTypeIDList = strDeliveryTypeIDList & "1"
							end if
							if p_strNotInRCD = "1" then	'Not in RCD, DeliverableTypeID = 32
								if strDeliveryTypeIDList <> "" then strDeliveryTypeIDList = strDeliveryTypeIDList & ","
								strDeliveryTypeIDList = strDeliveryTypeIDList & "32"
							end if
							if p_strInAppRecovery = "1" then	'In App Rcvry, DeliverableTypeID = 8
								if strDeliveryTypeIDList <> "" then strDeliveryTypeIDList = strDeliveryTypeIDList & ","
								strDeliveryTypeIDList = strDeliveryTypeIDList & "8"
							end if
							if p_strDIB = "1" then	'DIB, DeliverableTypeID = 16
								if strDeliveryTypeIDList <> "" then strDeliveryTypeIDList = strDeliveryTypeIDList & ","
								strDeliveryTypeIDList = strDeliveryTypeIDList & "16"
							end if

							Set dw2 = New DataWrapper
							Set cn2 = dw2.CreateConnection("PDPIMS_ConnectionString")
                       
							Set cmd2 = dw2.CreateCommandSP(cn2, "usp_ProductExplorer_AddComponentFromProduct")
							dw2.CreateParameter cmd2, "@p_intPDID", adInteger, adParamInput, 30, cint(p_strProductDropID)
							dw2.CreateParameter cmd2, "@p_chrPartNumber", adVarchar, adParamInput, len(p_strNewPartNumber), p_strNewPartNumber
							dw2.CreateParameter cmd2, "@p_chrDeliveryTypes", adVarchar, adParamInput, len(strDeliveryTypeIDList), strDeliveryTypeIDList
							dw2.CreateParameter cmd2, "@p_chrComponentPassID", adVarchar, adParamInput, 10, p_strNewIRSComponentPassID
							dw2.CreateParameter cmd2, "@p_chrUpdater", adVarchar, adParamInput, len(left(CurrentUserName,18)), left(CurrentUserName,18)
							dw2.CreateParameter cmd2, "@p_chrHTChangeReason", adVarchar, adParamInput, len("Added component by Excalibur"), "Added component by Excalibur" 
							dw2.CreateParameter cmd2, "@p_chrRetMsg", adVarchar, adParamOutput, 8000, 0
							dw2.ExecuteNonQuery(cmd2)
							strRetMsg = cmd2("@p_chrRetMsg")

							if Err.number = 0 and cn2.Errors.Count = 0 then
								if strRetMsg = "" then
									response.write strSpaces & strSpanDONE & "DONE</span></br>" & vbCrLf
                                elseif InStr(strRetMsg,"ERROR") > 0 then
									response.write  "<span id='ErrSpan|" & p_strProductDropID & "|" & p_strDeliverableName & "|'>"  & strRetMsg & "</span></span></br>" & vbCrLf	
								else
									response.write strRetMsg & "</span></br>" & vbCrLf	'spaces and blue font are done in stored procedure
								end if
							else
								Response.Write strSpaces & strSpanERROR & "<span id='ErrSpan|" & p_strProductDropID & "|" & p_strDeliverableName & "|'>" & "Error updating Product Explorer. "
								for each objErr in cn2.Errors
									response.write strSpaces & "Description: " & objErr.Description
								next
								response.write "</span></span></br>" & vbCrLf
								response.write strRetMsg
								'response.write " Error: " & cstr(cn2.Description) & "</br>"
							end if
							cn2.Close
							set cn2 = nothing
							set cmd2 = nothing

						case "rep"
							response.write "<b>Replace:</b> Old Part Number " & p_strOldPartNumber & " (" & p_strDeliverableName & ") with " & p_strNewPartNumber & " in Product " & p_strProduct & ", Product Drop (" & p_strProductDropName & ", ID=" & p_strProductDropID & ").</br>"

							Set dw2 = New DataWrapper
							Set cn2 = dw2.CreateConnection("PDPIMS_ConnectionString")

							Set cmd2 = dw2.CreateCommandSP(cn2, "usp_ProductExplorer_SWRefresh_ReplacePreview")
							dw2.CreateParameter cmd2, "@p_OldPart", adVarchar, adParamInput, len(p_strOldPartNumber), p_strOldPartNumber
							dw2.CreateParameter cmd2, "@p_NewPart", adVarchar, adParamInput, len(p_strNewPartNumber), p_strNewPartNumber
							dw2.CreateParameter cmd2, "@p_chrUserGroupIDs", adVarchar, adParamInput, len(strUserGroupIDs), strUserGroupIDs
							dw2.CreateParameter cmd2, "@p_intUserID", adInteger, adParamInput, 30, CurrentUserID
							dw2.CreateParameter cmd2, "@p_NewPartRevision", adVarchar, adParamInput, 5, ""
							dw2.CreateParameter cmd2, "@p_IsCD", adBoolean, adParamInput, 30, 0
							dw2.CreateParameter cmd2, "@p_ProductdropIDs", adVarChar, adParamInput, len(p_strProductDropID)+1, p_strProductDropID
							Set rs2 = dw2.ExecuteCommandReturnRS(cmd2)

							if Err.number <> 0 or cn2.Errors.Count <> 0 then
								Response.Write strSpaces & strSpanERROR & "<span id='ErrSpan|" & p_strProductDropID & "|" & p_strDeliverableName & "|'>"  & "Could not get Preview tables returned</span></span></br>" & vbCrLf
							else
								
								'Get the 2nd recordset
								set rs2 = rs2.NextRecordset()
								'printdebug(rs2)
								
								'Get the 3rd recordset
								set rs2 = rs2.NextRecordset()
								'printdebug(rs2)

								if rs2.Recordcount = 0 then
									Response.Write strSpaces & strSpanERROR & "<span id='ErrSpan|" & p_strProductDropID & "|" & p_strDeliverableName & "|'>"  & "The old part(s) selected do not exist in any product drop(s) where you have the approval to modify. Please verify you are the approver of the product(s) and resubmit your search and replace request. You can view approver(s) in Project Explorer\Product Drop Properties tab.</span></span></br>" & vbCrLf
								else
									rs2.Filter = "IsQualified=0"

									if rs2.Recordcount > 0 then
										set rs2 = rs2.NextRecordset()
									

										Response.Write strSpaces & "Could not replace in this product drop because:</br>" & vbCrLf
										Do while not rs2.eof
											response.write strSpaces & strSpaces & strSpanERROR & "<span id='ErrSpan|" & p_strProductDropID & "|" & p_strDeliverableName & "|'>"  & rs2("why") & "</span></span></br>" & vbCrLf
											rs2.MoveNext
										loop
									else
										'We can do the replace so replace the component now
										rs2.Close

										'Parameter @p_ProductdropIDs is in the format of "ProductDropID:New ComponentPassID:Message"
										strProductDropReplace = p_strProductDropID & ":" & p_strNewIRSComponentPassID & ":Replaced by Excalibur"

										Set cmd2 = dw2.CreateCommandSP(cn2, "usp_ProductExplorer_SearchReplace_Replace")
										dw2.CreateParameter cmd2, "@p_OldPart", adVarchar, adParamInput, len(p_strOldPartNumber), p_strOldPartNumber
										dw2.CreateParameter cmd2, "@p_NewPart", adVarchar, adParamInput, len(p_strNewPartNumber), p_strNewPartNumber
										dw2.CreateParameter cmd2, "@p_intUserID", adInteger, adParamInput, 30, CurrentUserID
										dw2.CreateParameter cmd2, "@p_ProductdropIDs", adVarChar, adParamInput, len(strProductDropReplace), strProductDropReplace
										dw2.CreateParameter cmd2, "@p_IsCD", adBoolean, adParamInput, 30, 0
										dw2.CreateParameter cmd2, "@p_InstallOptionRequest", adInteger, adParamInput, 30, p_InstallOptionRequest	'assuming that the part numbers are for the same deliverable
										dw2.CreateParameter cmd2, "@p_LockValue", adInteger, adParamInput, 30, 0
                                        dw2.CreateParameter cmd2, "@p_SortOrderOverride", adInteger, adParamInput, 30, p_SortOrderOverride	
                                        dw2.CreateParameter cmd2, "@p_DestOverride", adInteger, adParamInput, 30, p_DestOverride	
                                        dw2.CreateParameter cmd2, "@p_LocaleOverride", adInteger, adParamInput, 30, p_LocaleOverride	
										Set rs2 = dw2.ExecuteCommandReturnRS(cmd2)

										if Err.number <> 0 or cn2.Errors.Count <> 0 then
											Response.Write strSpaces & strSpanERROR & "<span id='ErrSpan|" & p_strProductDropID & "|" & p_strDeliverableName & "|'>"  & "No tables returned so could not Replace part number</span></span></br>" & vbCrLf
										else
											do until rs2 is nothing
												arrRS = rs2.GetRows()
												set rs2 = rs2.NextRecordset()
											loop


											'display the BOM and IsReplaced fields
											for x = 0 to ubound(arrRS,2) 'rows
												if lcase(arrRS(5,x)) = "yes" then	'ISReplaced field
													response.write strSpaces & "BOM (" & arrRS(2,x) & ") - " & strSpanDONE & "DONE</span></br>"
												else
													Response.Write strSpaces & "BOM (" & arrRS(2,x) & ") - " & strSpanERROR & "<span id='ErrSpan|" & p_strProductDropID & "|" & p_strDeliverableName & "|'>"  & "Error replacing part number. "
													for each objErr in cn2.Errors
														response.write " Description: " & objErr.Description
													next
													response.write "</span></span></br>" & vbCrLf
												end if
											next
											response.write "</br>" & vbCrLf

										end if
									end if
								end if

							end if
		
							set dw2 = nothing
							set cn2 = nothing
							set cmd2 = nothing
	
						case "rem"	'Remove a component
							response.write "<b>Remove:</b> Part Number " & p_strNewPartNumber & " (" & p_strDeliverableName & ") from Product " & p_strProduct & ", Product Drop (" & p_strProductDropName & ", ID=" & p_strProductDropID & ").</br>"

							Set dw2 = New DataWrapper
							Set cn2 = dw2.CreateConnection("PDPIMS_ConnectionString")

							Set cmd2 = dw2.CreateCommandSP(cn2, "usp_ProductExplorer_ProductDrop_RemoveComponents")
							dw2.CreateParameter cmd2, "@p_intPDID", adInteger, adParamInput, 30, cint(p_strProductDropID)
							strXML = "<?xml version='1.0' encoding='iso-8859-1' ?><components><component ComponentPassID='" & p_strNewIRSComponentPassID & "' /></components>"
							dw2.CreateParameter cmd2, "@p_xmlComponentPassID", adLongVarChar, adParamInput, len(strXML), strXML
							dw2.CreateParameter cmd2, "@p_chrHTChangeReason", adVarchar, adParamInput, len("Removed component by Excalibur"), "Removed component by Excalibur" 
							dw2.CreateParameter cmd2, "@p_chrUpdater", adVarchar, adParamInput, len(left(CurrentUserName,18)), left(CurrentUserName,18)
							dw2.CreateParameter cmd2, "@p_bitDeleteFromOtherDrop", adBoolean, adParamInput, 30, 0
							
							dw2.ExecuteNonQuery(cmd2)
							
							if Err.number = 0 and cn2.Errors.Count = 0 and cint(intError) = 0 then
								response.write strSpaces & strSpanDONE & "DONE</span></br>"
							else
								Response.Write strSpaces & strSpanERROR & "<span id='ErrSpan|" & p_strProductDropID & "|" & p_strDeliverableName & "|'>"  & "Error updating Product Explorer. "
								for each objErr in cn2.Errors
									response.write " Description: " & objErr.Description
								next
								response.write "</span></span></br>" & vbCrLf
	
							end if
							response.write "</br>" & vbCrLf
							cn2.Close
							set cn2 = nothing
							set cmd2 = nothing

					end select

                     ' update the RCDOnly column
                    if (p_strAction = "rep") then  
                        p_strPartNumber = p_strOldPartNumber& "," &  p_strNewPartNumber

                    else
                        p_strPartNumber = p_strNewPartNumber

                    end if
					Set dw2 = New DataWrapper			
					Set cn2 = dw2.CreateConnection("PDPIMS_ConnectionString")
					Set cmd2 = dw2.CreateCommandSP(cn2, "usp_ProductExplorer_UpdateRCDOnlyFromProduct")
					dw2.CreateParameter cmd2, "@ProductDotName", adVarchar, adParamInput, len(p_strProduct), p_strProduct
					dw2.CreateParameter cmd2, "@ProductDropID", adInteger, adParamInput, 30, cint(p_strProductDropID)
					dw2.CreateParameter cmd2, "@ComponentPartNumber", adVarchar, adParamInput, len(p_strPartNumber), p_strPartNumber
                    dw2.CreateParameter cmd2, "@UserID", adInteger, adParamInput, 30, CurrentUserID
					dw2.ExecuteNonQuery(cmd2)
					cn2.Close
					set cn2 = nothing
					set cmd2 = nothing

             else
                select case p_strAction
						case "add"
							response.write "<b>Add:</b> Part Number " & p_strNewPartNumber & " (" & p_strDeliverableName & ") to Product " & p_strProduct & ", Product Drop (" & p_strProductDropName & ", ID=" & p_strProductDropID & ").</br>"
                        case "rep"
							response.write "<b>Replace:</b> Old Part Number " & p_strOldPartNumber & " (" & p_strDeliverableName & ") with " & p_strNewPartNumber & " in Product " & p_strProduct & ", Product Drop (" & p_strProductDropName & ", ID=" & p_strProductDropID & ").</br>"
                        case "rem"	'Remove a component
							response.write "<b>Remove:</b> Part Number " & p_strNewPartNumber & " (" & p_strDeliverableName & ") from Product " & p_strProduct & ", Product Drop (" & p_strProductDropName & ", ID=" & p_strProductDropID & ").</br>"
                 end select
                Response.Write strSpaces & strSpanERROR & "<span id='ErrSpan|" & p_strProductDropID & "|" & p_strDeliverableName & "|'>"  & "The part(s) selected do not exist in any product drop(s) where you have the approval to modify. Please verify you are the approver of the product(s) and resubmit your search and request. You can view approver(s) in Product Explorer\Product Drop Properties tab.</span></span></br>" & vbCrLf
            end if
	
		next
	end if
	response.write ("</br>Completed")
%>
<br/>
<br />
<div id="CompUpdDiv"><a id="CompUpdLink" href="#"> Component Update List </a></div>
</BODY>
</HTML>
