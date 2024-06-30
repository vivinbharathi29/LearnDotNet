
<script runat="server" language="vbscript">

	dim cn
	dim rs
	dim cm
	dim p
	dim strRegionMatrix 

	dim strImageIDList
	dim strImageNameList 
	dim strImageTag
	
	dim strShowEditBoxes
	dim strActiveColor	
	dim strDevCenter
	dim CurrentUserPinPm
	dim ImageMasterSkuComp
    dim strBusinessSegmentID
	
    dim strTierDropDown
    dim strCommentDropDown
    dim strSelected
	dim ReleaseID
	
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	
		
	
	strBusinessSegmentID = request("BusinessSegmentID")
	ReleaseID = request("ProductReleaseID")
	CurrentUserPinPm = request("CurrentUserPinPm")
	strShowEditBoxes = request("ShowEditBoxes")
	strDevCenter = request("DevCenter")
	
if (request("ID") <>"" or request("ProdID") <>"") and ReleaseID > 0 then  
		rs.Open "usp_SelectProdBrandConfigs " & clng(request("ProdID")) & ",'" & ReleaseID &"'",cn,adOpenForwardOnly
	else 
		rs.Open "usp_SelectProdBrandConfigs " & clng(request("ProdID")) & "," & 0 ,cn,adOpenForwardOnly
	end if 
	
    

     strConfigs = ""

     do while not rs.EOF
        if strConfigs = "" then
            strConfigs = rs("OptionConfig")
        else
	        strConfigs = strConfigs & "," & rs("OptionConfig")
	    end if
	    rs.MoveNext
	 loop 
	 rs.Close


	'Build Region Matrix
	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	cm.CommandType = 4
	cm.CommandText = "spListRegionsForImage2"
		

	Set p = cm.CreateParameter("@ID", 3, &H0001)
	if request("CopyID") <> "" then
		p.Value = request("CopyID")
	elseif request("ID") = "" then
		p.Value = 0
	else
		p.Value = request("ID")
	end if
	cm.Parameters.Append p
	

	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	Set cm=nothing

	
	strRegionMatrix = "<TABLE  bgcolor=""cornsilk"" bordercolor=""tan"" border=""1"" cellpadding=""1"" cellspacing=""0"" width=""100%"" id=""regionTable"">"
	strImageIDList = ""
	strImageNameList = ""
	strImageTag = ""
	strCopyTag = ""
	strActiveColor =""
	strTabRowID = ""
	TabRowIndex = 0
	ImageMasterSkuComp = ""
    strTagPublish = ""
    strRegionMatrixBottom = ""
    strLastGeo = ""
    strSelected = ""
	do while not rs.EOF
        strIssues = ""
        strTemp = ""
        strTierDropDown = ""
		if rs("Active") or trim(rs("Priority") & "" ) <> "" then
			strImageIDList = strImageIDList & "," & rs("ID")
			strImageNameList = strImageNameList & "," & rs("Name")
			if rs("optionconfig") & "" <> "" then
			    strImageNameList = strImageNameList & " (" & rs("optionconfig") & ")"
			end if
			if rs("Active") then
				strActiveColor = " bgcolor=""cornsilk"" "
            else
				strActiveColor = " bgcolor=""mistyrose"" " 'grey
			    strIssues = strIssues & "<BR>Localization is Inactive"
            end if

			if rs("Geo") & "" <> strLastGeo then
				strRegionMatrix = strRegionMatrix & "<TR bgcolor=""wheat"" class=""Header""><TD colspan=13><font color=black size=2 face=verdana><b>" & rs("Geo") & "</b></font></td></tr>"
		        strRegionMatrix = strRegionMatrix & "<TR bgcolor=""cornsilk"" class=""Header""><TD valign=bottom width=10>&nbsp;</TD>"
                strRegionMatrix = strRegionMatrix & "<TD valign=bottom><font size=2 face=verdana><b>Tier</b></font></TD>"
                strRegionMatrix = strRegionMatrix & "<TD valign=bottom><font size=2 face=verdana><b>#Code</b></font></TD>"
		        strRegionMatrix = strRegionMatrix & "<TD valign=bottom nowrap><font size=2 face=verdana><b>Name</b></font></TD>"
		        strRegionMatrix = strRegionMatrix & "<TD valign=bottom><font size=2 face=verdana><b>OS&nbsp;Lang</b></font></TD>"
                strRegionMatrix = strRegionMatrix & "<TD valign=bottom><font size=2 face=verdana><b>Channel&nbsp;Partners</b></font></TD>"
		        strRegionMatrix = strRegionMatrix & "<TD valign=bottom><font size=2 face=verdana><b>GM&nbsp;</b></font></TD>"
		        if CurrentUserPinPm Then
			        strRegionMatrix = strRegionMatrix & "<TD valign=bottom><font size=2 face=verdana><b>Master SKU Comp.</b></font></TD>"
                End If
		        strRegionMatrix = strRegionMatrix & "<TD valign=bottom><font size=2 face=verdana><b>Issues&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b></font></TD>"
		        strRegionMatrix = strRegionMatrix & "<TD valign=bottom><font size=2 face=verdana><b>Regions Comment&nbsp;</b></font></TD>"
		        strRegionMatrix = strRegionMatrix & "<TD valign=bottom><font size=2 face=verdana><b>Publish&nbsp;</b></font></TD>"
                strRegionMatrix = strRegionMatrix & "<TD style=""display:none""><font size=2 face=verdana><b>channelpartnerid</b></font></TD>"
		        strRegionMatrix = strRegionMatrix & "</TR>"
            end if
			strLastGeo = rs("GEO") & ""
                
			Config = trim(rs("OptionConfig") & "")
            Priority = trim(rs("Priority") & "")
			if (instr(";" & trim(rs("businesssegmentids") & ""),";" & strBusinessSegmentID & ";")  or (trim(rs("Priority") & "" ) <> "")) then
			    if trim(Priority) <> "" and instr(strConfigs, Config) = 0 then
			        StrRegionClass = "NotSupported"
			    elseif trim(Priority) <> "" and instr(strConfigs, Config) > 0 then
			        StrRegionClass = "Image"
			    elseif trim(Priority) = "" and instr(strConfigs, Config) > 0 then
			        StrRegionClass = "Product"
			    else 
                    StrRegionClass = "All"
			    end if
                if instr(strConfigs, Config) = 0 then
       			    strIssues = strIssues & "<BR>Localization is not supported"
                end if
                if ((clng(strDevCenter) <> 2 and rs("Consumer") and not rs("Commercial")) ) and (trim(rs("Priority") & "" ) <> "") then
                    strActiveColor = " bgcolor=""mistyrose"" " 
                    strIssues = strIssues & "<BR>Consumer Image Selected"
                elseif ((clng(strDevCenter) = 2 and rs("Commercial") and not rs("Consumer") )) and (trim(rs("Priority") & "" ) <> "") then
                    strActiveColor = " bgcolor=""mistyrose"" " 
                    strIssues = strIssues & "<BR>Commercial Image Selected"
                end if
			    strTemp = strTemp & "<TR" & strActiveColor & " id=""regionRow" & rs("ID") & """ class=""" & StrRegionClass & """>"
            else
			    strTemp = strTemp & "<TR" & strActiveColor & "style=""display:none"" class=""Hidden"">"
			end if
			if trim(rs("priority") & "") = "1" then
                strTemp = strTemp & "<TD><input onclick=""javascript: PriorityChange(" & rs("ID") & ");"" checked id=""chkRegion"" name=""chkRegion"" type=""checkbox"" value=""" & rs("ID") & """ /></TD>"
                strImageTag = strImageTag & "," & trim(rs("ID"))        
            else
                strTemp = strTemp & "<TD><input onclick=""javascript: PriorityChange(" & rs("ID") & ");"" id=""chkRegion"" name=""chkRegion"" type=""checkbox"" value=""" & rs("ID") & """ /></TD>"
            end if
            for tier = 0 to 11
                if tier = rs("Tier") then
                    strSelected = "selected"
                else
                    strSelected = ""
                end if
                if tier = 0 then
                    strTierDropDown = strTierDropDown & "<OPTION " & strSelected & " Value=""" & tier & """></OPTION>"
                elseif tier = 11 then
                    strTierDropDown = strTierDropDown & "<OPTION " & strSelected & " Value=""" & tier & """>NA</OPTION>"
                else
                    strTierDropDown = strTierDropDown & "<OPTION " & strSelected & " Value=""" & tier & """>" & tier & "</OPTION>"
                end if
            next
            strTierDropDown = "<SELECT style=""width:40px"" id=cboTier" & rs("ID") & " name=cboTier" & rs("ID") & " >" & strTierDropDown & "</SELECT>"           
            strTemp = strTemp & "<TD><font face=verdana size=2>" & strTierDropDown & "</font><input type=""hidden"" name=""lblTier" & rs("ID") & """ id=""lblTier" & rs("ID") & """ value=""" & rs("Tier") & """ /></TD>" 'column value for Tier
			strTemp = strTemp & "<TD><font face=verdana size=2>" & rs("OptionConfig") & "&nbsp;</font></TD>"
			strTemp = strTemp & "<TD><INPUT type=""hidden"" id=txtDisplay name=txtDisplay value=""" & rs("DisplayName") & """><font face=verdana size=2 nowrap>" & rs("Name") & "</font></TD>"
			if trim(rs("OtherLanguage") & "") <> "" then
				strTemp = strTemp & "<TD><font face=verdana size=2><u>" & rs("OSLanguage") & "</u>," & rs("OtherLanguage") & "</font></TD>"
			else
				strTemp = strTemp & "<TD><font face=verdana size=2><u>" & rs("OSLanguage") & "</u> </font></TD>"
			end if
            if trim(rs("Priority") & "") <> "" then
                if rs("ChannelPartners") & "" <> "" then
                    strTemp = strTemp & "<TD><font face=verdana size=2><a href=""#"" id=""channelPartners" & rs("ID") & """ onclick=""channelPartners_onclick(" & rs("ID") & ");"">" & rs("ChannelPartners") & "</a>&nbsp;</font></TD>"
			    else
                    strTemp = strTemp & "<TD><font face=verdana size=2><a href=""#"" id=""channelPartners" & rs("ID") & """ onclick=""channelPartners_onclick(" & rs("ID") & ");"">Add</a>&nbsp;</font></TD>"
			    end if
            else
                strTemp = strTemp & "<TD><font face=verdana size=2><a style=""display:none"" href=""#"" id=""channelPartners" & rs("ID") & """ onclick=""channelPartners_onclick(" & rs("ID") & ");"">Add</a>&nbsp;</font></TD>"
            end if
            strTemp = strTemp & "<TD><font face=verdana size=2>" & rs("GMCode") & "&nbsp;</font></TD>"
			If CurrentUserPinPm and rs("ImageId") & "" <> "" Then
			    ImageMasterSkuComp = rs("DriveName") & ""
			    If ImageMasterSkuComp = "" Then ImageMasterSkuComp = "[ Default ]"
			    strTemp = strTemp & "<TD><font face=verdana size=2><a href=""#"" id=""msc" & rs("ImageId") & """ onclick=""msc_onclick(" & rs("ImageId") & ");"">" & ImageMasterSkuComp & "</a>&nbsp;</font></TD>"
			ElseIf CurrentUserPinPm Then
			    strTemp = strTemp & "<TD>&nbsp;</TD>"
			End If
		    if left(strIssues,4) = "<BR>" then
                strIssues = mid(strIssues,5)
            end if
            strTemp = strTemp & "<TD><font face=verdana size=2>" & strIssues & "&nbsp;</font></TD>"

            
            strCommentDropDown = "<SELECT face=verdana id=""ddRegionsComment" & rs("ID") & """ name=""ddRegionsComment" & rs("ID") & """><OPTION value=""""></OPTION>"         
    
            Select case rs("RegionsComment")
                case "JP-Office Personal"
                     strCommentDropDown = strCommentDropDown & "<OPTION selected value=""JP-Office Personal"">JP-Office Personal</OPTION>"
                     strCommentDropDown = strCommentDropDown &  "<OPTION value=""JP-Office Home&Business"">JP-Office Home&Business</OPTION>" 
                     strCommentDropDown = strCommentDropDown &  "<OPTION value=""JP-Office Professional"">JP-Office Professional</OPTION>"
                case "JP-Office Home&Business"
                     strCommentDropDown = strCommentDropDown & "<OPTION value=""JP-Office Personal"">JP-Office Personal</OPTION>"
                     strCommentDropDown = strCommentDropDown & "<OPTION selected value=""JP-Office Home&Business"">JP-Office Home&Business</OPTION>"
                     strCommentDropDown = strCommentDropDown &  "<OPTION value=""JP-Office Professional"">JP-Office Professional</OPTION>"
                case "JP-Office Professional"
                     strCommentDropDown = strCommentDropDown & "<OPTION value=""JP-Office Personal"">JP-Office Personal</OPTION>"
                     strCommentDropDown = strCommentDropDown &  "<OPTION value=""JP-Office Home&Business"">JP-Office Home&Business</OPTION>" 
                     strCommentDropDown = strCommentDropDown &  "<OPTION selected value=""JP-Office Professional"">JP-Office Professional</OPTION>"
                case else
                     strCommentDropDown = strCommentDropDown & "<OPTION value=""JP-Office Personal"">JP-Office Personal</OPTION>"
                     strCommentDropDown = strCommentDropDown &  "<OPTION value=""JP-Office Home&Business"">JP-Office Home&Business</OPTION>" 
                     strCommentDropDown = strCommentDropDown &  "<OPTION value=""JP-Office Professional"">JP-Office Professional</OPTION>"   
            End Select 
            
            strCommentDropDown = strCommentDropDown & "</SELECT>" 
            strTemp = strTemp & "<TD>" & strCommentDropDown & "<input type=""hidden"" name=""lblRegionsComment" & rs("ID") & """ id=""lblRegionsComment" & rs("ID") & """ value=""" & rs("RegionsComment") & """ /></TD>"
            strTemp = strTemp & "<TD align=""center"">"
            if not rs("Published") then
                if trim(rs("Priority") & "") <> "" then
                    strTemp = strTemp & "<input id=""chkPublish" & trim(rs("ID")) & """ value=""" & trim(rs("ID")) & """ name=""chkPublish"" type=""checkbox""/>"
                else
                    strTemp = strTemp & "<input disabled style=""display:none"" id=""chkPublish" & trim(rs("ID")) & """ value=""" & trim(rs("ID")) & """ name=""chkPublish"" type=""checkbox""/>"
                end if
            else
                if trim(rs("Priority") & "") <> "" then
                    strTemp = strTemp & "<input checked style=""display:" & strShowEditBoxes & """ id=""chkPublish" & trim(rs("ID")) & """ name=""chkPublish"" value=""" & trim(rs("ID")) & """ type=""checkbox""/>"
                    strTagPublish = strTagPublish & "," & rs("ID")
                else
                    strTemp = strTemp & "<input style=""display:none"" checked id=""chkPublish" & trim(rs("ID")) & """ name=""chkPublish"" value=""" & trim(rs("ID")) & """ type=""checkbox""/>"
                    strTagPublish = strTagPublish & "," & rs("ID")
                end if
            end if
            strTemp = strTemp & "&nbsp;</TD>"    
            strTemp = strTemp & "<TD style=""display:none""><input type=""text"" value=""" & rs("ChannelPartnerIDs") & """ name=""channelPartnerIDs" & rs("ID") & """ id=""channelPartnerIDs" & rs("ID") & """><input type=""text"" value=""" & rs("ChannelPartnerIDs") & """ name=""channelPartnerIDsOri" & rs("ID") & """ id=""channelPartnerIDsOri" & rs("ID") & """></TD>"
			strTemp = strTemp & "</TR>"
            strRegionMatrix = strRegionMatrix & strTemp

		end if
		rs.MoveNext
	loop
	rs.Close

	strRegionMatrix = strRegionMatrix & "</TABLE>"

    if len(strImageTag) > 0 then
		strImageTag = mid(strImageTag,2)
	end if

	if len(strCopyTag) > 0 then
		strCopyTag = mid(strCopyTag,2)
	end if

	if len(strImageIDList) > 0 then
		strImageIDList = mid(strImageIDList,2)
	end if

	if len(strImageNameList) > 0 then
		strImageNameList = mid(strImageNameList,2)
	end if
	
	if strTagPublish<> "" then
		strTagPublish = mid(replace(strTagPublish," ",""),2)
	end if

    strRegionMatrix = strRegionMatrix + "<INPUT type=""hidden"" id=txtImageIDList name=txtImageIDList value=" + strImageIDList + ">"
	if  request("CopyID") <>  "" then
		strRegionMatrix = strRegionMatrix + "<INPUT type=""hidden"" id=txtTag name=txtTag value=" + strCopyTag + ">"
    else
		strRegionMatrix = strRegionMatrix + "<INPUT type=""hidden"" id=txtTag name=txtTag value=" + strImageTag + ">"
	end if 
	strRegionMatrix = strRegionMatrix + "<INPUT type=""hidden"" id=txtimageNameList name=txtimageNameList value='" + strImageNameList + "'>"
	strRegionMatrix = strRegionMatrix + "<INPUT  type=""hidden"" id=tagPublish name=tagPublish value=" + strTagPublish + ">"
	
	
	response.write strRegionMatrix
	cn.Close
</script> 