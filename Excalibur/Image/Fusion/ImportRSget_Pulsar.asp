<!-- #include file="../../includes/DataWrapper.asp" -->

<script runat="server" language="vbscript">

dim strResult
strResult = ""

    dim intSourceProductID, intTargetProductID
    Dim rs, dw, cn, cmd
    Dim strReleaseList, strCmd, intReleaseId

    intSourceProductID = request.QueryString("SourceProductID")
    intTargetProductID = request.QueryString("TargetProductID")
    intReleaseId = request.QueryString("ProductReleaseID") 
    if intReleaseId = "" or IsNull(intReleaseId) then
		intReleaseId = 0 
	else
		intReleaseId = cLng(intReleaseId)
	end if 

  	Set dw = New DataWrapper
    Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
    Set rs = Server.CreateObject("ADODB.RecordSet")

    'get productversionrelease 
    strReleaseList = ""
    strcmd = "Select pvr.ID as ProductVersionReleaseID, r.Name as ReleaseName From ProductVersion_Release pvr inner join ProductVersionRelease r on pvr.ReleaseID = r.ID where ProductVersionID=" & request.QueryString("TargetProductID")
	Set cmd = dw.CreateCommandSQL(cn, strcmd)
	Set rs = dw.ExecuteCommAndReturnRS(cmd)
	 
	do while not rs.EOF
		if clng(intReleaseId) = rs("ProductVersionReleaseID") then
			strReleaseList = strReleaseList & "<Option selected value=""" & rs("ProductVersionReleaseID") & """>" & rs("ReleaseName") & "</Option>" 
		elseif rs("ProductVersionReleaseID") > 0 then
			strReleaseList = strReleaseList & "<Option value=""" & rs("ProductVersionReleaseID") & """>" & rs("ReleaseName") & "</Option>" 
		end if
		rs.MoveNext
	loop
	rs.Close


    Set cmd = dw.CreateCommAndSP(cn, "usp_Image_GetImageDefinitionList")
    dw.CreateParameter cmd, "@ProductID", adInteger, adParamInput, 8, intSourceProductID
    dw.CreateParameter cmd, "@Report", adInteger, adParamInput, 8, -1
    dw.CreateParameter cmd, "@ImageTypeID", adInteger, adParamInput, 8, null
    dw.CreateParameter cmd, "@TargetProductID", adInteger, adParamInput, 8, intTargetProductID
    dw.CreateParameter cmd, "@ProductReleaseID", adInteger, adParamInput, 8, 0

    Set rs = dw.ExecuteCommAndReturnRS(cmd)
	

	if rs.EOF and rs.BOF then 	
		strResult =  "<Table ID=""ImageTable"" width=100% border=0 cellpadding=1 cellspacing=0><TR bgcolor=cornsilk><TD><INPUT style=""width:16;height:16;"" type=""checkbox"" id=chkALL name=chkAll LANGUAGE=javascript></TD><TD><font size=1 face=verdana><b>Product&nbsp;Drop</b></font></TD><TD><font size=1 face=verdana><b>Brand&nbsp;&nbsp;</b></font></TD><TD><font size=1 face=verdana><b>OS&nbsp;&nbsp;</b></font></TD><TD><font size=1 face=verdana><b>Release&nbsp;</b></font></TD><TD><font size=1 face=verdana><b>Software&nbsp;&nbsp;</b></font></TD><TD><font size=1 face=verdana><b>Comments&nbsp;&nbsp;</b></font></TD></tr>"
		strResult = strResult &  "<TR><TD colspan=4><font size=1 face=verdana>The selected Product does not have any Operating Systems defined in the Image Definition.</font></TD></TR></table>"
	else
        if rs("UsedInPRL") = "0" then
            strResult = ""
        else
		    strResult =  "<div style=""margin-bottom:5px;margin-top:5px""><font size=1 face=verdana>Only Operating Systems matching your PRL are enabled for importing. Please follow the correct process to add additional Operating Systems to your PRL then import again.</font></div>"
        end if
        strResult = strResult & "<Table ID=""ImageTable"" width=100% border=0 cellpadding=1 cellspacing=0><TR bgcolor=cornsilk><TD><INPUT style=""width:16;height:16;"" type=""checkbox"" checked id=chkALL name=chkAll LANGUAGE=javascript onclick=""return chkAll_onclick()""></TD><TD><font size=1 face=verdana><b>Product&nbsp;Drop</b></font></TD><TD><font size=1 face=verdana><b>Brand&nbsp;&nbsp;</b></font></TD><TD><font size=1 face=verdana><b>OS&nbsp;&nbsp;</b></font></TD><TD><font size=1 face=verdana><b>Release&nbsp;&nbsp;</b></font></TD><TD><font size=1 face=verdana><b>Comments&nbsp;&nbsp;</b></font></TD></tr>"
		do while not rs.EOF 
                if rs("IsSelectable") = "0" then
				    strResult = strResult & "<TR valign=top bgcolor=lightgray><TD style=""BORDER-TOP: gray thin solid"">&nbsp</td>"
                else
                    strResult = strResult & "<TR valign=top bgcolor=Ivory><TD style=""BORDER-TOP: gray thin solid""><INPUT value=""" & rs("RowKey") & """ checked style=""width:16;height:16;"" type=""checkbox"" id=chkSelected name=chkSelected></td>"
                end if
				strResult = strResult & "<TD nowrap style=""BORDER-TOP: gray thin solid""><font size=1 face=verdana>" & rs("ProductDrop") & "&nbsp;</font></TD>"
				strResult = strResult & "<TD style=""BORDER-TOP: gray thin solid""><font size=1 face=verdana>" & rs("Brand") & "&nbsp;</font></TD>"
				strResult = strResult & "<TD style=""BORDER-TOP: gray thin solid""><font size=1 face=verdana>" & rs("OS") & "&nbsp;</font></TD>"
                if rs("IsSelectable") = "0" then
					strResult = strResult & "<TD style=""BORDER-TOP: gray thin solid""><font size=1 face=verdana>" & rs("ReleaseName") & "&nbsp;</font></TD>"
				else
					strResult = strResult & "<TD style=""BORDER-TOP: gray thin solid""><SELECT id=cboRelease" & rs("RowKey") &" name=cboRelease" & rs("RowKey") & ">" & strReleaseList &"</SELECT></TD>"
				end if
				strResult = strResult & "<TD style=""BORDER-TOP: gray thin solid""><font size=1 face=verdana>" & rs("Comments") & "&nbsp;</font></TD>"
			rs.MoveNext
		loop
		strResult = strResult &   "</table>"
	end if
	rs.Close
    set rs = nothing
    cn.Close
    set cn=nothing    

	response.Write strResult


</script> 

