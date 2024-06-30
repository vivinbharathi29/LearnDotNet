<%@ Language=VBScript %>
<!-- #include file="../../includes/lib_debug.inc" -->
<!-- #include file="../../includes/emailqueue.asp" -->
<!-- #include file="../../includes/UserInfo.asp" -->

<%
    'printrequest
    'response.End
%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    function window_onload() {
        var txtSuccess = document.getElementById("txtSuccess").value;
        var txtError = document.getElementById("txtError").value;

        if (txtSuccess != "0" && txtError == "") {            
            try {
                if ('<%=Request.QueryString("pulsarplusDivId")%>' != '') {
                    parent.window.parent.closeExternalPopup();
                    parent.window.parent.reloadFromPopUp('<%=Request.QueryString("pulsarplusDivId")%>');
                }
                else {
                    window.parent.AddPlatformCompleted();
                }
            }
            catch (e) {
                if (parent.window.parent.document.getElementById('modal_dialog')) {
                    parent.window.parent.modalDialog.cancel(true);
                } else {
                    window.returnValue = txtSuccess;
                    window.close();
                }
                
            }
        }
        else {
            alert(txtError);
            window.history.back();            
        }
}

//-->
</SCRIPT>
</HEAD>
<BODY LANGUAGE=javascript onload="return window_onload()">
<%
	dim cn
    dim rowschanged
	dim cm
	dim blnErrors
    dim strSeparator
    dim chrID
    dim chrTable
    dim chrFieldValues
    dim chrPersonName
    dim chrPersonFullName
    dim chrChildLinkID
    dim chrReturnID
    dim strBrand
    dim strCycle
    dim chkCycle
    dim strSelectedCycle
    dim i
    dim strSuccess
    dim ArrCycle
    dim strBusinessSegment
    dim strMode
    dim strSOARID
    dim sStatus
    dim sError
    dim chkeMMC
    dim followMKTName
    dim mktBrandId:mktBrandId=0
    dim newProductBrandID: newProductBrandID = 0
    dim mktSeries
    dim oldBrandId
    dim mktSeriesID:  mktSeriesID = 0
    dim isDefinition : isDefinition = 0
    dim oldSeriesID : oldSeriesID = 0
    dim changeHis : changeHis=""
    dim productBrandID : productBrandID = 0

    blnErrors = false
    sError = ""
    strSuccess = "0"


    sStatus = Request.Form("ddlStatus")
     if sStatus = "" then
        sStatus ="0"
    end if

    strBusinessSegment = Request.Form("txtBusinessSegmentID")
    if strBusinessSegment  = "" then
       strBusinessSegment = "269"
    end if

    chkCycle = request.form("ckWebCycle")
    ArrCycle  = split(chkCycle,",")
    strSelectedCycle = ""
    followMKTName = cInt(Request.Form("hdnFollowMKTName"))

    for i=LBound(ArrCycle) to UBound(ArrCycle)
      strSelectedCycle = strSelectedCycle + ArrCycle(i) + ","
    Next
    
    if(LEN(strSelectedCycle) > 0) then
            strSelectedCycle = LEFT(strSelectedCycle, LEN(strSelectedCycle) -1)
    end if

    if request.form("ckeMMC") = "on" then
        chkeMMC = 1
    else 
        chkeMMC = 0     
    end if  

    set cn = server.CreateObject("ADODB.Connection")
    cn.ConnectionString = Session("PDPIMS_ConnectionString") 
    cn.Open
    
	
    if (followMKTName = 1) then

        if (request("showDCRDiv") ="none") then
            isDefinition = 1
        end if
      

		if (request.form("txtPlatformID") > 0 and request.form("txtBrandsLoaded") <> request.form("txtBrandsAdded") and request.form("txtBrandFrom")="" and request.form("txtBrandTo")="") then 
            cn.BeginTrans
            set cm = server.CreateObject("ADODB.Command")
			Set cm.ActiveConnection = cn
			cm.CommandType = 4
			cm.CommandText = "spRemoveBrandFromProduct"
			cm.CommandTimeout = 0
			Set p = cm.CreateParameter("@ProductID", 3, &H0001)
			p.Value = request.form("txtProductionVersionID")
			cm.Parameters.Append p

	  	    oldBrandId = request.form("txtBrandsLoaded")
			Set p = cm.CreateParameter("@BrandID", 3, &H0001)
			p.Value = clng(oldBrandId)
	        cm.Parameters.Append p

			cm.execute
			if Err.Number <> 0 or cn.Errors.count <> 0 then
                cn.RollbackTrans
                response.write cn.Errors.number & "<BR>"
                response.write cn.Errors.description & "<BR>"
                response.write cn.Errors.source & "<BR>"
                response.write cn.Errors.sqlstate & "<BR>"
                response.write cn.Errors.nativeerror & "<BR>"  
                sError = cn.Errors.description
                On Error Goto 0
            else
				cn.CommitTrans
                strSuccess = 1
			end if		
            set cm = nothing
			
			oldSeriesID = trim(request.form("txtSeriesIDA" & oldBrandId))
			if (oldSeriesID <>"" ) then
				cn.BeginTrans
				set cm = server.CreateObject("ADODB.Command")
				Set cm.ActiveConnection = cn
				cm.CommandType = 4
				cm.CommandText = "spUpdateSeries"
					
				Set p = cm.CreateParameter("@ID", 3, &H0001)
				p.Value = oldSeriesID
				cm.Parameters.Append p

				Set p = cm.CreateParameter("@Name", 200, &H0001,50)
				p.Value = ""
				cm.Parameters.Append p

				cm.execute
				if Err.Number <> 0 or cn.Errors.count <> 0 then
					cn.RollbackTrans
					sError = cn.Errors.description
					response.write cn.Errors.description & "<BR>"
					On Error Goto 0
				else
					cn.CommitTrans
					strSuccess = 1
				end if
				set cm = nothing
			 end if
        end if 'end if of brand and series deletion and creation
		

        'add new brand/ change brand
        if ((request.form("txtPlatformID") = 0) or (request.form("txtPlatformID") > 0 and request.form("txtBrandsLoaded") <> request.form("txtBrandsAdded") and request.form("txtBrandFrom")="" and request.form("txtBrandTo")="" )) then
            cn.BeginTrans
            set cm = server.CreateObject("ADODB.Command")
	        cm.CommandType =  &H0004
	        cm.ActiveConnection = cn
            
            cm.CommandText = "spAddBrand2Product"
            Set p = cm.CreateParameter("@ProductID", 3, &H0001)
	        p.Value = request.form("txtProductionVersionID")
	        cm.Parameters.Append p
        
            mktBrandId = request.form("chkBrands")
	        Set p = cm.CreateParameter("@BrandID", 3, &H0001)
	        p.Value = clng(mktBrandId)
	        cm.Parameters.Append p
        
            mktSeries = trim(request.form("txtSeriesA" & mktBrandId))
            
	        Set p = cm.CreateParameter("@SeriesSummary", 200, &H0001,2000)
            p.Value = mktSeries
            cm.Parameters.Append p
    
            set p = cm.CreateParameter("@ProductBrandID", adInteger, adParamOutput)
            cm.Parameters.Append p
        
            cm.Execute
            If Err.Number <> 0 or cn.Errors.count > 0 Then 
                 cn.RollbackTran
                sError = Err.Description
                response.write cn.Errors.description & "<BR>"
                On Error Goto 0
            Else
                cn.CommitTrans
                strSuccess = "1"
                newProductBrandID = cm.Parameters.Item("@ProductBrandID").Value
            end if
            Set cm =nothing
        end if
        if (newProductBrandID > 0 )then
            productBrandID = newProductBrandID
        else
            productBrandID =  request.form("hidPBID")
        end if
        
        'add/update series
        if (mktBrandID = 0) then
            mktBrandId = request.form("chkBrands")
        end if
        if (mktSeriesID =0 and request.form("txtSeriesIDA" & mktBrandId)<>"" ) then
            mktSeriesID = clng(request.form("txtSeriesIDA" & mktBrandId))
        end if

        mktSeries = trim(request.form("txtSeriesA" & mktBrandId))
        if( mktSeries <>"" and (request.form("txtPlatformID") = 0 or request.form("tagSeriesA" & mktBrandId) = empty)) then 'add
            cn.BeginTrans
            set cm = server.CreateObject("ADODB.Command")
	        cm.CommandType =  &H0004
	        cm.ActiveConnection = cn
    
            cm.CommandText = "spAddSeries2Brand2"
    	
			Set p = cm.CreateParameter("@ProductID", 3, &H0001)
			p.Value = clng(request.form("txtProductionVersionID"))
			cm.Parameters.Append p
    
			Set p = cm.CreateParameter("@BrandID", 3, &H0001)
		    p.Value = clng(mktBrandId)
		    cm.Parameters.Append p
            
            Set p = cm.CreateParameter("@Name", 200, &H0001, 50)
			p.Value = left(trim(mktSeries), 50)
		    cm.Parameters.Append p
            
			Set p = cm.CreateParameter("@NewID", 3, &H0002)
			cm.Parameters.Append p
            cm.Execute
            
            If Err.Number <> 0 or cn.Errors.count > 0 Then  
                 cn.RollbackTrans
                 sError = Err.Description
                 response.write cn.Errors.description & "<BR>"
                 On Error Goto 0
             else
                 cn.CommitTrans
                 strSuccess = "1"
                 mktSeriesID = cm.Parameters.Item("@NewID").Value
            end if
            set cm=nothing
        end if
        
        if (request.form("tagSeriesA" & mktBrandId) <> empty and (mktSeries <> request.form("tagSeriesA" & mktBrandId))) then 'update series
            cn.BeginTrans
            set cm = server.CreateObject("ADODB.Command")
			Set cm.ActiveConnection = cn
			cm.CommandType = 4
            cm.CommandText = "spUpdateSeries"
					
            Set p = cm.CreateParameter("@ID", 3, &H0001)
            mktSeriesID = request.form("txtSeriesIDA" & mktBrandId)
            p.Value = clng(mktSeriesID)
            cm.Parameters.Append p

            Set p = cm.CreateParameter("@Name", 200, &H0001,50)
            p.Value = mktSeries
            cm.Parameters.Append p

            cm.execute
			if cn.Errors.count <> 0 then
                 cn.RollbackTrans
                 sError = Err.Description
                 response.write cn.Errors.description & "<BR>"
                 On Error Goto 0
            else
                 cn.CommitTrans
                 strSuccess = "1"
                 if (mktSeries ="") then
                        mktSeriesID = 0
                 end if
			end if  
            set cm=nothing
        end if 
        
        if (request.form("txtBrandFrom")<>"" and request.form("txtBrandTo")<>"") then ' only update the brand
             cn.BeginTrans
             set cm = server.CreateObject("ADODB.Command")
	         Set cm.ActiveConnection = cn
	         cm.CommandType = 4
	         cm.CommandText = " spUpdateProductBrand"
                					
	        Set p = cm.CreateParameter("@ProductID", 3, &H0001)
	        p.Value = clng(request.form("txtProductionVersionID"))
	        cm.Parameters.Append p

	        Set p = cm.CreateParameter("@OldBrandID", 3, &H0001)
	        p.Value = clng(request.form("txtBrandFrom"))
	        cm.Parameters.Append p

	        Set p = cm.CreateParameter("@NewBrandID", 3, &H0001)
	        p.Value = clng(request.form("txtBrandTo"))
	        cm.Parameters.Append p

	        cm.execute 
	        If Err.Number <> 0 or cn.Errors.count > 0 Then 
                cn.RollbackTran
                sError = Err.Description
                response.write cn.Errors.description & "<BR>"
                On Error Goto 0
            Else
                cn.CommitTrans
                strSuccess = "1"
                mktSeriesID = request.form("txtSeriesIDA" & request.form("txtBrandFrom"))
            end if
            Set cm =nothing
        end if

        cn.BeginTrans
        set cm = server.CreateObject("ADODB.Command")
	    cm.CommandType =  &H0004
	    cm.ActiveConnection = cn
    
        cm.CommandText = "usp_Product_UpdateBrandNames"					
	    Set p = cm.CreateParameter("@ProductversionID", 3, &H0001)
		p.Value = request.form("txtProductionVersionID")
		cm.Parameters.Append p

        if (mktBrandId = 0) then
            mktBrandId = request.form("chkBrands")
        end if

		Set p = cm.CreateParameter("@BrandID", 3, &H0001)
		p.Value = clng(mktBrandId)
		cm.Parameters.Append p
        
        strmodelnumber = trim(request("txtModelNumber" & mktBrandId))
		Set p = cm.CreateParameter("@ModelNumber", 200, &H0001, 10)
		p.Value = strmodelnumber
		cm.Parameters.Append p
                      
        strscreenSize = trim(request("txtScreenSize" & mktBrandId))
		Set p = cm.CreateParameter("@ScreenSize", 131, &H0001)
		p.Precision = 5
		p.NumericScale = 2
		if (strscreenSize) <> "" then
			p.Value = strscreenSize
		end if
		cm.Parameters.Append p
        
        strgeneration = trim(request("txtGeneration" & mktBrandId))
		Set p = cm.CreateParameter("@Generation", 200, &H0001, 5)
		p.Value = strgeneration
		cm.Parameters.Append p
			
        strFormFactor = trim(request("txtFormFactor" & mktBrandId))
		Set p = cm.CreateParameter("@FormFactor", 200, &H0001,15)
		p.Value = strFormFactor
		cm.Parameters.Append p    
        
        if request("txtSCMEnabled" & mktBrandId) ="1" then
            SCMNumber = trim(request("SelectSCM" & mktBrandId))
        else 'if disabled, get the value from the hidden txtbox
             SCMNumber = trim(request("txtSelectedSCM" & mktBrandId))
        end if
		Set p = cm.CreateParameter("@SCMNumber", 3, &H0001)
		p.Value = SCMNumber
		cm.Parameters.Append p                
        cm.execute
        If Err.Number <> 0 or cn.Errors.count > 0 Then  
            cn.RollbackTrans
            sError = Err.Description
            response.write cn.Errors.description & "<BR>"
            On Error Goto 0
        Else
            cn.CommitTrans
            strSuccess = "1"
        end if
        set cm=nothing
        
        if (request.form("hidSentPMSReq") = 1 and mktSeriesID = 0) then
            cn.BeginTrans
            cn.Execute "UPDATE Product_Brand SET SentPMSReq = " & request.form("hidSentPMSReq") & " WHERE ProductVersionId = " & request.form("txtProductionVersionID") & " AND BrandId = "& mktBrandId
		    if cn.Errors.count <> 0 then	
                cn.RollbackTrans
                sError = Err.Description
                response.write cn.Errors.description & "<BR>"
                On Error Goto 0
            else
                cn.CommitTrans
                strSuccess = "1" 
            end if
            set cm=nothing
        end if

        if(mktSeriesID > 0) then
            cn.BeginTrans
            cn.Execute "UPDATE Series Set SentPMSReq = " & request.form("hidSentPMSReq") & " WHERE ID = " & mktSeriesID
		    if cn.Errors.count <> 0 then	
                cn.RollbackTrans
                sError = Err.Description
                response.write cn.Errors.description & "<BR>"
                On Error Goto 0
            else
                cn.CommitTrans
                strSuccess = "1" 
            end if
            set cm=nothing
        end if
       
        if (request.form("txtPlatformID") > 0) then
            if (isDefinition = 0) then          
                if (request("tagBrandId") <> mktBrandId ) then
                    changeHis = changeHis & "<ProductBrand><Field>BrandId</Field><From>" & request("tagBrandId") & "</From><To>" & mktBrandId & "</To></ProductBrand>"
                end if
                if (request("tagSCMNumber") <> SCMNumber) then
                    changeHis = changeHis & "<ProductBrand><Field>SCMNumber</Field><From>" & request("tagSCMNumber") & "</From><To>" & SCMNumber & "</To></ProductBrand>"
                end if
                if (request("tagSeries") <> mktSeries) then
                    changeHis = changeHis & "<ProductBrand><Field>Series</Field><From>" & request("tagSeries") & "</From><To>" & mktSeries & "</To></ProductBrand>"
                end if
                if (request("tagModelNumber") <> strmodelnumber) then
                    changeHis = changeHis & "<ProductBrand><Field>ModelNumber</Field><From>" & request("tagModelNumber") & "</From><To>" & strmodelnumber & "</To></ProductBrand>"
                end if
                if (request("tagScreenSize") <> strscreenSize) then
                    changeHis = changeHis & "<ProductBrand><Field>ScreenSize</Field><From>" & request("tagScreenSize") & "</From><To>" & strscreenSize & "</To></ProductBrand>"
                end if
                if (request("tagGeneration") <> strgeneration) then
                    changeHis = changeHis & "<ProductBrand><Field>Generation</Field><From>" & request("tagGeneration") & "</From><To>" & strgeneration & "</To></ProductBrand>"
                end if
                if (request("tagFF") <> strFormFactor) then
                    changeHis = changeHis & "<ProductBrand><Field>FormFactor</Field><From>" & request("tagFF") & "</From><To>" & strFormFactor & "</To></ProductBrand>"
                end if
            
               if LEN(changeHis) > 0 then
                    changeHis = "<?xml version='1.0' encoding='iso-8859-1' ?>" + changeHis
                    cn.BeginTrans
                    set cm = server.CreateObject("ADODB.Command")
	                Set cm.ActiveConnection = cn
	                cm.CommandType = 4
	                cm.CommandText = "usp_ProductBrand_AddLog" 
	            
                    Set p = cm.CreateParameter("@ProductVersionID", 3, &H0001)
	                p.Value = request.form("txtProductionVersionID")
	                cm.Parameters.Append p
                
	                Set p = cm.CreateParameter("@BrandID", 3, &H0001)
	                p.Value = mktBrandId
	                cm.Parameters.Append p
    
	                Set p = cm.CreateParameter("@p_intDCRID", 3, &H0001)
	                p.Value = Request.Form("cboDCR")
	                cm.Parameters.Append p

                    Set p = cm.CreateParameter("@p_intUserID", 3, &H0001)
	                p.Value = request.form("txtCurrentUserID")
	                cm.Parameters.Append p

                    Set p = cm.CreateParameter("@p_xmlChangeLog",adLongVarChar, &H0001,len(changeHis))
		            p.Value = changeHis
		            cm.Parameters.Append p
    
	                cm.execute
	                set cm = nothing
	                If Err.Number <> 0 or cn.Errors.count > 0 Then    
                        cn.RollbackTrans
                        sError = Err.Description
                        response.write cn.Errors.description & "<BR>"
                        On Error Goto 0
                     Else
                        cn.CommitTrans
                        strSuccess = "1"
                    end if		
                end if 'end if of Len > 0
            end if 'end if of isDefinition = 0

            if (Request.Form("hidMktName") <> Request.Form("tagMktName")) then
                cn.BeginTrans
                set cm = server.CreateObject("ADODB.Command")
	            Set cm.ActiveConnection = cn
	            cm.CommandType = 4
	            cm.CommandText = "usp_ProductBrand_SentMKTChangeMail" 
            
                Set p = cm.CreateParameter("@PVID", 3, &H0001)
	            p.Value = request.form("txtProductionVersionID")
	            cm.Parameters.Append p

                Set p = cm.CreateParameter("@OldMktName", 200, &H0001, 60)
	            p.Value = request.form("tagMktName")
	            cm.Parameters.Append p
            
                Set p = cm.CreateParameter("@NewMktName", 200, &H0001, 60)
	            p.Value = request.form("hidMktName")
	            cm.Parameters.Append p

                cm.execute
	            set cm = nothing
	            If Err.Number <> 0 or cn.Errors.count > 0 Then    
                    cn.RollbackTrans
                    sError = Err.Description
                    response.write cn.Errors.description & "<BR>"
                    On Error Goto 0
                Else
                    cn.CommitTrans
                    strSuccess = "1"
                End if	
            end if 'end if of send mktName changed email
        end if ' end if of platformId > 0
    end if 'end if of followMKT
    '===================================end of add/update brand & series=================================
    
    'add/update platform
    chrPersonFullName= request.form("txtCurrentUserName")
		
	cn.BeginTrans
	set cm = server.CreateObject("ADODB.Command")
    '&H0004 means stored procedure
    '&H0001 means paramInput
    '&H0002 = paramOutput

	cm.CommandType =  &H0004
	cm.ActiveConnection = cn
    cm.CommandTimeout = 0

    cm.CommandText = "spFusion_AddPlatforms"
   
    Set p = cm.CreateParameter("@p_intProductVersionID", 3, &H0001)
	p.Value = request.form("txtProductionVersionID")
	cm.Parameters.Append p
   
    Set p = cm.CreateParameter("@p_intPlatformID", 3, &H0001)
	p.Value = request.form("txtPlatformID")
	cm.Parameters.Append p


	Set p = cm.CreateParameter("@p_intIntroYear", 200, &H0001, 4)
	p.Value = request.form("txtIntroYear")
	cm.Parameters.Append p

    Set p = cm.CreateParameter("@p_chrProductFamily", 200, &H0001, 64)
	p.Value = Request.Form("txtProductFamily")
	cm.Parameters.Append p

    Set p = cm.CreateParameter("@p_chrProduct", 200, &H0001, 64)
	p.Value = Request.Form("txtPCAName")
	cm.Parameters.Append p

    Set p = cm.CreateParameter("@p_intChassis", 3, &H0001)
	p.Value = Request.Form("ddlChassis")
	cm.Parameters.Append p
   
    Set p = cm.CreateParameter("@p_intBrandID", 3, &H0001)
    if followMKTName = 0 then 
	    p.Value = Request.Form("ddlBrand")
    else
        p.Value = productBrandID
    end if 
	cm.Parameters.Append p
   
    Set p = cm.CreateParameter("@p_chrGenericName", 200, &H0001, 60)
    if followMKTName = 0 then 
	    p.Value = Request.Form("txtGenericName")
    else
        p.Value = Request.Form("hidMktName") 
    end if 
	cm.Parameters.Append p

    Set p = cm.CreateParameter("@p_chrMarketingName ", 200, &H0001, 60)
    if followMKTName = 0 then 
	    p.Value =  Request.Form("txtMarketingName")  
    else
        p.Value = Request.Form("hidMktName")  
    end if 
	cm.Parameters.Append p

    Set p = cm.CreateParameter("@p_chrFullMarketingName", 200, &H0001, 120)
    if followMKTName = 0 then 
	    p.Value = Request.Form("txtFullMarketingName")
    else
        p.Value=""
    end if
	cm.Parameters.Append p

    Set p = cm.CreateParameter("@p_chrSOARDescription", 200, &H0001, 150)
	p.Value = Request.Form("ddlSoarProductDescription")
	cm.Parameters.Append p

    Set p = cm.CreateParameter("@p_chrBIOSName", 200, &H0001, 40)
	p.Value = Request.Form("txtBiosName")
	cm.Parameters.Append p

    Set p = cm.CreateParameter("@p_chrChinaIdentifier", 200, &H0001, 225)
	p.Value = ""
	cm.Parameters.Append p

    Set p = cm.CreateParameter("@p_chrComments", 200, &H0001, 64)
	p.Value = Request.Form("txtComments")
	cm.Parameters.Append p

    Set p = cm.CreateParameter("@p_chrSystemID", 200, &H0001, 4)
	p.Value = Request.Form("txtSystemID") 
	cm.Parameters.Append p

    Set p = cm.CreateParameter("@p_chrCycle", 200, &H0001, 256)
	p.Value = strSelectedCycle
	cm.Parameters.Append p

    Set p = cm.CreateParameter("@p_intBusinessSegmentID", 3, &H0001)
	p.Value = strBusinessSegment 'Request.Form("txtBusinessSegmentID")
	cm.Parameters.Append p

    Set p = cm.CreateParameter("@p_chrSPQDate", 200, &H0001, 10)
	p.Value = Request.Form("txtSPQNeededDate") 
	cm.Parameters.Append p

    Set p = cm.CreateParameter("@p_intStatus", 3, &H0001)
	p.Value = sStatus
	cm.Parameters.Append p

    Set p = cm.CreateParameter("@p_chrCreator", 200, &H0001, 64)
	p.Value = chrPersonFullName
	cm.Parameters.Append p

    Set p = cm.CreateParameter("@p_chrUpdater", 200, &H0001, 64)
	p.Value = chrPersonFullName
	cm.Parameters.Append p

    Set p = cm.CreateParameter("@p_chrPCAGraphicsType", 200, &H0001, 64)
	p.Value = Request.Form("ddlPCAGraphicsType")
	cm.Parameters.Append p

    Set p = cm.CreateParameter("@p_TouchId", 3, &H0001)
	p.Value = request.form("ddlTouch")
	cm.Parameters.Append p

    Set p = cm.CreateParameter("@p_ModelNumber", 200, &H0001, 40)
	p.Value = Request.Form("txtProductModelNumber")
	cm.Parameters.Append p

    Set p = cm.CreateParameter("@p_eMMConboard", 3, &H0001)
	p.Value = clng(chkeMMC)
	cm.Parameters.Append p

    Set p = cm.CreateParameter("@p_chrMemoryOnboard", 200, &H0001, 64)
	p.Value = Request.Form("ddlMemoryOnboard")
	cm.Parameters.Append p

    Set p = cm.CreateParameter("@p_chrGraphicCapacity", 200, &H0001, 64)
	p.Value = Request.Form("ddlGraphicCapacity")
	cm.Parameters.Append p

    Set p = cm.CreateParameter("@p_SystemIdComments", 200, &H0001, 256)
	p.Value = Request.Form("txtSystemIdComments") 
	cm.Parameters.Append p

    Set p = cm.CreateParameter("@p_chrDeployment", 200, &H0001, 30)
	p.Value = Request.Form("ddlDeployment") 
	cm.Parameters.Append p

    Set p = cm.CreateParameter("@p_BIOSFamilyName", 200, &H0001, 48)
	p.Value = Request.Form("txtBFName") 
	cm.Parameters.Append p

    if followMKTName = 1 then 
        Set p = cm.CreateParameter("@p_intSeriesID", 3, &H0001)
        p.Value = mktSeriesID
        if (mktSeriesID = 0) then
            p.Value = null
        end if
        cm.Parameters.Append p
                
        set p = cm.CreateParameter("@p_phWebFamilyName", 200, &H0001, 255)
        p.Value = Request.Form("hidphwFamilyName") 
        cm.Parameters.Append p

        set p = cm.CreateParameter("@p_matchPMS", 3, &H0001)
        if (Request.Form("hidPMS") = "Match") then
            p.Value = 1
        else 
            p.value = 0
        end if
        cm.Parameters.Append p
    else
        Set p = cm.CreateParameter("@p_intSeriesID", 3, &H0001)
		cm.Parameters.Append p

		set p = cm.CreateParameter("@p_phWebFamilyName", 200, &H0001, 255)
		cm.Parameters.Append p

		set p = cm.CreateParameter("@p_matchPMS", 3, &H0001)
		cm.Parameters.Append p
    end if 
    
    Set p = cm.CreateParameter("@p_intReturnID", 3, &H0002)
	cm.Parameters.Append p
 

    Err.Clear
    cm.Execute    
    If Err.Number <> 0 or cn.Errors.count > 0 Then        
        sError = Err.Description
        cn.RollbackTrans
        response.write cn.Errors.description & "<BR>"
        On Error Goto 0
        'Yong removed the code to save platform formfactor (as business function moved to PRL) instead of comment it out to make the page clean; 
        'if for some reason we need to put this back, we just need to look at tfs history
    Else
        cn.CommitTrans
        strSuccess = "1"
		
		Dim currentUserEmail
		set usrInfo = New UserInfo
		currentUserEmail = usrInfo.Currentuseremail
		
		'Notify with Email only when System Board ID changed
		if Request.form("txtPlatformID") <> 0 and trim(Request.Form("txtSystemID")) <> trim(Request.Form("txtInitialSystemID")) then
			set	oMessage = New EmailQueue 		
			oMessage.From = "pulsar.support@hp.com"
			oMessage.To = "psgsoftpaqsupport@hp.com;TWN.PDC.NB-ReleaseLab@hp.com;MobileExcalNotification-SysID@hp.com;" & currentUserEmail
			
			oMessage.Subject = Request.Form("txtProduct") & " System Board ID changed" 
			oMessage.HTMLBody = "<font face=Verdana size=4 color=black><b>" & Request.Form("txtProduct") & "</b><BR><BR></font>" & _
				   "<font face=Verdana size=2 color=black>System Board ID for '" & Request.Form("txtProduct") & "' <Base Unit Group> changed from: " & Request.Form("txtInitialSystemID") & " TO: " & formatsystemid(Request.Form("txtSystemID")) & "<BR><BR></font>" & _
				   "<BR><font face=Arial size=2 color=black>Updates By: " & chrPersonFullName & " - " & now() 
					
			oMessage.SendWithOutCopy
			Set oMessage = Nothing 	
		end if
        
		'Notify with email when BU created or updated
        dim platformID, logoBadge
		If followMKTName = 1 then
            if (request.form("txtPlatformID") > 0) then
				platformID = request.form("txtPlatformID")
			else
				platformID = cm.Parameters.Item("@p_intReturnID").Value
			end if
            
			set cm = server.CreateObject("ADODB.Command")
	        Set cm.ActiveConnection = cn
			set rs = server.CreateObject("ADODB.recordset")
			rs.Open " SELECT pb.LogoBadge " &_
					" FROM ProductVersion_Platform pp WITH (NOLOCK) JOIN Product_Brand pb WITH (NOLOCK) on pp.ProductBrandID = pb.Id "&_
					" WHERE pp.PlatformID = "& platformID & " AND ISNULL(pp.SeriesID,0) = 0 "&_
					" UNION " &_
					" SELECT s.LogoBadge "&_
					" FROM ProductVersion_Platform pp WITH (NOLOCK) JOIN Product_Brand pb WITH (NOLOCK) on pp.ProductBrandID = pb.Id "&_
					" JOIN Series s WITH (NOLOCK) on pp.SeriesID = s.ID where pp.PlatformID = " & platformID, cn, adOpenForwardOnly
					
			if not (rs.EOF and rs.BOF) then
				logoBadge =  rs("LogoBadge")
			end if
			rs.Close

			set	oMessage = New EmailQueue 		
			oMessage.From = "pulsar.support@hp.com"
			oMessage.To = "MobileExcalNotification-ProductNames@hp.com;" & currentUserEmail 
			dim strOutput
			strOutput = strOutput & "<TABLE border=1 dellpadding=2 cellspacing=0>"
			strOutput = strOutput & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>Marketing</b></font></TD></TR>"
			strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Marketing Name</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & Request.Form("hidMktName") & "</font></TD></TR>"
            strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Logo Badge C Cover</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & logoBadge & "</font></TD></TR>"
            if(request.form("txtPlatformID") > 0 and Request.Form("hidfirstPfID") = platformID and Request.Form("tagPhwFamilyName") <> Request.Form("hidphwFamilyName")) then
                strOutput = strOutput & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>Old PHweb</b></font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Family Name</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & Request.Form("tagPhwFamilyName") & "</font></TD></TR>"
                strOutput = strOutput & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>New PHweb</b></font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Family Name</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & Request.Form("hidphwFamilyName") & "</font></TD></TR>"
			    strOutput = strOutput & "</TABLE><BR><BR>"
            elseif(Request.Form("hidfirstPfID") = 0) then
                strOutput = strOutput & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>PHweb</b></font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Family Name</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & Request.Form("hidphwFamilyName") & "</font></TD></TR>"
			    strOutput = strOutput & "</TABLE><BR><BR>"
            else
                strOutput = strOutput & "<TR><TD bgcolor=gainsboro colspan=2><font size=2 face=verdana><b>PHweb</b></font></TD></TR>"
			    strOutput = strOutput & "<TR><TD bgcolor=ivory><b><font size=2 face=verdana>Family Name</font></b></TD><TD bgcolor=ivory><font size=2 face=verdana>" & Request.Form("hidfirstPwfName") & "</font></TD></TR>"
			    strOutput = strOutput & "</TABLE><BR><BR>"
            end if
			
			If request.form("txtPlatformID") = 0 Then
				'PlatformID=0: Create
				oMessage.Subject = Request.Form("txtProduct") & " series definitions created in Pulsar" 
				oMessage.HTMLBody = "<font face=Arial size=3 color=black><b>Series Added</b></font>" & strOutput
				oMessage.SendWithOutCopy
			Else
				'Update
				If Request.Form("initialMktName") <> Request.Form("hidMktName") Then
					'Actually did changes
					oMessage.Subject = Request.Form("txtProduct") & " series definitions updated in Pulsar" 
					oMessage.HTMLBody = "<font face=Arial size=3 color=black><b>Series Updated</b></font>" & strOutput
					oMessage.SendWithOutCopy
				End if
			End if
			Set oMessage = Nothing

            cn.BeginTrans
            set cm = server.CreateObject("ADODB.Command")
	        Set cm.ActiveConnection = cn
	        cm.CommandType = 4
	        cm.CommandText = "spUpdateSeriesSummary" 
	        
	        Set p = cm.CreateParameter("@ID", 3, &H0001)
	        p.Value = request.form("txtProductionVersionID")
	        cm.Parameters.Append p
            
	        Set p = cm.CreateParameter("@ID", 3, &H0001)
	        p.Value = mktBrandId
	        cm.Parameters.Append p
		        
	        cm.execute
            if cn.Errors.count <> 0 then
                cn.RollbackTrans
                sError = Err.Description
                response.write cn.Errors.description & "<BR>"
                On Error Goto 0
            else
                cn.CommitTrans
                strSuccess = "1"
		    end if  		
		end if

    End If   
	Set cm=nothing	
		
	cn.Close
	set cn = nothing	
    On Error Goto 0
	

	function FormatSystemID(strValue)
		dim RowArray
		dim Row
		dim strOutput
		dim IDArray
		
		if instr(strValue,"^")=0 and instr(strValue,"|")=0 then
			FormatSystemID = strValue
		else
			strOutput = ""
			RowArray = split(strValue,"|")
			for each Row in Rowarray
				if instr(Row,"^")=0 then
					strOutput = strOutput & ", " & Row
				else
					IDArray = split(Row,"^")
					strOutput = strOutput & ", " & IDArray(0)
					if Ubound(IDArray) > 0 then
						if trim(IDArray(1)) <> "" and trim(IDArray(1)) <> "&nbsp;" then
							strOutput = strOutput & "&nbsp;(" & replace(IDArray(1)," ","&nbsp;") & ")"
						end if
					end if
				end if
			next
			if strOutput = "" then
				FormatSystemID = "&nbsp;"
			else
				FormatSystemID = mid(strOutput,3)
			end if
			
		end if
	end function

%>
    <INPUT type="hidden" id="txtSuccess" name="txtSuccess" value="<%=strSuccess%>">
    <INPUT type="hidden" id="txtError" name="txtError" value="<%=sError%>"> 
</BODY>
</HTML>