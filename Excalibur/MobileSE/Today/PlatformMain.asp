<%@ Language=VBScript %>
<%	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"	  
%>	
<%
    Dim AppRoot
    AppRoot = Session("ApplicationRoot")
    dim nMode
    dim cn
	dim rs
	dim cm
	dim p
	dim strID
	dim CurrentUser
	dim CurrentUserID, CurrentUserName
    dim sIntroYear, sProductFamily, sProduct, sChassis, sBrand, sGenericName, sMarketingName, strBusinessSegmentID, sPCAGraphicsType, sPCAName, sMemoryOnboard, beMMC, sGraphicCapacity
    dim sFullMarketingName, sSOARID, sBiosName, sChinaGPIdentifier, sComments, sDeployment
    dim sSystemID,sSystemIdComments, sWebCycle, sSPQNeededDate, sStatus, sSpTimeCreated, sSpTimeUpdated, sSpCreator, sSpUpdater,sID
    dim sProductVersionID, strCategoryID, sPlatformID, sBusinessSegmentID, ArrayCycle
    dim sDisableBrand, bolHasActiveBU
    dim sFormFactorID, sSpindleID, sSpindle, sTouchID, sTouch, sTypeID, sType, sWidthUnit, sWidthValue, sDepthUnit, sDepthValue, sHeightFrontUnit, sHeightFrontValue, sHeightRearUnit, sHeightRearValue, sWeightUnit, sWeightValue, sNotes, sModelNumber
    dim bFName
    dim followMKTName : followMKTName = 0
    dim showMKTRows: showMKTRows ="none"
    dim showOthRows: showOthRows =""
    dim strBrandsAdded: strBrandsAdded =""
    dim strBrandNameLoaded: strBrandNameLoaded=""
    dim strDCRs : strDCRs=""
    dim showDCR : showDCR ="none"
    dim strSCMNumber : strSCMNumber =""
    dim strModelNumber : strModelNumber =""
    dim strScreenSize : strScreenSize =""
    dim strGeneration : strGeneration =""
    dim strFF : strFF ="" 
    dim strSeries: strSeries=""
    dim strMktName
    dim strPhwFamilyName: strPhwFamilyName=""
    dim strPMS: strPMS ="No Match" 
    dim showRequestPMSDiv: showRequestPMSDiv = "inline"
    dim tagSeriesName: tagSeriesName=""
    dim activeBtnSentPMReq:activeBtnSentPMReq = 0
    dim productVersion: productVersion =0
    dim productStatus:productStatus=0

    response.Flush
 
    sID = request("ID")
    sProduct = Request.Cookies("ProductName")
    sProductFamily = Request.Cookies("ProductFamily")
    sBusinessSegmentID = Request.Cookies("BusinessSegmentId")
    sProductVersionID = request("ProductVersionID")

    
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
	
	'Get User
	dim CurrentDomain
	dim CurrentUserPartner
	CurrentUser = lcase(Session("LoggedInUser"))

	if instr(currentuser,"\") > 0 then
		CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
		Currentuser = mid(currentuser,instr(currentuser,"\") + 1)
	end if

	set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	set rs = server.CreateObject("ADODB.recordset")

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

	set cm=nothing	
	

	CurrentUserID = 0
	if not (rs.EOF and rs.BOF) then
		CurrentUserID = rs("ID")
        CurrentUserName = rs("Name")
	end if
	rs.Close

    'get product
    set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	set rs = server.CreateObject("ADODB.recordset")
    rs.Open "Select ProductName, Pf.Name as ProductFamily, BusinessSegmentId, Version, pv.ProductStatusID From ProductVersion pv with (nolock)" &_
             "JOIN ProductFamily pf with (nolock) on pv.ProductFamilyId  = pf.ID Where pv.ID = " & sProductVersionID, cn, adOpenForwardOnly
    if not (rs.EOF and rs.BOF) then
		sProduct = rs("ProductName")
        sProductFamily = rs("ProductFamily")
        sBusinessSegmentID = rs("BusinessSegmentId")
        productVersion = rs("Version")
        productStatus = rs("ProductStatusID")
	end if
	rs.Close
    


    'check if follow marketing name
    If Not(IsNull(Request("FollowMktName"))) then
        followMKTName = Request("FollowMktName")
    end if
    IF followMKTName = 1 then
        showMKTRows = ""
        showOthRows = "none"
    end if

 if request("ID") = ""  then
	'Response.Write "<BR>&nbsp;Not enough information supplied"
    nMode = 1 'Create
    sPlatformID = "0"
    sFormFactorID = "0"

    if sBusinessSegmentID = 5 or sBusinessSegmentID = 6 or sBusinessSegmentID = 7 or sBusinessSegmentID = 8 or sBusinessSegmentID = 10 or sBusinessSegmentID = 11 then
     sChassis = 1578 ' default Notebooks chassis
    end if

 else
    nMode = 2 'update
    sPlatformID = request("ID")

    set cm = server.CreateObject("ADODB.Command")
	Set cm.ActiveConnection = cn
	set rs = server.CreateObject("ADODB.recordset")

	cm.CommandType = 4
	cm.CommandText = "usp_GetPlatformsandById "
	
	Set p = cm.CreateParameter("@p_chrID", 200, &H0001, 30)
	p.Value = sPlatformID
	cm.Parameters.Append p

    Set p = cm.CreateParameter("@pvID", 200, &H0001, 30)
	p.Value = sProductVersionID
	cm.Parameters.Append p

	rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly
	Set rs = cm.Execute 
	
	if not (rs.EOF and rs.BOF) then
     sPlatformID = rs("PlatformID") & ""
     sIntroYear = rs("IntroYear") & ""
     sProductFamily = rs("ProductFamily") & ""
     sPCAName = rs("PCA") & ""
     sPCAGraphicsType = rs("PCAGraphicsType") & ""
     sGraphicCapacity = rs("GraphicCapacity") & ""
     sMemoryOnboard = rs("MemoryOnboard") & ""
     beMMC = rs("eMMConboard") & ""
     sChassis = rs("ChassisID") & ""
     sGenericName = rs("GenericName") & ""
     sMarketingName  = rs("MktNameMaster") & ""
     sFullMarketingName   = rs("MktFullName") & ""
     sSOARID   = rs("SOAROID") & ""
     sBiosName   = rs("BIOSName") & ""
     sChinaGPIdentifier  = rs("ChinaGPIdentifier") & ""
     sComments   = rs("Comments") & ""
     sSystemID  = rs("SystemID") & ""
     sSystemIdComments  = rs("SystemIdComments") & ""
     sSPQNeededDate   = rs("SoftPaqNeedDate") & ""
     sStatus   = rs("StatusID") & ""
     sTimeCreated   = rs("TimeCreated") & ""
     sTimeUpdated   = rs("TimeChanged") & ""
     sCreator   = rs("Creator") & ""
     sUpdater   = rs("Updater") & ""
     sWebCycle = rs("WebCycle") & ""
     sTouchID = rs("TouchId") & ""
     sModelNumber = rs("ModelNumber") & ""
     sDeployment = rs("Deployment") & ""
     bFName = rs("BIOSFamilyName") &""
     strPhwFamilyName = rs("PHWebFamilyName") &""
     if (rs("MatchProductMaster")) then
        strPMS = "Match"
        showRequestPMSDiv ="none"
     end if
     if (rs("SentPMSReq") ) then
		activeBtnSentPMReq = 1
	 end if 
	end if

	rs.Close

    sBrand = 0
    rs.Open "Select ProductBrandID = isnull(ProductBrandID,0) From ProductVersion_Platform Where ProductVersionID = " & sProductVersionID & " and PlatformID = " & sPlatformID, cn, adOpenForwardOnly
	do while not rs.EOF        
        sBrand = rs("ProductBrandID")                              
        rs.MoveNext                  
	loop
	rs.Close 
    
    sDisableBrand = ""
    bolHasActiveBU = false
    rs.Open "Select COUNT(*) as Total " &_ 
			"from ProductVersion_Platform PP WITH (NOLOCK) " &_
			"INNER JOIN  PRL_Chassis Pl WITH (NOLOCK) on Pl.PlatformID = PP.PlatformID " &_
			"INNER JOIN	PRL_Products PRLP WITH (NOLOCK) on PRLP.ProductVersionID = PP.ProductVersionID " &_
			"INNER JOIN  PRL_Revision PR WITH (NOLOCK) ON PR.PRLRevisionID = PRLP.PRLRevisionID and PR.PRLID = Pl.PRLID " &_
			"Where PP.PlatformID = " & sPlatformID &_
			" and PP.ProductVersionID = " & sProductVersionID &_ 
			" and PR.StatusID in (Select StatusID from PRLStatus where Constant in ('MOL_ENG_TEST','MOL_POST_POR'))", cn
    if rs("Total") > 0 then
        sDisableBrand = "Disabled"                     
    else         
        rs.Close
        rs.Open "SELECT  count(1) as Total " &_
                "FROM	Feature FE INNER JOIN " &_
		        "Alias A ON FE.AliasID = A.AliasID INNER JOIN " &_
		        "Platform_Alias PA ON PA.AliasID = A.AliasID INNER JOIN " &_
		        "AvDetail AV ON AV.FeatureID = FE.FeatureID INNER JOIN " &_
		        "AvDetail_ProductBrand AVPB ON AVPB.AvDetailID = AV.AvDetailID " &_
                "WHERE	PA.PlatformID = " & sPlatformID & " AND ltrim(rtrim(A.Name)) <> '' AND AVPB.ProductBrandID = " & sBrand & " AND Status in ('A')", cn
        if rs("Total") > 0 then
            bolHasActiveBU = true
            sDisableBrand = "Disabled"
        else 
            bolHasActiveBU = false
        end if
        rs.Close 
	end if
    
    if (followMKTName = 1 And productStatus = 2 ) then '1:definision;2:development'
        showDCR =""
		set cm = server.CreateObject("ADODB.Command")
		Set cm.ActiveConnection = cn
		cm.CommandType = 4
		cm.CommandText = "spListApprovedDCRs"
		    
		Set p = cm.CreateParameter("@ID", 3, &H0001)
		p.Value = sProductVersionID
		cm.Parameters.Append p
	        
        set rs2 = server.CreateObject("ADODB.recordset")
		rs2.CursorType = adOpenForwardOnly
		rs2.LockType=AdLockReadOnly
		Set rs2 = cm.Execute 
		    
		strDCRs = "<Option selected></option>"
		do while not rs2.EOF
		    strDCRs = strDCRs & "<option value=""" & rs2("ID") & """>" & rs2("ID") & " - " & rs2("Summary") & "</option>"
		    rs2.MoveNext
		loop
		rs2.Close 	
    end if
    
    
    ' ************************* Get Form Factor Information **************************
    'Yong removed the code (as business function moved to PRL) instead of comment it out to make the page clean; 
    'if for some reason we need to put this back, we just need to look at tfs history
    ' ********************************************************************************
 end if

dim cycleList 
    cycleList = sWebCycle
dim retVal

function isChecked(cycle)
   retVal = ""
   if inStr(cycleList, cycle) > 0 then  
      retval = " checked "  
   end if
   isChecked = retval
end function	
%>

<HTML>
<HEAD>
    <TITLE>Choose Product Cycles</TITLE>
    <META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
    <link href="<%= AppRoot %>/style/wizard style.css" type="text/css" rel="stylesheet" />
    <link href="<%= AppRoot %>/style/Excalibur.css" type="text/css" rel="stylesheet" />
    <link href="<%= AppRoot %>/style/redmond/jquery-ui-1.8.7.custom.css" rel="stylesheet" type="text/css" />
    <link rel="stylesheet" type="text/css" href="<%= AppRoot %>/SupplyChain/style.css" />
    <link rel="stylesheet" type="text/css" href="<%= AppRoot %>/SupplyChain/sample.css" />
    <style type="text/css">
        .spinDiv {
                z-index: 99999;
            }
    </style>
</HEAD>

<script type="text/javascript" src="<%= AppRoot %>/_ScriptLibrary/jsrsClient.js"></script>
<script type="text/javascript" src="<%= AppRoot %>/includes/client/popup.js"></script>
<!-- #include file="../../includes/bundleConfig.inc" -->
<script src="<%= AppRoot %>/includes/client/jquery.blockUI.js" type="text/javascript"></script>

<script type="text/javascript" src="<%= AppRoot %>/includes/client/json2.js"></script>
<script src="/Pulsar/Scripts/spin/spin.js"></script>
<script src="/Pulsar/Scripts/spin/jquery.spin.js"></script>

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
    var preBrandID = 0;
    var curBrandID = 0;
    var preBrandName = "";
    var removedItem = 0;
    var sentMKTName = "";
    function cmdDate_onclick(FieldID) {
        var strID;
        var oldValue = window.frmMain.elements(FieldID).value;

        strID = window.showModalDialog("caldraw1.asp", FieldID, "dialogWidth:320px;dialogHeight:265px;edge: Raised;center:Yes; help: No;resizable: No;status: No");
        if (typeof (strID) == "undefined")
            return

        window.frmMain.elements(FieldID).value = strID;
    }

    function ddlBrand_onchange(ddl, OldBrandId, sDisableBrand, bolHasActiveBU) {
        if ((sDisableBrand == "Disabled" && bolHasActiveBU == "False") || bolHasActiveBU == "True") {
            var options = ddl.options;
            for (var i = 0; i < options.length; i++) {
                if (options[i].value == OldBrandId) {
                    options[i].selected = true;
                    break;
                }
            }

            if (sDisableBrand == "Disabled" && bolHasActiveBU == "False")
                alert("Once the PRL moves to POR and Post-POR status, you do not allow changing the Brand.");
            else if (bolHasActiveBU == "True")
                alert("Active Base Unit AVs are already set up in the SCM for this Base Unit Group and Brand.  Must set all Base Units to Obsolete before changing the Base Unit Group's Brand");
        }
    }

    function adjustWidth(percent) {
        return document.documentElement.offsetWidth * (percent / 100);
    }

    function adjustHeight(percent) {
        return (document.documentElement.offsetHeight * (percent / 100)) + 100;
    }

    function PlatformFormFactor_onclick(ProductVersionID, PlaformID) {
        var url = "../../../ipulsar/Product/Platform_FormFactor.aspx?ProductVersionID=" + ProductVersionID + "&PlatformID=" + PlaformID;

        var sWidth = 400;
        var sHeight = 650;

        window.parent.OpenDialog2(url, "Base Unit Group Form Factor", sWidth, sHeight, true, true, false, false);
    }

    //function ShowPlatformFormFactor(url, Title, DlgWidth, DlgHeight) {
    //    $("#iframeDialog").dialog({ width: DlgWidth, height: DlgHeight });

    //    $("#modalDialog").attr("width", "100%");
    //    $("#modalDialog").attr("height", "100%");
    //    $("#modalDialog").attr("src", url);
    //    $("#iframeDialog").dialog("option", "title", Title);
    //    $("#iframeDialog").dialog("open");
    //    $("#iframeDialog").attr("height", DlgHeight + "px");
    //}

    function ReloadIframe(selector) {
        $(selector).contents().find("body").html('');
    }

    function ReSubmit() {
        window.close();
        return false;
    }

       
    $(function () {
        $("#iframeDialog").dialog({
            modal: true,
            autoOpen: false,
            width: 900,
            height: 800,
            close: function () {
                $("#modalDialog").attr("src", "about:blank");
            },
            resizable: false

        });

        $('#iframeDialog').bind('dialogclose', function (event) {
            $("#iframeDialog").dialog('destroy');
            ReloadIframe("#modalDialog");
        });

         $(".ModelNumber").on("keyup", function (e) {

                var AllowedCharacters = /^[[0-9a-zA-Z \-]+$/;
                if ($(this).val() != "" && !($(this).val().match(AllowedCharacters))) 
                  {
                    alert("Model Number can only contain alphanumeric and dashes");
                    $(this).val("");
                    return;
                  }
                 // do not allow spaces
                $(this).val(this.value.replace(/^\s+|\s+$/g, ""));

         });

        if ($("#hdnFormFactorID").val() == "0")
            $("#divFormFactorFrm").hide();
        else
        {
            if ($('#ddlWeightUnit option:selected').index() == 1) {
                $('#lblounces').show();
                getOunces();
                $('#lblouncesvalue').show();
            }
            else {
                $('#lblounces').hide();
                $('#lblouncesvalue').hide();
            }
            
        }

        $("#txtWidthValue").keydown(function (e) {
            // Allow: backspace, delete, tab, escape, enter and .
            checkNumber(e);
        });

        $("#txtDepthValue").keydown(function (e) {
            // Allow: backspace, delete, tab, escape, enter and .
            checkNumber(e);
        });

        $("#txtHeightFrontValue").keydown(function (e) {
            // Allow: backspace, delete, tab, escape, enter and .
            checkNumber(e);
        });

        $("#txtHeightRearValue").keydown(function (e) {
            // Allow: backspace, delete, tab, escape, enter and .
            checkNumber(e);
        });

        $("#txtWeightValue").keydown(function (e) {
            // Allow: backspace, delete, tab, escape, enter and .
            checkNumber(e);
        });

        // on keyup
        $( "#txtWeightValue" ).keyup(function() {
            var num = $("#txtWeightValue").val();
            var ounces = (num * 0.035274).toFixed(2);
            $('#lblouncesvalue').html(ounces);
        });

        $("#btnReqPms").click(function () {
            RequestNewProductMasterSeries(); 
        });
            
        $("#cboBSBrandFilter").change(function () {
            getBrandWithBusinessSegment();
            getBrandFormulaWithBusinessSegment();
        });

        if($("input[name='chkBrands']:checked").val()!=null)
            curBrandID = $("input[name='chkBrands']:checked").val();

        $(".scmNo").change(function () {
            getFirstPhWebName($(this).val());
        });

        $(".SeriesChange").on('keyup', function () {
            if (this.id.indexOf(curBrandID) > -1 )
                brandColumnsChange(this);
        });
        $(".ModelChange").on('keyup', function () {
            if (this.id.indexOf(curBrandID) > -1) {
                brandColumnsChange(this);
            }
        });
        $(".ScreenChange").on('keyup', function () {
            if (this.id.indexOf(curBrandID) > -1)
                brandColumnsChange(this);
        });
        $(".GenChange").on('keyup',function () {
            if (this.id.indexOf(curBrandID) > -1)
                brandColumnsChange(this);
        });
        $(".FFChange").on('keyup', function () {
            if (this.id.indexOf(curBrandID) > -1)
                 brandColumnsChange(this);
        });
        if ($("#hdnFollowMKTName").val() == 1) {
            getBrandWithBusinessSegment();
            getBrandFormulaWithBusinessSegment();
            getFirstPhWebName();
            if ($("#hidPdStatus").val() > 2)
                $('#mktBrandsRow').find('*').attr('disabled', true);
        }
        	
		$(".SeriesChange").bind("mouseup", function(e){
            inputTextClear($(this));
		});
        $(".ModelChange").bind("mouseup", function(e){
            inputTextClear($(this));
		});
        $(".ScreenChange").bind("mouseup", function(e){
            inputTextClear($(this));
		});
        $(".GenChange").bind("mouseup", function(e){
            inputTextClear($(this));
		});
        $(".FFChange").bind("mouseup", function(e){
            inputTextClear($(this));
		});
		
        DisableEnableGraphicCapacity();
        modalDialog.load();
    });

   

    function checkNumber(e) {
        // Allow: backspace, delete, tab, escape, enter and .
        if ($.inArray(e.keyCode, [46, 8, 9, 27, 13, 110, 190]) !== -1 ||
            // Allow: Ctrl+A, Command+A
            (e.keyCode == 65 && (e.ctrlKey === true || e.metaKey === true)) ||
            // Allow: home, end, left, right, down, up
            (e.keyCode >= 35 && e.keyCode <= 40)) {
            // let it happen, don't do anything
            return;
        }
        // Ensure that it is a number and stop the keypress
        if ((e.shiftKey || (e.keyCode < 48 || e.keyCode > 57)) && (e.keyCode < 96 || e.keyCode > 105)) {
            e.preventDefault();
        }
    }


    function ddlWeightUnit_onchange() {
            if ($('#ddlWeightUnit option:selected').index() == 1) {
                $('#lblounces').show();
                getOunces();
                $('#lblouncesvalue').show();
            }
            else {
                $('#lblounces').hide();
                $('#lblouncesvalue').hide();
            }
    }

    function getOunces() {
        var num = $("#txtWeightValue").val();
        var ounces = (num * 0.035274).toFixed(2);
        $('#lblouncesvalue').html(ounces);
    }

    function ddlPCAGraphicsType_onchange()
    {
        DisableEnableGraphicCapacity();
    }

    function DisableEnableGraphicCapacity()
    {
        if ($('#ddlPCAGraphicsType option:selected').val() == "Discrete") {
            $("#ddlGraphicCapacity").prop("disabled", false);            
        }
        else {
            $("#ddlGraphicCapacity").prop("disabled", true);
            $("#ddlGraphicCapacity option:eq(0)").prop("selected", true);
        }
    }

    function getBrandWithBusinessSegment()
    {
       
        var brandTableRows = $("#TableBrand tbody tr");
        var businessSegmentId = $("#cboBSBrandFilter").val();

        if (businessSegmentId > 0) {
            var brands = $("#txtBrandsLoaded").val() + $("#txtBrandsAdded").val();
            if (brands[0] == ',')
                brands = brands.substring(1);

            $.ajax({
                url: "/Pulsar/Product/GetBrands?BusinessSegmentId=" + businessSegmentId,
                method: "GET",
                beforeSend: function () {
                    $('#mktBrandsRow').find('*').attr('disabled', true);
                    $("#tabDiv").spin("medium", "#0096D6");    
                },
                success: function (returnData) {
                    var data = jQuery.parseJSON(returnData);
                    var rows = $(brandTableRows).each(function (rowNum, val) {
                        var row = $(this);
                        var itemID;
                        var rowID = row[0].id;
                        var bdisplay;
                        rowID = rowID + ",";
                        bdisplay = false;
                        for (var i = 0; i <= data.brands.length - 1; i++) {

                            itemID = "Brand" + data.brands[i].ID + ",";
                            if (rowID.indexOf(itemID) > -1) {
                                bdisplay = true;
                                break;
                            }
                        }
                        if (!bdisplay && brands != "") {
                            $.each(brands.split(','), function (index, value) {
                                itemID = "Brand" + value + ",";
                                if (rowID.indexOf(itemID) > -1) {
                                    bdisplay = true;
                                }
                            });
                        }
                        if (bdisplay) {
                            $(row).show();
                        } else {
                            $(row).hide();
                        }

                    });
                },
                cache: false,
                complete: function () {
                    if($("#hidPdStatus").val() < 3)
                        $('#mktBrandsRow').find('*').removeAttr("disabled");
                    $("#tabDiv").spin(false); 
                },
                error: function (msg, status, error) {
                    alert("Please try again later. " + msg.responseText);
                }
            });

        }
        else {
            $.ajax({
                url: "/Pulsar/Product/GetBrands?BusinessSegmentId=0",
                method: "GET",
                success: function (returnData) {
                    var data = jQuery.parseJSON(returnData);
                    var rows = $(brandTableRows).each(function (rowNum, val) {
                        var row = $(this);
                        $(row).show();
                    });
                },
                cache: false
            });
        }
    }

    function BrandCheck_onclick(ID, PVID) {
        curBrandID = ID;
        if ($("#txtPlatformID").val() > 0) {
            preBrandID = $("#txtBrandsLoaded").val();
			preBrandName = $("#txtBrandNameLoaded").val();
        }
        if (preBrandID == 0) {
            preBrandID = ID;
            preBrandName = $("#txtBrandNameLoaded").val();
        }
        
        if ($("#txtPlatformID").val() == 0) {
            document.all("DivSeries" + preBrandID).style.display = "none";
            document.all("DivSeries" + ID).style.display = "";
            $("#txtBrandsLoaded").val(ID);
            $("#txtBrandsAdded").val(ID);
            preBrandID = ID;

            //TODO: change from new to old ==> brand1 sent --> brand2 --> brand 1 again
            if ($("#btnReqPms").is("[disabled=disabled]")) {
                $("#btnReqPms").removeAttr("disabled");
                $("#hidSentPMSReq").val(0);
            }
        }
        else
            RemoveOldBrand(preBrandID, preBrandName, ID, PVID);

        getNameByBrand(curBrandID);
    }

    function RemoveOldBrand(preID, preName, newID, PVID) {

          //if the brand has published SCM, do not allow user to delete the brand
         if (document.getElementById("txtSCMEnabled" + preID).value != "1") {
            alert("The brand has already been published in an SCM and cannot be deleted");
            event.srcElement.checked = true;
            return false;
         }

         Result = window.showModalDialog("BrandDeleteWarning.asp?ProductName='<%=sProduct %>'&BrandName=" + preName + "&BrandID=" + preID + "&PVID=" + PVID, "", "dialogWidth:700px;dialogHeight:300px;edge: Sunken;center:Yes; help: No;maximize:Yes;resizable: Yes;status: No")
         if (Result == "1") { //change to new product
             document.all("DivSeries" + preID).style.display = "none";
             document.all("DivSeries" + newID).style.display = "";
            removedItem = preID;
            preBrandID = newID;
            preBrandName = event.srcElement.BrandName;
             $("#txtBrandsAdded").val(newID);
         }
         else { 
            document.all("DivSeries" + preID).style.display = "";
            document.all("DivSeries" + newID).style.display = "none";
            $("#chkBrands" + newID).prop("checked", false)
            $("#chkBrands" + preID).prop("checked", true)
            preBrandID = preID;
            preBrandName = preName;
            $("#txtBrandsAdded").val(preID);
            curBrandID = preID;
        } 
    }

    function ChooseNewBrand(strID) { 
        var strIDs = "";
        var currentBSegmentID = $("#cboBSBrandFilter").val();
        if($("#txtBrandsLoaded").val()!="")
            strIDs = "," + $("#txtBrandsLoaded").val();
        if ($("#txtBrandsAdded").val() != "")
            strIDs = "," + $("#txtBrandsAdded").val();

        if (strIDs != "")
            strIDs = strIDs.substring(1);

        modalDialog.open({ dialogTitle: 'Brand', dialogURL: 'BrandUpdate_Pulsar.asp?ProductID=' + $("#txtProductionVersionID").val() + '&BrandID=' + strID + '&ExcludeIDList=' + strIDs + '&BSegmentID=' + currentBSegmentID + '', dialogHeight: 300, dialogWidth: 450, dialogResizable: false, dialogDraggable: true });
        globalVariable.save(strID, 'brand_id');
    }

    var BrandNSList, FamilyNSList;
    var BrandsColumns = [{ name: "Series", value: "txtSeriesA" },
                    { name:"ProductModelNumber", value: "txtModelNumber" },
                    { name:"Generation", value: "txtGeneration" },
                    { name:"ScreenSize", value: "txtScreenSize" },
                    { name: "FormFactor", value: "txtFormFactor" }];

    function getBrandFormulaWithBusinessSegment() {
        var busId = $("#cboBSBrandFilter").val(); 
        $.ajax({
            type: "POST",
            contentType: "application/json; charset=utf-8",
            url: "/Pulsar/Product/GetBrandFormulaNameByBusinessSegment",
            data: JSON.stringify({ bsId: busId }), 
            dataType: "json",
            beforeSend: function () {
                $('#mktBrandsRow').find('*').attr('disabled', true);
                $("#tabDiv2").spin("medium", "#0096D6");
            },
            success: function (returnData) {
                if (returnData == null || returnData.length == 0)
                    alert("can not get the brand formula. please make sure you have naming standard for 'Long Name'.")
                else {
                    BrandNSList = returnData.BrandName;
                    FamilyNSList = returnData.FamilyName;
                    if ($("#txtPlatformID").val() > 0 && curBrandID > 0) 
                        getNameByBrand(curBrandID);
                }
            },
            complete: function () {
                if($("#hidPdStatus").val() < 3)
                    $('#mktBrandsRow').find('*').removeAttr("disabled");
                $("#tabDiv2").spin(false);
            },
            error: function (msg, status, error) {
                alert("Please try again later. " + msg.responseText);
            }
        });
    }

    function brandColumnsChange(element) {   
        setTimeout(function () {
            getNameByBrand(curBrandID);
        }, 1000); 
    }

    function getNameByBrand(brandId) {
        getMKTName(brandId, BrandNSList);
        getFamilyName(brandId, FamilyNSList);
        CheckProductMaster();
    }

    function getMKTName(brandId, brandFormulaList) {
        var brandMktName = getBrandFormulaName(brandId, brandFormulaList);
        if (brandMktName.result == "") {
            $("#divMkt").html("");
            $("#hidMktName").val("");
            alert("The marketing name composed of following columns: " + brandMktName.formular + ". Please enter at least one of these columns."); 
            return false;
        }
        else {
            $("#divMkt").html(brandMktName.result);
            $("#hidMktName").val(brandMktName.result);
            
            if ($("#hidMktName").val() != $("#tagMktName").val() && $("#btnReqPms").is("[disabled=disabled]")) {
                $("#btnReqPms").removeAttr("disabled");
                $("#hidSentPMSReq").val(0);
            }
            if (brandMktName.result.length > 60) {
                alert("Marketing Name cannot exceed 60 characters. Please try to change the following column: " + brandMktName.formular);
                return;
            }
        }
    }

    function getFamilyName(brandId, brandFormulaList) {
       
        var brandFamilyName = getBrandFormulaName(brandId, brandFormulaList);

        brandFamilyName.result = brandFamilyName.result.replace("ProductName", $("#txtProduct").val());
        if ($("#hidVersion").val() != "")
            brandFamilyName.result = brandFamilyName.result.replace("ProductVersion", $("#hidVersion").val())
        else if (brandFamilyName.result.indexOf("ProductVersion ") > -1)
            brandFamilyName.result = brandFamilyName.result.replace("ProductVersion ", "")
        else
            brandFamilyName.result = brandFamilyName.result.replace("ProductVersion", "")    
        brandFamilyName.result = $.trim(brandFamilyName.result);
        if (brandFamilyName.result == "") {
            alert("The Family name composed of following columns: " + brandFamilyName.formular + ". Please enter at least one of these columns."); 
            $("#divPhwFamily").html("");
            $("#hidphwFamilyName").val("");
            return false;
        }
        else if ($('#hidfirstPfID').val() > 0 && ($('#hidfirstPfID').val() != $('#txtPlatformID').val())) {
            $("#divPhwFamily").html($('#hidfirstPwfName').val()); 
            $("#hidphwFamilyName").val(brandFamilyName.result);
        }
        else {
            $("#divPhwFamily").html(brandFamilyName.result); 
            $("#hidphwFamilyName").val(brandFamilyName.result);
        }
    }

    function getBrandFormulaName(brandId, brandFormulaList) {
         var brandNameFormula;
         var brandName;
         var screenSizeUnit;
    
         for (var i = 0; i < brandFormulaList.length; i++) {
             if (brandFormulaList[i].BrandId == brandId) {
                 brandNameFormula = brandFormulaList[i].FormulaName;
                 screenSizeUnit = brandFormulaList[i].ScreenSizeUnit;
                 break;
             }
         }
    
         brandName = $.trim(brandNameFormula);
         if (brandName.indexOf("ScreenSize") > -1 && $('#txtScreenSize' + brandId).val() == "") {
             brandName= brandName.replace(new RegExp(screenSizeUnit, "g"), "");
         }
         for (var j = 0; j < BrandsColumns.length; j++) {
             if (brandName.indexOf(BrandsColumns[j].name) > -1 && $('#' + BrandsColumns[j].value +brandId ).val() !="") {
                     brandName = brandName.replace(BrandsColumns[j].name, $('#' + BrandsColumns[j].value +brandId ).val());
             }
             else if(brandName.indexOf(BrandsColumns[j].name+ " ") > -1)
                 brandName = $.trim(brandName.replace(BrandsColumns[j].name+" " , ""));
             else
                 brandName = $.trim(brandName.replace(BrandsColumns[j].name , ""));
         }
         return { formular:brandNameFormula, result : $.trim(brandName) };
    }


    function CheckProductMaster() {
        var matched = 0;
        var productversionID = $("#txtProductionVersionID").val();
        $.ajax({
            type: "POST",
            contentType: "application/json; charset=utf-8",
            url: "/Pulsar/Product/GetProductMastrSeries",
            data: JSON.stringify({ pvId: productversionID }),
            dataType: "json",
            beforeSend: function () {   
                $('#mktBrandsRow').find('*').attr('disabled', true);
                $("#tabDiv").spin("medium", "#0096D6");
            },
            success: function (returnData) {
                if (returnData != null) {
                    for (var i = 0;i<returnData.length;i++ ){
						if (returnData[i] == $('#divMkt').html()){
							$("#divPms").html("Match");
							$("#hidPMS").val("Match");
							$("#requestPMSDiv").hide();
							$("#btnReqPms").attr("disabled", true); 
							matched =1;
							break;
						}
					}
                }
					
                if (matched==0){
                    $("#divPms").html("No Match");
					$("#hidPMS").val("No Match");
                    $("#requestPMSDiv").show();
					if ($('#hidSentPMSReq').val() == 0)
						$("#btnReqPms").removeAttr("disabled");
                }                
            },
            complete: function () {
                if($("#hidPdStatus").val() < 3)
                    $('#mktBrandsRow').find('*').removeAttr("disabled");
                $("#tabDiv").spin(false);
            },
            error: function (msg, status, error) {
                alert("Please try again later. " + msg.responseText);
            }
        });
    }

    function RequestNewProductMasterSeries() {
        var curMktName = $('#divMkt').html();
        $("#btnReqPms").attr("disabled", true);
        var productVersionID = $('#txtProductionVersionID').val();
        var curUserID = $("#txtCurrentUserID").val(); 
        var brandId = curBrandID;
        var sId = ($("#txtSeriesIDA" + brandId).val()=="")? 0 : $("#txtSeriesIDA" + brandId).val();
        var oldMktName, tagMktName
        tagMktName = $("#tagMktName").val();
       
        if (curMktName != tagMktName) {
            oldMktName = tagMktName;
        } else if ($('#sentMKTName').val()!="" && $('#sentMKTName').val() !=curMktName) {
            oldMktName = $('#sentMKTName').val();
        }
        $.ajax({
            type: "POST",
            contentType: "application/json; charset=utf-8",
            url: "/Pulsar/Product/RequestNewProductMasterSeries",
            data: JSON.stringify({pvId:productVersionID, bId:brandId, seriesId:sId, userId:curUserID, curMktName:curMktName, oldMktName:oldMktName }),
            dataType: "json",
            beforeSend: function () {
                $("#tabDiv").spin("medium", "#0096D6");
            },
            success: function (returnData) {
                if (returnData.ReturnValue == 0) {
                    $('#hidSentPMSReq').val(1);
                    $('#sentMKTName').val(curMktName);
                    alert("Your request has been sent."); 
                }
                else {
                    alert(returnData.Message);
                }
            },
            complete: function () {
                $("#tabDiv").spin(false);
            },
            error: function (msg, status, error) {
                alert("Please try again later. " + msg.responseText);
            }
        });
    }

    function getFirstPhWebName(selectedScmNo) {
        var pvId = $('#txtProductionVersionID').val();
        if (selectedScmNo == null && curBrandID == 0) {
            selectedScmNo = 0;
        }
         $.ajax({
            type: "POST",
            contentType: "application/json; charset=utf-8",
            url: "/Pulsar/Product/GetFirstPhWebName",
            data: JSON.stringify({productVersionID:pvId, scmNo: selectedScmNo}),
            dataType: "json",
            cache:false,
            beforeSend: function () {
                $("#tabDiv").spin("medium", "#0096D6");
            },
            success: function (returnData) {
                if (returnData.ReturnValue == 0 && returnData.phwebNames!= null) {
                    $("#hidfirstPfID").val(returnData.phwebNames.PlatformId);
                    $("#hidfirstPwfName").val(returnData.phwebNames.PhWebFamilyName);
                    if ($("#txtPlatformID").val() > 0 || curBrandID > 0) {
                        $("#divPhwFamily").html($('#hidfirstPwfName').val()); 
                    }
                } else if (returnData.ReturnValue == 0 && returnData.phwebNames == null) {
                    $("#txtPlatformID").val() == 0 ? $("#hidfirstPfID").val(0) : $("#hidfirstPfID").val($("#txtPlatformID").val());
                    getFamilyName(curBrandID, FamilyNSList);
                } else if (returnData.ReturnValue == 1){
                    alert(returnData.Message);
                }
            },
            complete: function () {
                $("#tabDiv").spin(false);
            },
            error: function (msg, status, error) {
                alert("Please try again later. " + msg.responseText);
            }
        });
    }

    var KeyString = "";
    function combo_onkeypress() {
	    if (event.keyCode == 13){
		    KeyString = "";
		}
	    else{
		    KeyString=KeyString+ String.fromCharCode(event.keyCode);
		    event.keyCode = 0;
		    var i;
		    var regularexpression;
		
		    for (i=0;i<event.srcElement.length;i++){
				    regularexpression = new RegExp("^" + KeyString,"i")
				    if (regularexpression.exec(event.srcElement.options[i].text)!=null) 
					    event.srcElement.selectedIndex = i;
			}
		    return false;
		}	
    }

    function combo_onfocus() {
	    KeyString = "";
    }

    function combo_onclick() {
	    KeyString = "";
    }

    function combo_onkeydown() {
	    if (event.keyCode==8){
		    if (String(KeyString).length >0)
			    KeyString= Left(KeyString,String(KeyString).length-1);
		    return false;
		}
    }

     function ChoosenewBrandResult() {
        var strID;
        var strResult;

        strResult = modalDialog.getArgument('brand_update_result');
        strID = globalVariable.get('brand_id');

        if (typeof (strResult) != "undefined") {

            for (i = 0; i < frmMain.chkBrands.length; i++) {

                if (frmMain.chkBrands[i].value == strID) {
                    frmMain.chkBrands[i].checked = false;
                    document.all("Brand" + strID).style.display = "none";
                    document.all("DivSeries" + strID).style.display = "none";
                }

                if (frmMain.chkBrands[i].value == strResult) {
                    frmMain.chkBrands[i].checked = true;
                    document.all("DivSeries" + strResult).style.display = "";
                }
            }

            $("#txtBrandsAdded").val(strResult) ;

            document.all("txtSeriesA" + strResult).value = document.all("txtSeriesA" + strID).value;

            document.getElementById('DIV3').scrollTop = document.all("Brand" + strResult).offsetTop - 35; 
            document.all("txtBrandFrom").value = strID;
            document.all("txtBrandTo").value = strResult;
            curBrandID = strResult;
        }
    }
	
	function inputTextClear(element){
        var $input = $(element),
        oldValue = $input.val();
        var newValue;
        
        setTimeout(function(){
            newValue = $input.val();  
            if (oldValue != newValue && newValue=='' && $input.attr('id').indexOf(curBrandID) > -1)
                brandColumnsChange($input);      
        }, 1);
    }
    //-->
</SCRIPT>

<BODY bgcolor=ivory>

<link rel="stylesheet" type="text/css" href="../Style/programoffice.css">
<font size=3 face=verdana><b>
    <%   
    if request("ID") =  "" then
        response.write "Add Base Unit Group"
    else
        response.write "Update Base Unit Group"
    end if
   %>
</b></font>

<div id="tabDiv" class="spinDiv">
<form ID=frmMain method=post action="PlatformSave.asp?pulsarplusDivId=<%=request("pulsarplusDivId")%>">
<div id="tabDiv2" class="spinDiv">
<table ID="tabAdd" WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
   <!-- Hide Baseunit groupID field task 24231-->
    <!--<% if nMode=2 then %>
    <tr>
        <td width="160" nowrap valign=top><b>Base Unit Group ID</b>&nbsp;</td>
        <td width="100%"><%=Server.HTMLEncode(sPlatformID)%></td>
    </tr>
    <% end if %>-->
    
	<tr>
		<td width="160" nowrap valign=top><b>Intro Year:</b><font color="red" size="1"> *</font></td>
		<td width="100%"><input type="text" id="txtIntroYear"  name="txtIntroYear" value="<%=Server.HTMLEncode(sIntroYear)%>" maxlength="4" /></td>
    </tr>
    <tr>
		<td width="160" nowrap valign=top><b>Product Family:</b></td>
		<td width="100%"><%=Server.HTMLEncode(sProductFamily)%></td>
    </tr>
    <tr>
		<td width="160" nowrap valign=top><b>Product:</b></td>
		<td width="100%"><%=Server.HTMLEncode(sProduct)%></td>
    </tr>
    <tr id="brandsRow" style="display:<%=showOthRows%>">
		<td width="160" nowrap valign=top><b>Brand:</b><font color="red" size="1"> *</font></td>
		<td width="100%">
            <table><tr><td style="vertical-align:top; padding:0">
            <select id="ddlBrand" name="ddlBrand" onchange="ddlBrand_onchange(this, '<%= trim(sBrand) %>', '<%= sDisableBrand %>', '<%= bolHasActiveBU %>')">
            <option selected value="" >Select Brand to assign Base Unit Group to SCM</option>
              <%         ' get brands
                dim strSQL
                set cm = server.CreateObject("ADODB.Command")
	            Set cm.ActiveConnection = cn
                set rs = server.CreateObject("ADODB.recordset")
                rs.open "usp_GetProductCombinedSCMs " & clng(sproductVersionID), cn, adOpenForwardOnly
              while Not rs.EOF
              if trim(rs("ProductBrandID")) = trim(sBrand) then
            %>
                 <option selected value="<%=rs("ProductBrandID")%>"><%=rs("Name")%></option>
                        <%else%>
                            <option value="<%=rs("ProductBrandID")%>"><%=rs("Name")%></option>
                        <%end if%>
        <%
           rs.MoveNext
           Wend
           rs.Close

           %>
           </select>
           </td><td style="vertical-align:top">
            <font color="green" size="1">&nbsp;If brands are not displayed, go to General tab and select/save a brand then try again.</font>
               </td></tr></table>
		   </td>
    </tr>
    <!--Brand Start -->
    <tr id="mktBrandsRow" style="display:<%=showMKTRows%>">
        <td>
            <strong>Brands:</strong><p style="color:red;font-size:10px;display:inline;"> *</p>
               <p>Use SCM column to combine multiple brands into 1 SCM. Blanks create a unique SCM. </p>
               <i>Note: Brands must share the same Software and Product Drop to be combined into 1 SCM.</i>
        </td>
        <td >
            <%
                response.write "<input id=""txtBrandFrom"" name=""txtBrandFrom"" type=""hidden"" value"""">"
                response.write "<input id=""txtBrandTo"" name=""txtBrandTo"" type=""hidden"" value"""">"
            %>
            <br />
            <!--DCR start-->
            <div id="divDCR" style="Display:<%=showDCR%>">
                <b>Approved DCR:  </b><p style="color:red;font-size:10px;display:inline;"> *</p> <SELECT style="WIDTH:100%" id=cboDCR name=cboDCR onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()"><%=strDCRs%></SELECT>
            </div>
            <!--DCR END-->
            <br />
            Filter Brands by:&nbsp;&nbsp;
                <select id="cboBSBrandFilter" name="cboBSBrandFilter" style="width: 200px;">
                    <option selected value=""></option>
                        <%  
                        rs.open "spPULSAR_Product_ListBusinessSegments",cn 
                        do while not rs.eof
                            if trim(rs("BusinessSegmentID")) = trim(sBusinessSegmentID) then%>
                                <option selected value="<%=rs("businesssegmentID")%>"><%=rs("name")%></option>
                        <%else%>
                            <option value="<%=rs("businesssegmentID")%>"><%=rs("name")%></option>
                        <%end if%>
                    <%
                            rs.movenext
                        loop
                        rs.close  
                    %>
                </select><br /><br />
                    <div style="border-right: steelblue 1px solid; border-top: steelblue 1px solid; overflow-y: scroll;
                        border-left: steelblue 1px solid; border-bottom: steelblue 1px solid; height: 160px; width:700px; 
                        background-color: white" id="DIV3">
                        <table width="900px" id="TableBrand">
                            <thead>
                                <tr style="position: relative; top: expression(document.getElementById('DIV3').scrollTop-2);">
                                    <td width="10" style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;border-bottom: 1px outset" bgcolor="#c9ddff">
                                        &nbsp;
                                    </td>
                                    <td width="70" style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;border-bottom: 1px outset" bgcolor="#c9ddff">
                                        &nbsp;Brand
                                    </td>
                                    <td width="70" style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;border-bottom: 1px outset" bgcolor="#c9ddff">
                                        &nbsp;SCMs
                                    </td>
                                    <td width="220" style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;border-bottom: 1px outset" bgcolor="#c9ddff">
                                        &nbsp;Series
                                    </td>
                                    <td width="90" style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;border-bottom: 1px outset" bgcolor="#c9ddff">
                                        &nbsp;Model Number
                                    </td>
                                     <td width="40" style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;border-bottom: 1px outset" bgcolor="#c9ddff">
                                        &nbsp;Screen Size
                                    </td>
                                     <td width="40" style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;border-bottom: 1px outset" bgcolor="#c9ddff">
                                        &nbsp;Generation
                                    </td>
                                     <td width="70" style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset; border-bottom: 1px outset" bgcolor="#c9ddff">
                                        &nbsp;Form Factor
                                    </td>
                                    <td width="40" style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;border-bottom: 1px outset" bgcolor="#c9ddff">
                                        &nbsp;Suffix&nbsp;
                                    </td>
                                </tr>
                            </thead>
                            <%
		                    if sPlatformID > 0  then
			                    rs.open  "spListBrands4Product_Pulsar " & clng(sProductVersionID) & ", 0," & clng(sPlatformID)  ,cn,adOpenForwardOnly '
		                    else
			                    rs.open  "spListBrands4Product_Pulsar " & clng(sProductVersionID) & ", 0, 0" ,cn,adOpenForwardOnly
		                    end if
		
		                    dim strRow
		                    dim strRowWait
		                    strRowWait = ""
                            dim strDisableBrandInfo
                            dim SCMEnabled
		                    do while not rs.eof
    		                    strRow = ""
			                    if rs("Active") or not isnull(rs("ProductBrandID")) then
				                    if rs("selectedBrand") = 0  then 
					                    strChecked = ""
				                    else
					                    strChecked = "checked"
					                    strBrandsLoaded = trim(rs("ID"))
                                        strBrandsAdded = trim(rs("ID"))
                                        strBrandNameLoaded = trim(rs("Name"))
                                        strSCMNumber = rs("SCMNumber")
                                        strModelNumber = rs("ProductModelNumber")
                                        strScreenSize = rs("ScreenSize")
                                        strGeneration = rs("Generation")
                                        strFF = rs("FormFactor")
                                    end if
                                    
				                    if isnull(rs("LastPublishDt")) then
					                    strDisableBrandInfo = ""
                                        SCMEnabled ="1"
				                    else
					                    strDisableBrandInfo = " disabled "
                                        SCMEnabled ="0"
				                    end if 
				                    strAbbr = rs("name") & "" 
                                    if sPlatformID = 0 then
                                        strRow = strRow &  "<TR ID=Brand" & trim(rs("ID")) & "><TD nowrap><INPUT type=""radio"" " & strChecked & " BrandName=""" & rs("Name") & """ BNWOFormula=""" & rs("BrandsWOFormula") & """ title=" & rs("ID") & " id=""chkBrands" & trim(rs("ID")) & """ name=chkBrands LANGUAGE=javascript onclick=""return BrandCheck_onclick(" & rs("ID") & "," & "0" & ")"" value=""" & rs("ID") & """></TD>"
                                    else
		                                strRow = strRow &  "<TR ID=Brand" & trim(rs("ID")) & "><TD nowrap><INPUT type=""radio"" " & strChecked & " BrandName=""" & rs("Name") & """ BNWOFormula=""" & rs("BrandsWOFormula") & """ title=" & rs("ID") & " id=""chkBrands" & trim(rs("ID")) & """ name=chkBrands LANGUAGE=javascript onclick=""return BrandCheck_onclick(" & rs("ID") & "," & request("ID") & ")"" value=""" & rs("ID") & """></TD>"
                                    end if
                                    if strChecked <> "" then 
                                        if rs("Active") then
                                            strRow = strRow &  "<TD nowrap><font face=verdana size=2><a href=""javascript: ChooseNewBrand(" & rs("ID")& ");"">" & strAbbr & "</a>&nbsp;</font></TD>"
                                        else
                                            strRow = strRow &  "<TD nowrap><font face=verdana size=2><a href=""javascript: ChooseNewBrand(" & rs("ID")& ");"">" & strAbbr & "&nbsp;(old)</a></font></TD>"
                                        end if
                                    else
                                        if rs("Active") then
                                            strRow = strRow &  "<TD nowrap><font face=verdana size=2>" & strAbbr & "&nbsp;</font></TD>"
                                        else
                                            strRow = strRow &  "<TD nowrap><font face=verdana size=2>" & strAbbr & "&nbsp;(old)</font></TD>"
                                        end if
                                    end if
                                       
                                    dim iSCMNum
                                    iSCMNum=0
                                    dim iSelectedSCMNumber
                                    iSelectedSCMNumber =clng(rs("SCMNumber"))
                
                                    strRow = strRow &  "<TD nowrap><select id=SelectSCM" & trim(rs("ID")) & " name=SelectSCM" & trim(rs("ID")) & " style='width: 70px;' class='scmNo'" 
                                    strRow = strRow & strDisableBrandInfo
                                    strRow = strRow & ">"
                                    strRow = strRow & " <option "
                   
                                    strRow = strRow & " value='0'></option>"
               
                                    iSCMNum=1
                                    do while  iSCMNum<21
				                        strRow = strRow & "  <option value='" & cstr(iSCMNum) & "'"
                                        if iSCMNum = iSelectedSCMNumber then
                                            strRow = strRow & " selected "
                                        end if  
                                        strRow = strRow &             " >SCM " & cstr(iSCMNum) &  "</option>"
			                            iSCMNum = iSCMNum +1 
                                    loop    
                                    strRow = strRow &     "</select>"
                                    strRow = strRow &     "<INPUT type=""text"" id=txtSelectedSCM" & trim(rs("ID")) & " name=txtSelectedSCM" & trim(rs("ID")) & " style=""Display:none"" value=""" & iSelectedSCMNumber & """>"
                                    strRow = strRow &     "<INPUT type=""text"" id=txtSCMEnabled" & trim(rs("ID")) & " name=txtSCMEnabled" & trim(rs("ID")) & " style=""Display:none"" value=""" & SCMEnabled & """>"
                                    strRow = strRow &     "</TD>"  
                                    'end of SCM number
                                    strRow = strRow &  "<TD><INPUT type=""text"" id=tagSeries" & trim(rs("ID")) & "  name=tagSeries" & trim(rs("ID")) & "  style=""Display:none"" value=""" & rs("SeriesSummary") & """>"
				
				                    if strChecked = "" then
					                    showSeries = "none"
				                    else
					                    showSeries = ""
				                    end if
				                    'Pull series
				                    strSeriesID = ""
				                    strSeriesName = ""
				                    if  not isnull(rs("ProductBrandID")) and (rs("selectedBrand") > 0) then 
					                    set rs2 = server.CreateObject("ADODB.recordset")
					                    rs2.Open "spListSeries4BrandAndPlatform " &  rs("ProductBrandID") & ", " & sPlatformID ,cn,adOpenForwardOnly 
					                    strSeriesID = ""
					                    strSeriesName = ""
                                        if not (rs2.EOF and rs2.BOF) then 
                                            strSeriesID = rs2("ID")
                                            strSeriesName = rs2("Name") 
                                            tagSeriesName = strSeriesName
					                     end if   
				                    end if

				                    strRow = strRow &  "<DIV style=""display:" & showSeries & """ id=DivSeries" & trim(rs("ID")) & ">"
                                    if (rs("selectedBrand") = 0) then
                                       strSeriesID="0"
                                    end if
				                    strRow = strRow &  "<INPUT style=""width:100"" type=""text"" class=""SeriesChange"" id=txtSeriesA" & trim(rs("ID")) & " name=txtSeriesA" & trim(rs("ID")) & " value=""" & strSeriesName & """"                          
                                    strRow = strRow &  ">"
				                    strRow = strRow &  "<INPUT type=""hidden"" id=tagSeriesA" & trim(rs("ID")) & " name=tagSeriesA" & trim(rs("ID")) & " value=""" & strSeriesName & """>"
				                    strRow = strRow &  "<INPUT type=""hidden"" id=txtSeriesIDA" & trim(rs("ID")) & " name=txtSeriesIDA" & trim(rs("ID")) & " value=""" & strSeriesID & """>"

				                    strRow = strRow &  "</DIV></TD>"
                
                                    
                                    'Model number cells
                                    strRow = strRow &  "<TD nowrap>" & "<INPUT style=""width:90""  maxlength=""10"" type=""text"" class=""ModelChange"" id=txtModelNumber" & trim(rs("ID")) & " name=txtModelNumber" & trim(rs("ID")) & " value=""" & rs("ProductModelNumber") & """>" & "</TD>"

                                    'Screen size cells
                                    strRow = strRow &  "<TD nowrap>" & "<INPUT style=""width:90"" maxlength=""10"" type=""text"" class=""ScreenChange"" id=txtScreenSize" & trim(rs("ID")) & " name=txtScreenSize" & trim(rs("ID")) & " value=""" & rs("ScreenSize") & """>" & "</TD>"

                                    ' generation and form factor cells
                                    strRow = strRow &  "<TD nowrap>" & "<INPUT style=""width:40"" maxlength=""5"" type=""text"" class=""GenChange"" id=txtGeneration" & trim(rs("ID")) & " name=txtGeneration" & trim(rs("ID")) & " value=""" & rs("Generation") & """>" & "</TD>"

                                    strRow = strRow &  "<TD nowrap>" & "<INPUT style=""width:70"" type=""text"" class=""FFChange"" id=txtFormFactor" & trim(rs("ID")) & " name=txtFormFactor" & trim(rs("ID")) & " value=""" & rs("FormFactor") & """>" & "</TD>"
                                 
            

				                    if len(rs("Suffix") & "") > 13 then
                                        strRow = strRow &  "<TD nowrap><font face=verdana size=2>" & left(rs("Suffix"),13) & "...&nbsp;</font><div id=suffix" & trim(rs("ID")) &" style=""display:none;"">"& rs("Suffix") &"</div></TD></TR>"
                                    else
				                        strRow = strRow &  "<TD nowrap><font face=verdana size=2>" & rs("Suffix") & "&nbsp;</font><div id=suffix" & trim(rs("ID")) &" style=""display:none;"">"& rs("Suffix") &"</div></TD></TR>"
                                    end if

				                    if strChecked = "checked" then
				                        Response.Write strRow
				                    else
				                        strRowWait = strRowWait & strRow
				                    end if
			                    end if
			                    rs.MoveNext
		                    loop
		                    rs.Close
                            set rs = Nothing
		                    if strRowWait <> "" then
		                        response.write strRowWait
		                    end if
                            %>
                        </table>
                    </div>
        </td>
    </tr>  
    <!--Brand End -->
    <tr id="mktRow" style="display:<%=showMKTRows%>">
        <td><strong>Marketing Name:</strong></td>
        <td id="tdMkt">
            <div id="divMkt"><%=Server.HTMLEncode(sMarketingName)%></div>
        </td>
    </tr>
    <tr id="prodMasterRow" style="display:<%=showMKTRows%>">
        <td><strong>Product Master Series:</strong></td>
        <td id="tdPms">
            <table>
                <tr>
                    <td><div id="divPms"><%=Server.HTMLEncode(strPMS)%></div></td>
                    <td>
                        <div id="requestPMSDiv" style="display:<%=showRequestPMSDiv%>;padding-left:100px;">
                            <input type="button" id="btnReqPms" value="Send" <%if strPMS="Match" or sPlatformID = 0 or activeBtnSentPMReq = 1 then%> disabled <% end if %>/> Request New Product Master Series 
                        </div>
                    </td> 
                </tr>
            </table>
        </td>
    </tr>
    <tr id="phWebFamilyRow" style="display:<%=showMKTRows%>">
        <td><strong>PHWeb Family Name:</strong></td> 
        <td id="tdPhwFamily"><div id="divPhwFamily" style="display:inline;"><%=Server.HTMLEncode(strPhwFamilyName)%></div></td>
    </tr>
    <tr>
        <td width="160" nowrap valign=top><b>Product Model Number:</b></td>
        <td><input type="text" id="txtProductModelNumber" name="txtProductModelNumber" value="<%=Server.HTMLEncode(sModelNumber)%>" maxlength="10" size="10" /></td>
    </tr>
    <tr>
        <td width="160" nowrap valign=top><b>System Board:</b><font color="red" size="1"> *</font></td>
        <td><input type="text" id="txtPCAName" name="txtPCAName" value="<%=Server.HTMLEncode(sPCAName)%>" maxlength="30" size="75" /></td>
    </tr>

    <tr>
        <td style="width:160px; text-wrap:none; vertical-align:top;"><b>eMMC onboard:</b></td>
        <td>
            <% if beMMC then %>
                <input type="checkbox" name="ckeMMC" ID="ckeMMC" checked />
            <% else %>
                <input type="checkbox" name="ckeMMC" ID="ckeMMC" />
            <% end if %>
        </td>
    </tr>

    <tr>
        <td style="width:160px; text-wrap:none; vertical-align:top;"><b>Memory onboard:</b></td>
        <td style="width:100%">
            <select id="ddlMemoryOnboard" name="ddlMemoryOnboard">
                <option value=""></option>
                <% if sMemoryOnboard = "2GB" then %>
                    <option selected value="2GB">2GB</option>
                    <option value="4GB">4GB</option>
                <% elseif sMemoryOnboard = "4GB" then %>
                    <option value="2GB">2GB</option>
                    <option selected value="4GB">4GB</option>
                <% else %>
                    <option value="2GB">2GB</option>
                    <option value="4GB">4GB</option>
                <% end if %>                   
	      </select>
        </td>
    </tr>

    <tr>
		<td width="160" nowrap valign=top><b>PCA Graphic Type:</b></td>
		<td width="100%">
            <select id="ddlPCAGraphicsType" name="ddlPCAGraphicsType" onchange="ddlPCAGraphicsType_onchange();">
                <option value=""></option>
                <% if sPCAGraphicsType = "UMA" then %>
                    <option selected value="UMA">UMA</option>
                    <option value="Discrete">Discrete</option>
                <% elseif sPCAGraphicsType = "Discrete" then %>
                    <option value="UMA">UMA</option>
                    <option selected value="Discrete">Discrete</option>
                <% else %>
                    <option value="UMA">UMA</option>
                    <option value="Discrete">Discrete</option>
                <% end if %>                   
	      </select></td>
    </tr>  

    <tr>
        <td style="width:160px; text-wrap:none; vertical-align:top;"><b>Graphic Capacity:</b></td>
        <td style="width:100%">
            <select id="ddlGraphicCapacity" name="ddlGraphicCapacity">
                <option value=""></option>
                <% if sGraphicCapacity = "2GB" then %>
                    <option selected value="2GB">2GB</option>
                    <option value="4GB">4GB</option>
                <% elseif sGraphicCapacity = "4GB" then %>
                    <option value="2GB">2GB</option>
                    <option selected value="4GB">4GB</option>
                <% else %>
                    <option value="2GB">2GB</option>
                    <option value="4GB">4GB</option>
                <% end if %>                   
	      </select>
        </td>
    </tr>

    <tr id="displayRow" style="display:<%=showOthRows%>">
        <td style="vertical-align:top; text-wrap:none; width:150px">
                    <b>Display:</b></td>
        <td>
            <select id="ddlTouch" name="ddlTouch" style="width: 200px">
                <option value="0"></option>
                <%  'get Spindle
                     set cm = server.CreateObject("ADODB.Command")
	                 Set cm.ActiveConnection = cn
                     set rs = server.CreateObject("ADODB.recordset")
                     rs.open "Select * from FormFactor_Touch order by Value", cn, adOpenForwardOnly
                     while Not rs.EOF
                        if trim(rs("Id")) = trim(sTouchID) then
                %>
                        <option selected value="<%=rs("Id")%>"><%=rs("Value")%></option>
                        <%else%>
                        <option value="<%=rs("Id")%>"><%=rs("Value")%></option>
                        <%end if
                    
                        rs.movenext
                    Wend
                    rs.Close
                    set rs = Nothing   
                %>
            </select>
        </td></tr>

    <tr>
		<td width="160" nowrap valign=top><b>Chassis:</b><font color="red" size="1"> *</font></td>
		<td width="100%">
            <select id="ddlChassis" name="ddlChassis">
            <%  ' get chassis
                 set cm = server.CreateObject("ADODB.Command")
	            Set cm.ActiveConnection = cn
                set rs = server.CreateObject("ADODB.recordset")
                 rs.open "usp_GetLisOfChassis ", cn, adOpenForwardOnly
                while Not rs.EOF
                if trim(rs("CategoryID")) = trim(sChassis) then
           %> 
                <option selected value="<%=rs("CategoryID")%>"><%=rs("Description")%></option>
                        <%else%>
                            <option value="<%=rs("CategoryID")%>"><%=rs("Description")%></option>
                        <%end if%>
                    <%
                 rs.movenext
                 Wend
                rs.Close
           set rs = Nothing   
                    %>      
	      </select></td>
    </tr>    
    <tr id="gnRow" style="display:<%=showOthRows%>">
		<td width="160" nowrap valign=top><b>Generic Name (Service Tags):</b><font color="red" size="1"> *</font></td>
		<td width="100%"><input type="text" id="txtGenericName" name="txtGenericName" value="<%=Server.HTMLEncode(sGenericName)%>" maxlength="40" size="75" />
            <font color="green" size="1">&nbsp;(40 maximum characters Do NOT include code names.)</font></td>
	</tr>
    <tr id="oldMktRow" style="display:<%=showOthRows%>">
		<td width="160" nowrap valign=top><b>Marketing Name (Product Master, SoftPaq/Web support):</b><font color="red" size="1"> *</font></td>
		<td width="100%"><input type="text" id="txtMarketingName" name="txtMarketingName" value="<%=Server.HTMLEncode(sMarketingName)%>" maxlength="50" size="75" />
            <font color="green" size="1">&nbsp;(50 maximum characters)</font></td>
	</tr>
	<tr id="fMKTRow" style="display:<%=showOthRows%>">
		<td width="160" nowrap valign=top><b>Full Marketing&nbsp;Name:</b></td>
		<td width="100%"><input type="text" id="txtFullMarketingName" name="txtFullMarketingName" value="<%=Server.HTMLEncode(sFullMarketingName)%>" maxlength="120" size="75" />
            <font color="green" size="1">&nbsp;(120 maximum characters)</font></td>
	</tr>
    <tr id="deploymentRow" style="display:<%=showOthRows%>">
		<td width="160" nowrap valign=top><b>Deployment:</b></td>
		<td width="100%">
            <select id="ddlDeployment" name="ddlDeployment">
                <option selected=""></option>
            <%  ' get chassis
                 set cm = server.CreateObject("ADODB.Command")
	            Set cm.ActiveConnection = cn
                set rs = server.CreateObject("ADODB.recordset")
                 rs.open "usp_ADMIN_GetNamingStandardDeploymentValues", cn, adOpenForwardOnly
                while Not rs.EOF
                if trim(rs("Value")) = trim(sDeployment) then
           %> 
                <option selected value="<%=rs("Value")%>"><%=rs("Value")%></option>
                        <%else%>
                            <option value="<%=rs("Value")%>"><%=rs("Value")%></option>
                        <%end if%>
                    <%
                 rs.movenext
                 Wend
                rs.Close
           set rs = Nothing   
                    %>      
	      </select></td>
    </tr>  
    
    <tr style="display:<%=showOthRows%>">
		<td width="160" nowrap valign=top><b>SOAR Product Description:</b>&nbsp;</td>
        <td width="100%">
            <select id="ddlSoarProductDescription" name="ddlSoarProductDescription">
                <option selected="0"></option>
           <%   ' get SOAR Products
                set cm = server.CreateObject("ADODB.Command")
	            Set cm.ActiveConnection = cn
                set rs = server.CreateObject("ADODB.recordset")
                rs.open "usp_Platform_SOARProdDesc", cn, adOpenForwardOnly
           while Not rs.EOF
          if trim(rs("OID")) = trim(sSOARID) then
        %>
                <option selected value="<%=rs("OID")%>"><%=rs("Name")%></option>
         <%else%>
                <option value="<%=rs("OID")%>"><%=rs("Name")%></option>
         <%
           end if
           rs.MoveNext
           Wend
           rs.Close
           set rs = Nothing
           %>
           </select></td>
	</tr>
   
    <tr>
		<td width="150" nowrap valign=top><b>Comments:</b>&nbsp;</td>
		<td width="100%"><input type="text" id="txtComments" name="txtComments" value="<%=Server.HTMLEncode(sComments)%>" maxlength="63" size="75" />
            <font color="green" size="1">&nbsp;(64 maximum characters)</font>
		</td>
	</tr>
    <tr>
		<td width="150" nowrap valign=top><b>System Board ID: </b>&nbsp;</td>
		<td width="100%"><input type="text" id="txtSystemID" name="txtSystemID" value="<%=Server.HTMLEncode(sSystemID)%>" maxlength="4" />
            <font color="green" size="1">&nbsp;(HEX, 4 maximum characters.)</font>
            <br />
            <b>System Board Comment: </b>&nbsp;
            <input type="text" id="txtSystemIdComments" name="txtSystemIdComments" value="<%=Server.HTMLEncode(sSystemIdComments)%>" maxlength="120" />
		</td>
	</tr>
    <tr>
        <td style="width:150px"><b>SMBIOS Family Name Override: </b>&nbsp;</td>
		<td style="width:100%"><input type="text" id="txtBFName" name="txtBFName" value="<%=Server.HTMLEncode(bFName)%>" maxlength="48" size="75"/>
    </tr>
        <tr>
		<td width="150" nowrap valign=top><b>Web Releases Cycles: </b>&nbsp;</td>
		<td width="10%">

             <%  ' get chassis
                 dim cyclecount
		         cyclecount = 0

                set cm = server.CreateObject("ADODB.Command")
	            Set cm.ActiveConnection = cn
                set rs = server.CreateObject("ADODB.recordset")
                 rs.open "IRS_usp_LCM_Spq_ViewWebReleasedCycle", cn, adOpenForwardOnly
                   ' cm.CommandType = 4
	                'cm.CommandText = "IRS_usp_LCM_Spq_ViewWebReleasedCycle"
                'set rs = cm.Execute
                while Not rs.EOF
                   cyclecount = cyclecount + 1
                 %>
                    <input type="checkbox"   name="ckWebCycle" ID="<%=rs("Description") %>" value="<%=rs("Description")%>"   <%=isChecked(rs("Value"))%> /> <%=rs("Value")%>
                            
                    <% if cyclecount = 3  then%>
                          <br /> 
                    <% cyclecount = 0

              end if %>                             
                 <% 
                
                rs.movenext
                Wend
                rs.Close
                set rs = Nothing   
                    %> 

		</td>
	</tr>
        <tr>
		<td width="150" nowrap valign=top><b>SoftPaq Need Date (on HP Support Web site): </b>&nbsp;</td>
		<td width="100%"><input type="text" id="txtSPQNeededDate" name="txtSPQNeededDate" readonly value="<%=Server.HTMLEncode(sSPQNeededDate)%>" maxlength="10" size="10" />
            <a id="hrefSPQNeededDate" href="javascript: cmdDate_onclick('txtSPQNeededDate')"><img id="imgProjectStartDtPicker" src="images/calendar.gif" alt="Choose Date" border="0" WIDTH="26" HEIGHT="21" /></a>
		</td>
	</tr>
<% if nMode=2 then %>
    <tr>
		<td width="150" nowrap valign=top><b>Status:</b>&nbsp;</td>
		<td width="100%"><select id="ddlStatus" name="ddlStatus">
              <% if sStatus = "0" then%>
                      <option selected value="0">In Progress</option>
                      <option  value="1">SW RTMed</option>
            <% else %>
                      <option  value="0">In Progress</option>
                      <option selected value="1">SW RTMed</option>
             <%end if%>
		                 </select></td>
    </tr>
    <tr>
		<td width="150" nowrap valign=top><b>Created By: </b>&nbsp;</td>
		<td width="100%"><span id="Creator" /><%=Server.HTMLEncode(sCreator)%></td>
	</tr>      
     <tr>
		<td width="150" nowrap valign=top><b>Created: </b>&nbsp;</td>
		<td width="100%"><span id="TimeCreated"  /><%=Server.HTMLEncode(sTimeCreated)%></td>
	</tr>
    <tr>
		<td width="150" nowrap valign=top><b>Updated By: </b>&nbsp;</td>
		<td width="100%"><span id="Updater" /><%=Server.HTMLEncode(sUpdater)%></td>
	</tr>
    <tr>
		<td width="150" nowrap valign=top><b>Updated: </b>&nbsp;</td>
		<td width="100%"><span id="TimeUpdated" /><%=Server.HTMLEncode(sTimeUpdated)%></td>
	</tr>
         
        
<% end if %>
</table>
</div>
    <INPUT type="hidden" id="txtID" name="txtID" value="<%=request("ProductVersionID")%>">
    <input type="hidden" id="txtProductionVersionID" name="txtProductionVersionID" value="<%=Server.HTMLEncode(sProductVersionID)%>" />
    <input type="hidden" id="txtPlatformID" name="txtPlatformID" value="<%=Server.HTMLEncode(sPlatformID)%>" />
    <input type="hidden" id="txtBusinessSegmentID" name="txtBusinessSegmentID" value="<%=Server.HTMLEncode(sBusinessSegmentID)%>" />
    <input type="hidden" id="txtCurrentUserID" name="txtCurrentUserID" value="<%=Server.HTMLEncode(CurrentUserID)%>" />
    <input type="hidden" id="txtCurrentUserName" name="txtCurrentUserName" value="<%=Server.HTMLEncode(CurrentUserName)%>" />
    <input type="hidden" id="txtProductFamily" name="txtProductFamily" value="<%=Server.HTMLEncode(sProductFamily)%>" />   
    <input type="hidden" id="txtProduct" name="txtProduct" value="<%=Server.HTMLEncode(sProduct)%>" />   
    <input type="hidden" id="hidVersion" name="hidVersion" value="<%=Server.HTMLEncode(productVersion)%>" />   
    <input type='hidden' id='hdnFormFactorAction' name='hdnFormFactorAction' value='<%= Server.HTMLEncode(sFormFactorID) %>' />
    <input type='hidden' id='hdnFormFactorID' name='hdnFormFactorID' value='<%= Server.HTMLEncode(sFormFactorID) %>' />
	<input type="hidden" id="txtInitialSystemID" name="txtInitialSystemID" value="<%=Server.HTMLEncode(sSystemID)%>">
    <input type="hidden" id="hdnFollowMKTName" name="hdnFollowMKTName" value="<%=followMKTName%>">
    <input type="hidden" id="txtBrandsAdded" name="txtBrandsAdded" value="<%=strBrandsAdded%>">
    <input type="hidden" id="txtBrandsLoaded" name="txtBrandsLoaded" value="<%=strBrandsLoaded%>">
    <input type="hidden" id="txtBrandNameLoaded" name="txtBrandNameLoaded" value="<%=strBrandNameLoaded%>">
    <input type="hidden" id="tagBrandId" name="tagBrandId" value ="<%=strBrandsLoaded%>" />
    <input type="hidden" id="tagSCMNumber" name="tagSCMNumber" value ="<%=strSCMNumber%>" />
    <input type="hidden" id="tagSeries" name="tagSeries" value ="<%=tagSeriesName%>" />
    <input type="hidden" id="tagModelNumber" name="tagModelNumber" value ="<%=strModelNumber%>" />
	<input type="hidden" id="tagScreenSize" name="tagScreenSize" value ="<%=strScreenSize%>" />
    <input type="hidden" id="tagGeneration" name="tagGeneration" value ="<%=strGeneration%>" />
    <input type="hidden" id="tagFF" name="tagFF" value ="<%=strFF%>" />
    <input type="hidden" id="tagMktName" name="tagMktName" value="<%=sMarketingName%>" />
    <input type="hidden" id="tagPhwFamilyName" name="tagPhwFamilyName" value="<%=strPhwFamilyName%>" />
    <input type="hidden" id="tagMatchPMS" name="tagMatchPMS" value="<%=strPMS%>" />
    <input type="hidden" id="showDCRDiv" name="showDCRDiv" value="<%=showDCR%>" />
    <input type="hidden" id="initialMktName" name="initialMktName" value="<%=sMarketingName%>" />
    <input type="hidden" id="hidMktName" name="hidMktName" value="<%=sMarketingName%>" />
    <input type="hidden" id="hidphwFamilyName" name="hidphwFamilyName" value="<%=strPhwFamilyName%>" />  
    <input type="hidden" id="hidSentPMSReq" name="hidSentPMSReq" value=<%=activeBtnSentPMReq%> />    
    <input type="hidden" id="hidPBID" name="hidPBID" value="<%=sBrand%>" />
    <input type="hidden" id="hidPMS" name="hidPMS" value="<%=strPMS%>" />
    <input type="hidden" id="hidPdStatus" name="hidPdStatus" value="<%=productStatus%>" />
    <input type="hidden" id="hidfirstPfID" name="hidfirstPfID" value="" />
    <input type="hidden" id="hidfirstPwfName" name="hidfirstPwfName" value="" />
</form>
<%

	set rs = nothing
	set cn = nothing

%>
</div>
<div style="display: none;">
    <div id="iframeDialog" title="Coolbeans">
        <iframe frameborder="0" name="modalDialog" id="modalDialog"></iframe>
    </div>
</div>
</BODY>
</HTML>