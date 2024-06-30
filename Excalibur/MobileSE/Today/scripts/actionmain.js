//******************************************************************
//File Description:     USER INTERFACE MANIPULATION
//Details:              Functions that manipulate UI 
//Created:              Harris, Valerie (2/5/2016) - PBI 15660/ Task 16234
//******************************************************************
var $browser = null;
var $browserversion = null;

//*****************************************************************
//Description:  OnLoad, on page load instantiate functions
//*****************************************************************
$(window).load(function () {
    //*********APPLY JS FUNCTIONS PAGE ELEMENTS*******//
    //Browser Detection
    BrowserDetection();

    //Instantiate Elements
    LoadUIElement(true);

});

//*****************************************************************
//Description:  Load Jquery UI Elements
//Function:     LoadUIElement();
//*****************************************************************
function LoadUIElement(bPageLoad) {
    if (bPageLoad === true) {
        //Get Browser Name and Version: ---
        $browser = $("html").attr("data-browser");
        $browserversion = $("html").attr("data-version");

        //Add click event to select-menu icons
        ClickReleaseMenu();

        //Add click event to product checkboxes
        ClickProductCheckbox();

        //Add click event to product line checkboxes
        ClickProductLineCheckbox();
    }
}

//*****************************************************************
//Description:  OnClick, check/uncheck Product checkboxes, it's Release checkboxes and rows            
//Function:     SelectMultipleProductCheckbox();
//*****************************************************************
function SelectMultipleProductCheckbox(sProductList, sSelectionType, iReleaseID) {
    var aProduct = sProductList.split(',');
    var ProdCount = 0;
    var UpdateCount = 0;
    var sProductName = "";

    for (i = 0; i < aProduct.length; i++) {
        ProdCount++;
        for (j = 0; j < ProgramInput.chkProducts.length; j++) {
            if (aProduct[i] == ProgramInput.chkProducts[j].value) {
                iProductID = ProgramInput.chkProducts[j].value;
                
                //If Product has releases, change icon, show and check releases
                if($('table#TBLProducts span[data-product="' + iProductID + '"] em.row-option').length > 0){
                    $('table#TBLProducts #product_' + iProductID + '').prop('checked',true);
                    ChangeMenuIcon(iProductID);
                    SelectReleaseMenu(iProductID, true);
                    if (sSelectionType != "Business Segment")
                        SelectReleaseCheckbox(iProductID, true, sSelectionType);
                    else
                        SelectReleaseCheckbox2(iProductID, true, true, iReleaseID);
                }else{
                    //Check Product checkbox
                    $('table#TBLProducts #product_' + iProductID + '').prop('checked',true);
                }
                if ($('table#TBLProducts #product_' + iProductID + '').length > 0) {
                    if (sProductName == "") {
                        sProductName = $('table#TBLProducts #product_' + iProductID + '').attr("data-productname");
                    } else {
                        sProductName = $('table#TBLProducts #product_' + iProductID + '').attr("data-productname") + ',' + sProductName
                    }
                }
                UpdateCount++;
                break;
            }
        }
    } 

    //Remove trailing comma: ----
    if(sProductName === ""){
        sProductName = "None";
    }else{
        sProductName = sProductName.replace(/,\s*$/, "");
        $.trim(sProductName);
    }

    //Auto select product line checkbox only if using Product Group or Business Segment: ---
    if (sProductName != "None" && sSelectionType != "Product Line") {
        SelectProductLineCheckbox();
    }

    if (UpdateCount == ProdCount && UpdateCount == 1){
        alert("Automatically selected the product defined for this " + sSelectionType + ".\r\rPlease verify the product list was updated correctly.");
    }else if (UpdateCount == ProdCount) {
        alert("Automatically selected all " + UpdateCount + " products defined for this " + sSelectionType + ".\r\rPlease verify the product list was updated correctly.");
    }else if (ProdCount == 1) {
        alert("Unable to find products defined for this " + sSelectionType + ".  No products have been selected.");
    }else{
        alert("Automatically selected only " + UpdateCount + " of the " + ProdCount + " products defined for this " + sSelectionType + ".\r\rPlease verify the product list was updated correctly.\n\nProducts Selected:\n " + sProductName + ".");
    }
}


//*****************************************************************
//Description:  Onload, Add Onclick event to product checkboxes            
//Function:     ClickProductCheckbox();
//*****************************************************************
function ClickProductCheckbox() {
    $('table#TBLProducts .chk-product').click(function (e) {
        var iProductID = $(this).val();
        var iProductLineID = $(this).attr("data-productline");
        var bChecked = $(this).prop('checked');

        ChangeMenuIcon(iProductID);
        SelectReleaseMenu(iProductID, bChecked, false);
        SelectReleaseCheckbox(iProductID, bChecked);
        ProductLineStatus(iProductLineID);
    });
}


//*****************************************************************
//Description:  Onload, Add Onclick event to select-menu             
//Function:     ClickReleaseMenu();
//*****************************************************************
function ClickReleaseMenu() {    
    $('table#TBLProducts span.select-menu-release').click(function (e) {
        var iProductID = $(this).attr("data-product");
        var iProductLineID = $(this).attr("data-productline");
        var bPlusIcon = null;
        var bChecked = null;
        var oCheckBox = null;

        oCheckBox = $('table#TBLProducts #product_' + iProductID + '');
        bPlusIcon = $('table#TBLProducts span[data-product="' + iProductID + '"] em.row-option').hasClass('fa fa-plus-square');

        //Check product checkbox if it's not checked
        if(bPlusIcon == true){
            oCheckBox.attr("checked", true);
            oCheckBox.prop('checked', true);
            bChecked = true;
        } else {
            oCheckBox.attr("checked", false);
            oCheckBox.prop('checked', false);
            bChecked = false;
        }
        
        ChangeMenuIcon(iProductID);
        SelectReleaseMenu(iProductID, bChecked, false);
        SelectReleaseCheckbox(iProductID, bChecked);
        ProductLineStatus(iProductLineID);
    });
}


//*****************************************************************
//Description:  OnClick, Show/Hide Release Rows for selected Product             
//Function:     SelectReleaseMenu();
//*****************************************************************
function SelectReleaseMenu(iProductID, bChecked, bGroupSelect) {

    if (bGroupSelect == true) { //for multiple selections, display hidden release rows
        if ($('table#TBLProducts tr[data-product="' + iProductID + '"]').hasClass(ShowRowClass()) === false) {
            $('table#TBLProducts tr[data-product="' + iProductID + '"]').removeClass('hide').addClass(ShowRowClass());
        }
    } else {
        if (bChecked === true) {
            //if($('table#TBLProducts tr[data-product="' + iProductID + '"]').hasClass(ShowRowClass()) === false){
                $('table#TBLProducts tr[data-product="' + iProductID + '"]').removeClass('hide').addClass(ShowRowClass());
            //}
        } else {
            //if($('table#TBLProducts tr[data-product="' + iProductID + '"]').hasClass('hide') === false){
                $('table#TBLProducts tr[data-product="' + iProductID + '"]').removeClass(ShowRowClass()).addClass('hide');
            //}
        }
    }
}

//*****************************************************************
//Description:  OnClick, Check/Uncheck Release Checkboxes for selected Product             
//Function:     SelectReleaseCheckbox();
//*****************************************************************
function SelectReleaseCheckbox(iProductID, bChecked, sSelectionType) {
    var oCheckBoxes = document.getElementsByName("chkRelease_" + iProductID);
    var oCheckBox = $('[name=chkRelease_' + iProductID + ']');

    if (bChecked === true) {
        if (oCheckBoxes.length == 1) {
            oCheckBoxes[0].checked = true;
        }
        else if (sSelectionType !== undefined) {
            oCheckBox.attr("checked", true);
            oCheckBox.prop('checked', true);
        }
    }
    else {
        for (var i = 0; i < oCheckBoxes.length; i++) {
            if (oCheckBoxes[i].checked === true) {
                oCheckBoxes[i].checked = false;
            }
        }
    }
}



//*****************************************************************
//Description:  OnClick, Check/Uncheck Child Release Checkboxes for selected Product - release            
//Function:     SelectReleaseCheckbox2();
//*****************************************************************
function SelectReleaseCheckbox2(iProductID, bChecked, ckbChecked, sReleaseID) {
    var chldCheckBox = null;
    var pReleaseId = sReleaseID + "_" + iProductID;
        chldCheckBox = $('table#TBLProducts .release_' + pReleaseId + '');

    if (bChecked === true) {                                
        //Check all checkboxes for the selected product
       chldCheckBox.attr("checked", true);
       chldCheckBox.prop('checked', true);

    } else {
        //Un-check all checkboxes for the selected product
        chldCheckBox.attr("checked", false);
        chldCheckBox.prop('checked', false);
    }
}


//*****************************************************************
//Description:  Change Plus/Minus Icon by Product ID
//Function:     ChangeMenuIcon();
//*****************************************************************
function ChangeMenuIcon(iProductID) {
    if ($('table#TBLProducts span[data-product="' + iProductID + '"] em.row-option').hasClass('fa fa-plus-square')) {
        $('table#TBLProducts span[data-product="' + iProductID + '"] em.row-option').removeClass('fa fa-plus-square').addClass('fa fa-minus-square');
    } else {
        $('table#TBLProducts span[data-product="' + iProductID + '"] em.row-option').removeClass('fa fa-minus-square').addClass('fa fa-plus-square');
    }
}


//*****************************************************************
//Description:  Based on browser, use class name to show table row
//Function:     ShowRowClass();
//*****************************************************************
function ShowRowClass() {
    var sClassShowRowName = null;

    if ($browser === 'Chrome' || $browser === 'Opera' || $browser === 'Firefox' || $browser === 'Netscape' || $browser == 'MSIE' && $browserversion > 10) {
        sClassShowRowName = 'show-row';
    } else {
        sClassShowRowName = 'show';
    }
    return sClassShowRowName;
}

//*****************************************************************
//Description:  Detect Browser Type
//Function:     BrowserDetection();
//*****************************************************************
function BrowserDetection() {
    //Browser Detection
    $("html").attr("data-browser", get_browser());
    $("html").attr("data-version", parseInt(get_browser_version()));

    function get_browser() {
        var N = navigator.appName, ua = navigator.userAgent, tem;
        var M = ua.match(/(opera|chrome|safari|firefox|msie)\/?\s*(\.?\d+(\.\d+)*)/i);
        if (M && (tem = ua.match(/version\/([\.\d]+)/i)) !== null) M[2] = tem[1];
        M = M ? [M[1], M[2]] : [N, navigator.appVersion, '-?'];
        return M[0];
    }
    function get_browser_version() {
        var N = navigator.appName, ua = navigator.userAgent, tem;
        var M = ua.match(/(opera|chrome|safari|firefox|msie)\/?\s*(\.?\d+(\.\d+)*)/i);
        if (M && (tem = ua.match(/version\/([\.\d]+)/i)) !== null) M[2] = tem[1];
        M = M ? [M[1], M[2]] : [N, navigator.appVersion, '-?'];
        return M[1];
    }
}

//*****************************************************************
//Description:  Determine if one or more Product checkboxes of the same group is/isn't checked 
//Function:     isProductLineSelected();
//*****************************************************************
function isProductLineSelected(iGroupID) {
    var bChecked = false;
    var iChecked= 0;

    iChecked = $('table#TBLProducts input.chk-product[type="checkbox"][data-productline="' + iGroupID + '"]:checked').length;
   
    if (iChecked > 0) {
        bChecked = true;
    }

    if (bChecked === true) {
        return true;
    } else {
        return false;
    }
}

//*****************************************************************
//Description:  Onload, Add Onclick event to product line checkboxes            
//Function:     ClickProductLineCheckbox();
//*****************************************************************
function ClickProductLineCheckbox() {
    $('#divProductLine .chk-productline').click(function (e) {
        var iProductLineID = $(this).attr("data-productline");
        var sProductList = $(this).val();
        var bChecked = $(this).prop('checked');

        if (bChecked === true) {
            SelectMultipleProductCheckbox(sProductList, 'Product Line');
            ProductLineStatus(iProductLineID);
        } else {
            UnSelectMultipleProductCheckbox(iProductLineID, bChecked);
        }
    });
}


//*****************************************************************
//Description:  Uncheck multiple Product checkboxes associated with a Product Line            
//Function:    UnSelectMultipleProductCheckbox();
//*****************************************************************
function UnSelectMultipleProductCheckbox(iProductLineID, bChecked) {
    var iProductID = null;

    //For each product associated with the Product Line, uncheck
    $('table#TBLProducts input[type="checkbox"][data-productline="' + iProductLineID + '"]').each(function () {
        iProductID = $(this).val();

        $(this).prop('checked', false);
        ChangeMenuIcon(iProductID);
        SelectReleaseMenu(iProductID, bChecked, false);
        SelectReleaseCheckbox(iProductID, bChecked);
    });
}

//*****************************************************************
//Description:  Check/UnCheck Product Line checkbox           
//Function:     ProductLineStatus();
//*****************************************************************
function ProductLineStatus(iProductLineID) {
    var oProductLineChkBox = $('#divProductLine input[type="checkbox"][data-productline="' + iProductLineID + '"]');

    //Un-check Product Line checkbox under Select Product by Product Line 
    if (isProductLineSelected(iProductLineID) === false) {
        oProductLineChkBox.attr("checked", false);
        oProductLineChkBox.prop('checked', false);
    } else {
        oProductLineChkBox.attr("checked", true);
        oProductLineChkBox.prop('checked', true);
    }
}

//*****************************************************************
//Description:  Clear all Product selections           
//Function:     ClearMultipleProductCheckbox();
//*****************************************************************
function ClearProductList() {
    var iProductID = null;

    //For each product associated with the Product Line, uncheck
    $('table#TBLProducts input.chk-product[type="checkbox"]').each(function () {
        iProductID = $(this).val();

        $(this).prop('checked', false);
        ChangeMenuIcon(iProductID);
        SelectReleaseMenu(iProductID, false, false);
        SelectReleaseCheckbox(iProductID, false);
    });

    $('input.chk-productline[type="checkbox"]:not(:disabled)').prop('checked', false);
}

//*****************************************************************
//Description:  Check if one or more Products are selected in list 
//              and select Product Line checkbox
//Function:     SelectProductLineCheckbox();
//*****************************************************************
function SelectProductLineCheckbox() {
    var sProductLineIDs = $("#inpProductLineIDs").val();
    var aProductLineIDs = null;
    var iProductLineID = null;
    var x;
    
    if(sProductLineIDs != ""){
        aProductLineIDs = sProductLineIDs.split(",");
    }

    for (x in aProductLineIDs) {
        iProductLineID = aProductLineIDs[x];
        ProductLineStatus(iProductLineID);
    }   
}

//*****************************************************************
//Description:  Show/Hide Business Segment's Releases in Select Product by option list
//Function:     ShowBusSegReleases();
//*****************************************************************
function ShowBusSegReleases(iBusinessSegmentID) {
    var oReleaseList = $("ul#busseg_" + iBusinessSegmentID + "");

    if (oReleaseList.hasClass('hide')) {
        oReleaseList.removeClass('hide').addClass('show');
    } else {
        oReleaseList.removeClass('show').addClass('hide');
    }
}