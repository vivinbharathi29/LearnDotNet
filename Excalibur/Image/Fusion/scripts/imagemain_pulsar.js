//******************************************************************
//File Description:     USER INTERFACE MANIPULATION
//Details:              Functions that manipulate UI 
//Created:              3/15/2016 Harris, Valerie - PBI 17835/ Task 18059
//Modified:             
//******************************************************************

//*****************************************************************
//Description:  OnLoad, on page load instantiate functions
//*****************************************************************
$(window).load(function () {
    //*********APPLY JS FUNCTIONS PAGE ELEMENTS*******//
    //Instantiate Elements
    LoadUIElement(true);

});

//*****************************************************************
//Description:  Load Jquery UI Elements
//Function:     LoadUIElement();
//*****************************************************************
function LoadUIElement(bPageLoad) {
    if (bPageLoad === true) {
        CopyWithTargeting();
    }
}

//*****************************************************************
//Description:  Onload, change form elements for Copy With Targeting
//Function:     CopyWithTargeting();
//Created By:   Harris, Valerie (3/15/2016) - PBI 17835/ Task 18059   
//*****************************************************************
function CopyWithTargeting() {
    var oSelect = $('#cboOS');
    var bCopyWithTarget = $('#inpCopyWithTarget').val();

    if (bCopyWithTarget === 'True') {
        //Prevent User from changing operating system: --
        oSelect.attr("readonly", true);
        oSelect.addClass("disabled");
        $('#cboOS option:not(:selected)').remove();

        if ($("#cmdAddOS").length > 0) {
            $("#cmdAddOS").addClass("hide");
        } else {
            $("#cmdAddOS").addClass("show");
        }
    }
}

