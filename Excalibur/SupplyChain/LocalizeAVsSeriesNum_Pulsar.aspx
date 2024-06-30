<%@ Page Language="VB" AutoEventWireup="false"
    Inherits="DummyVBApp.LocalizeAVsSeriesNum_Pulsar" Codebehind="LocalizeAVsSeriesNum_Pulsar.aspx.vb" %>
<%@ Register Assembly="eWorld.UI, Version=2.0.6.2393, Culture=neutral, PublicKeyToken=24d65337282035f2"
    Namespace="eWorld.UI" TagPrefix="ew" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<script type="text/javascript" src="../Scripts/jquery-1.4.4.js"></script>
<script  src="../includes/client/jquery-1.10.2.js" type="text/javascript" ></script>
<script  src="../includes/client/jquery-ui-1.10.4.js" type="text/javascript"></script>    
<script src="../Scripts/shared_functions.js" type="text/javascript"></script>
<script type="text/javascript">
    $(function () {        
        //load_datePicker();
        $("#txtRTPDate").datepicker({
            showOn: 'button',
            buttonText: 'RTP Date',
            buttonImageOnly: true,
            buttonImage: '/Excalibur/images/calendar.gif',
            dateFormat: "mm/dd/yy",
            changeMonth: true,
            changeYear: true,
            firstDay: 7
        });
        $("#txtEMDate").datepicker({
            showOn: 'button',
            buttonText: 'EM Date',
            buttonImageOnly: true,
            buttonImage: '/Excalibur/images/calendar.gif',
            dateFormat: "mm/dd/yy",
            changeMonth: true,
            changeYear: true,
            firstDay: 7
        });
        LoadProductRelease();
        //LoadDefaultDates();
        $("input[name=chkRelease]").change(function () {
            var sRTPDate = "";
            var sEMDate = "";

            $('input[name=chkRelease]:checkbox:checked').each(function () {
                if (sRTPDate == "") {
                    sRTPDate = $(this).attr('RTPDate');
                }
                else {
                    if ((new Date(sRTPDate).getTime() > new Date($(this).attr('RTPDate')).getTime())) {
                        sRTPDate = $(this).attr('RTPDate');
                    }
                }

                if (sEMDate == "") {
                    sEMDate = $(this).attr('EMDate');
                }
                else {
                    if ((new Date(sEMDate).getTime() < new Date($(this).attr('EMDate')).getTime())) {
                        sEMDate = $(this).attr('EMDate');
                    }
                }
            });

            $('#txtRTPDate').val(sRTPDate);
            $('#txtEMDate').val(sEMDate);
        });

    });
    function cmdCancel_onclick() {
        var pulsarplusDivId = $('#pulsarplusDivId').val();
        if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
            // For Closing current popup if Called from pulsarplus
            parent.window.parent.closeExternalPopup();
        }
        else {
            window.parent.modalDialog.cancel();
        }
    }
    function LoadDefaultDates()
    {
        $('#txtRTPDate').val($("#hdnRTPDate").val());
        $('#txtEMDate').val($("#hdnEMDate").val());
    }
    function ValidateData() {
        //get the selected releases
        var strReleaseIDs = "";
        $('input:checkbox[name="chkRelease"]:checked').each(function () {
            strReleaseIDs = strReleaseIDs == "" ? $(this).val() : strReleaseIDs + "," + $(this).val();
        });
        if (strReleaseIDs == "")
        {
            alert("Please select at least one release.");
            return false;
        }
        var strRTPDate = $('#txtRTPDate').val();
        var strEMDate = $('#txtEMDate').val();

        $("#hdnRTPDate").val(strRTPDate);
        $("#hdnEMDate").val(strEMDate);        
        $("#hdnSelectedReleases").val(strReleaseIDs);
            
        return true;
    }
    function LoadProductRelease() {
        $("#divRelease").empty();
        var productreleases = $("#hdnProductReleases").val();
        if (productreleases.length > 0)
        {
            var releases = productreleases.split(';');
            for (var i = 0; i < releases.length; i++)
            {
                var items = releases[i].split(',');

                if (items[0] != "") {
                    var checkBox = document.createElement("input");
                    checkBox.type = "checkbox";
                    if (productreleases.split(';').length == 1) {
                        checkBox.checked = true;
                        $('#txtRTPDate').val(items[2]);
                        $('#txtEMDate').val(items[3]);
                    }
                    else {
                        checkBox.checked = false;
                    }

                    $(checkBox).attr('name', 'chkRelease');
                    $(checkBox).attr('id', items[1]);
                    $(checkBox).attr('value', items[1]);
                    $(checkBox).attr('RTPDate', items[2]);
                    $(checkBox).attr('EMDate', items[3]);
                    $(checkBox).attr('style', 'margin:0px; padding-top:0px;padding-bottom:0px;margin-left:5px;');

                    var spRelease = document.createElement('span');
                    spRelease.innerHTML = items[0];
                    $(spRelease).attr('id', 'spRelease' + items[0]);
                    $(spRelease).attr('style', 'margin:0px; padding-top:0px;padding-bottom:0px;margin-right:5px;');

                    $('#divRelease').append(checkBox);
                    $('#divRelease').append(spRelease);               

                }
            }            
        }        
    }
</script>

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
    <link href="../style/general.css" rel="stylesheet" type="text/css" />    
    <link href="../includes/client/jqueryui/jquery-ui-1.11.4/jquery-ui.css" rel="stylesheet" />

</head>
<body runat="server" id="thisBody">
    <form id="form1" runat="server">
    <textarea id="holdtext" style="display: none;" rows="0" cols="0"></textarea>
    <div style="width: 99%; height: 99%;">
        <asp:Label ID="lblError" runat="server" Text="Feature Categories haven't been added to Pulsar. Try to add Localized AVs later or contact a Pulsar Administrator." Width="274px" Style="font-size: 10px;
                                font-weight: bold; font-family: Verdana; display:none;"></asp:Label>
        
            <table width="100%" cellspacing="0" cellpadding="2" >                
                <tr>
                    <td style=" font-size:10px;font-weight: bold; font-family: Verdana;">AV Type:</td>
                    <td >
                        <asp:DropDownList ID="ddlAvType" runat="server" Style="height: 23px; margin-bottom: 0px;" DataTextField="Name"
                            DataValueField="FeatureCategoryId" AutoPostBack="false" EnableViewState="true">            
                        </asp:DropDownList> 
                    </td>
                </tr>               
                <tr>
                    <td style=" font-size:10px;font-weight: bold; font-family: Verdana;">Release(s):<span style="color:red">*</span></td>
                    <td><div id="divRelease" style="margin-right:5px;"></div></td>
                </tr>
                <tr>
                    <td style=" font-size:10px;font-weight: bold; font-family: Verdana;">RTP/MR Date:</td>                             
                    <td> 
                        <input type="text" id="txtRTPDate" name="txtRTPDate" onkeydown="return false;" class="dateselection-validate" />				                             
                    </td>
                </tr>
                <tr>
                    <td style=" font-size:10px;font-weight: bold; font-family: Verdana;">EM Date:</td>                             
                    <td> 
                        <input type="text" id="txtEMDate" name="txtEMDate" onkeydown="return false;" class="dateselection-validate"  />				                                                        
                    </td>
                    
                </tr>
                <tr>
                    <td colspan="2"><span style="color:red; font-weight:bold">Note: The selected date will be used for all the AVs.</span></td>
                </tr>
                <tr>
                    <td colspan="2">
                    <asp:Checkbox ID="cbxShowAllLocs" runat="server">
                            </asp:Checkbox>
                            <asp:Label ID="lblShowAllLocs" runat="server" Text="Show all product localizations from admin regardless of the Localization Tab Support"
                                Style=" width: 250px;">
                            </asp:Label>
                    </td>
                </tr>
            </table>      
        
        <br />
        <br />
        <hr />        
        <asp:Button ID="btnCancel" runat="server" Text="Cancel" Style="position: absolute;
            left: 220px; width: 61px; height: 24px" OnClientClick="cmdCancel_onclick()" />
		<asp:Button ID="btnSubmit" OnClientClick="return ValidateData();" runat="server" Text="OK" Style="position: absolute; left: 179px;
            width: 35px; height: 24px" />
    </div>
    <asp:HiddenField ID="hdnPath" runat="server" Value="" />
        <asp:HiddenField ID="hdnProductVersionId" runat="server" Value="0" />
        <asp:HiddenField ID="hdnProductReleases" runat="server" Value="" />
        <asp:HiddenField ID="hdnSelectedReleases" runat="server" Value="" />
        <asp:HiddenField ID="hdnRTPDate" runat="server" Value="" />
        <asp:HiddenField ID="hdnEMDate" runat="server" Value="" />
        <input type="hidden" id="pulsarplusDivId" value="<%=Request("pulsarplusDivId")%>" />
    </form>
</body>
</html>
<script type="text/javascript">
    //*****************************************************************
    //Description:  On page load instantiate functions
    //*****************************************************************
    $(window).load(function () {        
        //Bug 15291 / Task 15294 - clear message when "Add Localized AVs" but Localized is not assigned to any Feature
        var iAVTotal = $('#ddlAvType option').length;
        if(iAVTotal === 0){
            $("#lblError").show();
            $("#lblHeader").hide();
            $("#lblHeader").hide();
            $("#ddlAvType").hide();
            $("#btnSubmit").hide();
        }else{
           $("#lblError").hide(); 
        }   
    });
</script>
