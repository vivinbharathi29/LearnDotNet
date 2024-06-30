<%@ Language=VBScript %>

<!-- #include file = "../includes/noaccess.inc" -->

<html>
<head>
    <title></title>
    <link href="../style/redmond/jquery-ui-1.8.7.custom.css" rel="stylesheet" type="text/css" />
    <script src="../Scripts/jquery-1.10.2.js"></script>
    <script src="../includes/client/jquery.min.js" type="text/javascript"></script>
    <script src="../includes/client/jquery-ui.min.js" type="text/javascript"></script>
    <script type="text/javascript">
        $(function () {
            $("input:button").button();
        });
    </script>

<script type="text/javascript" language="JavaScript" src="../includes/client/Common.js"></script>
<script src="../Scripts/PulsarPlus.js"></script>
<SCRIPT type="text/javascript" language="JavaScript" >
<!--

function frmMilestoneVerify()
{
	with (window.parent.frames["UpperWindow"].frmMilestone)
	{
	if (!validateDateInput(txtProjectedStartDt, 'Projected Start Date')){	return false; }
	if (!validateDateInput(txtProjectedEndDt, 'Projected End Date')){	return false; }
	if (!validateDateInput(txtActualStartDt, 'Actual Start Date')){	return false; }
	if (!validateDateInput(txtActualEndDt, 'Actual End Date')) { return false; }
	if ((hidPorStartDt.value != "") && ((hidProjectedStartDt.value != txtProjectedStartDt.value) || (hidProjectedEndDt.value != txtProjectedEndDt.value)) && (!validateTextInput(txtItemNotes, 'Changes Notes'))){ return false; }
	}
	return true;
}

function cmdCancel_onclick() {
    if (pulsarplusDivId.value != undefined && pulsarplusDivId.value != "") {

        // For Closing current popup
        parent.window.parent.closeExternalPopup();
    } else {
        var iframeName = parent.window.name;
        if (iframeName != '') {
            parent.window.parent.ClosePropertiesDialog();
        } else {
            if (parent.window.parent.document.getElementById('modal_dialog')) {
                parent.window.parent.modalDialog.cancel();
            } else {
                window.parent.close();
            }
        }
    }
}

function cmdOK_onclick() 
{
	cmdCancel.disabled = true;
	cmdOK.disabled = true;
	if (window.parent.frames["UpperWindow"].frmMilestone)
	{
		if (frmMilestoneVerify())
		{
			with (window.parent.frames["UpperWindow"].frmMilestone)
			{
				if (hidMilestone.value.toUpperCase() == "TRUE")
				{	
					txtProjectedEndDt.value = txtProjectedStartDt.value;
					txtActualEndDt.value = txtActualStartDt.value;
				}
			}
			window.returnValue=1;
			window.parent.frames["UpperWindow"].frmMilestone.submit();
		}
		else
		{
			cmdCancel.disabled = false;
			cmdOK.disabled = false;
		}
	}	
	else
	{
		window.parent.frames["UpperWindow"].frmSchedule.submit();
	}
}

//-->
</SCRIPT>
</head>
<body bgcolor="ivory">

    <INPUT type="hidden" id=pulsarplusDivId name=pulsarplusDivId value="<%=request("pulsarplusDivId")%>">
<table BORDER="0" CELLSPACING="1" CELLPADDING="1" align="right">
	<tr>
		<TD><INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()"></TD>
		<TD><INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel  LANGUAGE=javascript onclick="return cmdCancel_onclick()"  ></TD>
</TR></table>
</body>
</html>