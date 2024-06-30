<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS type="text/javascript">
<!--

function cmdCancel_onclick() {
	window.close();
}

function cmdOK_onclick() {

    frmMain.submit();

    window.parent.opener = 'X';
    window.parent.open('', '_parent', '')
    window.parent.close();
}

function chkAll_onclick() {
    frmMain.Section1.checked = frmMain.chkAll.checked;
    if (frmMain.ReportType.value != "3") {
        frmMain.Section2.checked = frmMain.chkAll.checked;
        frmMain.Section3.checked = frmMain.chkAll.checked;
    }
    frmMain.Section4.checked = frmMain.chkAll.checked;
    frmMain.Section5.checked = frmMain.chkAll.checked;
    frmMain.Section6.checked = frmMain.chkAll.checked;
    frmMain.Section7.checked = frmMain.chkAll.checked;
    frmMain.Section8.checked = frmMain.chkAll.checked;
    frmMain.Section9.checked = frmMain.chkAll.checked;
    frmMain.Section10.checked = frmMain.chkAll.checked;
    frmMain.Section11.checked = frmMain.chkAll.checked;
}

//-->
</SCRIPT>
</HEAD>
<STYLE>
BODY
{
    FONT-FAMILY: Verdana;
    FONT-SIZE: x-small;
}
TD
{
    FONT-FAMILY: Verdana;
    FONT-SIZE: x-small;
}

H1
{
    FONT-FAMILY: Verdana;
    FONT-SIZE: small;
}
.DateInput
{
	font-family: Verdana;
	font-size: x-small;
	height: 20;
	width: 80;
	border: solid 1px gray;
}
</STYLE>
<body style="background-color:ivory;" id="MainBody">
<form action="ReadinessReport.asp" target=_blank method="post" id="frmMain" name="frmMain">
<h1>Readiness Report Options</h1>
<font size=2><b>Choose Report Sections</b><BR></font>
    <div style="border-right: steelblue 1px solid; border-top: steelblue 1px solid; overflow-y: auto; border-left: steelblue 1px solid; width: 100%; border-bottom: steelblue 1px solid; background-color: white" id="DIV1">
    <table id="tabSections" width="100%">
        <thead>
            <tr style="position: relative; top: expression(document.getElementById('DIV1').scrollTop-2);">
                <td bgcolor="lightsteelblue" nowrap style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset; border-bottom: 1px outset"><input type="checkbox" id="chkAll" name="chkAll" language=javascript onclick="chkAll_onclick();" ></td>
                <td bgcolor="lightsteelblue" style="width:100%; border-right: 1px outset; border-top: 1px outset;border-left: 1px outset; border-bottom: 1px outset">Sections</td>
            </tr>
        </thead>
        <tbody>
        <tr><td width=10><input type="checkbox" id="Section1" name="Sections" checked value=1>&nbsp;</td><td>Build Level Alerts</td></tr>
        <%if trim(request("ReportType")) <> "3" then%>
            <tr><td width=10><input type="checkbox" id="Section2" name="Sections" checked value=2>&nbsp;</td><td>Distribution Alerts</td></tr>
            <tr><td width=10><input type="checkbox" id="Section3" name="Sections" checked value=3>&nbsp;</td><td>Certification Alerts</td></tr>
        <%end if%>
        <tr><td width=10><input type="checkbox" id="Section4" name="Sections" checked value=4>&nbsp;</td><td>Workflow Alerts</td></tr>
        <tr><td width=10><input type="checkbox" id="Section5" name="Sections" checked value=5>&nbsp;</td><td>Availability Alerts</td></tr>
        <tr><td width=10><input type="checkbox" id="Section6" name="Sections" checked value=6>&nbsp;</td><td>Developer Alerts</td></tr>
        <tr><td width=10><input type="checkbox" id="Section7" name="Sections" checked value=7>&nbsp;</td><td>Root Deliverable Alerts</td></tr>
        <tr><td width=10><input type="checkbox" id="Section8" name="Sections" checked value=8>&nbsp;</td><td>OTS Alerts - Primary Product</td></tr>
        <tr><td width=10><input type="checkbox" id="Section9" name="Sections" value=9>&nbsp;</td><td>OTS Alerts - Related Products</td></tr>
        <tr><td width=10><input type="checkbox" id="Section10" name="Sections" checked value=10>&nbsp;</td><td>OTS Alerts - Related Product Observation Counts</td></tr>
        <tr><td width=10><input type="checkbox" id="Section11" name="Sections" value=11>&nbsp;</td><td>OTS Alerts - Primary and Related Products</td></tr>
        </tbody>
    </table>
    </div>
<input type="hidden" id=ProdID name=ProdID value="<%=request("ProdID")%>" />
<input type="hidden" id=ReportType name=ReportType value="<%=request("ReportType")%>" />
<input type="hidden" id=TeamID name=TeamID value="<%=request("TeamID")%>" />
</FORM>

</BODY>
</HTML>
