<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<script language="javascript" src="../../_ScriptLibrary/jsrsClient.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

    var SyntaxStatus = 0;

    function AddValue(strValue){
        
        var strPrefix="";

        if (strValue.substr(0,4) == "like")
            strPrefix="";
        else if (strValue.substr(0,7).toLowerCase() == "between")
            strPrefix="";
        else if (strValue.substr(0,1).toLowerCase() == "<")
            strPrefix="";
        else if (strValue.substr(0,1).toLowerCase() == ">")
            strPrefix="";
        else if (strValue.substr(0,7) == "is null")
            strPrefix="";
        else if (strValue.substr(0,4) == "in (")
            strPrefix="";
        else
            strPrefix="= ";

        if (frmMain.txtAdvanced.value =="")
            insertAtCursor(frmMain.txtAdvanced, strValue + " ");
        else
            insertAtCursor(frmMain.txtAdvanced," " + strPrefix + strValue + " ");

}


    function AddField(strField){
 
        if (frmMain.txtAdvanced.value =="")
            insertAtCursor(frmMain.txtAdvanced, strField + " ");
        else
            insertAtCursor(frmMain.txtAdvanced," " + strField + " ");

        divTypicalValues.style.display = "";
       // divFieldName.innerText = strField;
       if(strField == "Status")
            divTypical.innerHTML = "Status: " + typicalStatus.innerHTML + "<BR>" + divTypical.innerHTML.replace("Status: " + typicalStatus.innerHTML + "<BR>","");
       else if (strField == "Priority")
            divTypical.innerHTML = "Priority: " + typicalPriority.innerHTML + "<BR>" + divTypical.innerHTML.replace("Priority: " + typicalPriority.innerHTML + "<BR>","");
       else if (strField == "PrimaryProduct")
            divTypical.innerHTML = "PrimaryProduct: " + typicalProduct.innerHTML + "<BR>" + divTypical.innerHTML.replace("PrimaryProduct: " + typicalProduct.innerHTML + "<BR>","");
       else if (strField == "ProductFamily")
            divTypical.innerHTML = "ProductFamily: " + typicalProductFamily.innerHTML + "<BR>" + divTypical.innerHTML.replace("ProductFamily: " + typicalProductFamily.innerHTML + "<BR>","");
       else if (strField == "AffectedProduct")
            divTypical.innerHTML = "AffectedProduct: " + typicalProduct.innerHTML + "<BR>" + divTypical.innerHTML.replace("AffectedProduct: " + typicalProduct.innerHTML + "<BR>","");
       else if (strField == "AffectedState")
            divTypical.innerHTML = "AffectedState: " + typicalAffectedState.innerHTML + "<BR>" + divTypical.innerHTML.replace("AffectedState: " + typicalAffectedState.innerHTML + "<BR>","");
       else if (strField == "State")
           divTypical.innerHTML = "State: " + typicalState.innerHTML + "<BR>" + divTypical.innerHTML.replace("State: " + typicalState.innerHTML + "<BR>", "");
       else if (strField == "FailedFixes")
           divTypical.innerHTML = "FailedFixes: " + typicalFailedFixes.innerHTML + "<BR>" + divTypical.innerHTML.replace("FailedFixes: " + typicalFailedFixes.innerHTML + "<BR>", "");
       else if (strField == "GatingMilestone")
            divTypical.innerHTML = "GatingMilestone: " + typicalMilestone.innerHTML + "<BR>" + divTypical.innerHTML.replace("GatingMilestone: " + typicalMilestone.innerHTML + "<BR>","");
       else if (strField == "Owner")
            divTypical.innerHTML = "Owner: " + typicalEmail.innerHTML + "<BR>" + divTypical.innerHTML.replace("Owner: " + typicalEmail.innerHTML + "<BR>","");
       else if (strField == "Approver")
            divTypical.innerHTML = "Approver: " + typicalEmail.innerHTML + "<BR>" + divTypical.innerHTML.replace("Approver: " + typicalEmail.innerHTML + "<BR>","");
       else if (strField == "ComponentPM")
            divTypical.innerHTML = "ComponentPM: " + typicalEmail.innerHTML + "<BR>" + divTypical.innerHTML.replace("ComponentPM: " + typicalEmail.innerHTML + "<BR>","");
       else if (strField == "ComponentTestLead")
            divTypical.innerHTML = "ComponentTestLead: " + typicalEmail.innerHTML + "<BR>" + divTypical.innerHTML.replace("ComponentTestLead: " + typicalEmail.innerHTML + "<BR>","");
       else if (strField == "Developer")
            divTypical.innerHTML = "Developer: " + typicalEmail.innerHTML + "<BR>" + divTypical.innerHTML.replace("Developer: " + typicalEmail.innerHTML + "<BR>","");
       else if (strField == "Originator")
            divTypical.innerHTML = "Originator: " + typicalEmail.innerHTML + "<BR>" + divTypical.innerHTML.replace("Originator: " + typicalEmail.innerHTML + "<BR>","");
       else if (strField == "ProductPM")
            divTypical.innerHTML = "ProductPM: " + typicalEmail.innerHTML + "<BR>" + divTypical.innerHTML.replace("ProductPM: " + typicalEmail.innerHTML + "<BR>","");
       else if (strField == "ProductTestLead")
            divTypical.innerHTML = "ProductTestLead: " + typicalEmail.innerHTML + "<BR>" + divTypical.innerHTML.replace("ProductTestLead: " + typicalEmail.innerHTML + "<BR>","");
       else if (strField == "Tester")
            divTypical.innerHTML = "Tester: " + typicalEmail.innerHTML + "<BR>" + divTypical.innerHTML.replace("Tester: " + typicalEmail.innerHTML + "<BR>","");
       else if (strField == "Division")
            divTypical.innerHTML = "Division: " + typicalDivision.innerHTML + "<BR>" + divTypical.innerHTML.replace("Division: " + typicalDivision.innerHTML + "<BR>","");
       else if (strField == "ApproverGroup")
            divTypical.innerHTML = "ApproverGroup: " + typicalGroup.innerHTML + "<BR>" + divTypical.innerHTML.replace("ApproverGroup: " + typicalGroup.innerHTML + "<BR>","");
       else if (strField == "ComponentTestLeadGroup")
            divTypical.innerHTML = "ComponentTestLeadGroup: " + typicalGroup.innerHTML + "<BR>" + divTypical.innerHTML.replace("ComponentTestLeadGroup: " + typicalGroup.innerHTML + "<BR>","");
       else if (strField == "ComponentPMGroup")
            divTypical.innerHTML = "ComponentPMGroup: " + typicalGroup.innerHTML + "<BR>" + divTypical.innerHTML.replace("ComponentPMGroup: " + typicalGroup.innerHTML + "<BR>","");
       else if (strField == "DeveloperGroup")
            divTypical.innerHTML = "DeveloperGroup: " + typicalGroup.innerHTML + "<BR>" + divTypical.innerHTML.replace("DeveloperGroup: " + typicalGroup.innerHTML + "<BR>","");
       else if (strField == "OriginatorGroup")
            divTypical.innerHTML = "OriginatorGroup: " + typicalGroup.innerHTML + "<BR>" + divTypical.innerHTML.replace("OriginatorGroup: " + typicalGroup.innerHTML + "<BR>","");
       else if (strField == "OwnerGroup")
            divTypical.innerHTML = "OwnerGroup: " + typicalGroup.innerHTML + "<BR>" + divTypical.innerHTML.replace("OwnerGroup: " + typicalGroup.innerHTML + "<BR>","");
       else if (strField == "ProductPMGroup")
            divTypical.innerHTML = "ProductPMGroup: " + typicalGroup.innerHTML + "<BR>" + divTypical.innerHTML.replace("ProductPMGroup: " + typicalGroup.innerHTML + "<BR>","");
       else if (strField == "ProductTestLeadGroup")
            divTypical.innerHTML = "ProductTestLeadGroup: " + typicalGroup.innerHTML + "<BR>" + divTypical.innerHTML.replace("ProductTestLeadGroup: " + typicalGroup.innerHTML + "<BR>","");
       else if (strField == "TesterGroup")
            divTypical.innerHTML = "TesterGroup: " + typicalGroup.innerHTML + "<BR>" + divTypical.innerHTML.replace("TesterGroup: " + typicalGroup.innerHTML + "<BR>","");
       else if (strField == "ApprovalCheck")
            divTypical.innerHTML = "ApprovalCheck: " + typicalApprovalCheck.innerHTML + "<BR>" + divTypical.innerHTML.replace("ApprovalCheck: " + typicalApprovalCheck.innerHTML + "<BR>","");
       else if (strField == "OnBoard")
            divTypical.innerHTML = "OnBoard: " + typicalOnBoard.innerHTML + "<BR>" + divTypical.innerHTML.replace("OnBoard: " + typicalOnBoard.innerHTML + "<BR>","");
       else if (strField == "ImplementationCheck")
            divTypical.innerHTML = "ImplementationCheck: " + typicalImplementationCheck.innerHTML + "<BR>" + divTypical.innerHTML.replace("ImplementationCheck: " + typicalImplementationCheck.innerHTML + "<BR>","");
       else if (strField == "ComponentType")
            divTypical.innerHTML = "ComponentType: " + typicalComponentType.innerHTML + "<BR>" + divTypical.innerHTML.replace("ComponentType: " + typicalComponentType.innerHTML + "<BR>","");
       else if (strField == "Impacts")
            divTypical.innerHTML = "Impacts: " + typicalImpacts.innerHTML + "<BR>" + divTypical.innerHTML.replace("Impacts: " + typicalImpacts.innerHTML + "<BR>","");
       else if (strField == "Severity")
            divTypical.innerHTML = "Severity: " + typicalSeverity.innerHTML + "<BR>" + divTypical.innerHTML.replace("Severity: " + typicalSeverity.innerHTML + "<BR>","");
       else if (strField == "EAStatus")
            divTypical.innerHTML = "EAStatus: " + typicalEAStatus.innerHTML + "<BR>" + divTypical.innerHTML.replace("EAStatus: " + typicalEAStatus.innerHTML + "<BR>","");
       else if (strField == "TestEscape")
            divTypical.innerHTML = "TestEscape: " + typicalTestEscape.innerHTML + "<BR>" + divTypical.innerHTML.replace("TestEscape: " + typicalTestEscape.innerHTML + "<BR>","");
       else if (strField == "DaysCurrentOwner")
            divTypical.innerHTML = "DaysCurrentOwner: " + typicalDays.innerHTML + "<BR>" + divTypical.innerHTML.replace("DaysCurrentOwner: " + typicalDays.innerHTML + "<BR>","");
       else if (strField == "DaysInState")
            divTypical.innerHTML = "DaysInState: " + typicalDays.innerHTML + "<BR>" + divTypical.innerHTML.replace("DaysInState: " + typicalDays.innerHTML + "<BR>","");
       else if (strField == "DaysOpen")
            divTypical.innerHTML = "DaysOpen: " + typicalDays.innerHTML + "<BR>" + divTypical.innerHTML.replace("DaysOpen: " + typicalDays.innerHTML + "<BR>","");
       else if (strField == "DateOpened")
            divTypical.innerHTML = "DateOpened: " + typicalDate.innerHTML + "<BR>" + divTypical.innerHTML.replace("DateOpened: " + typicalDate.innerHTML + "<BR>","");
       else if (strField == "DateClosed")
            divTypical.innerHTML = "DateClosed: " + typicalDate.innerHTML + "<BR>" + divTypical.innerHTML.replace("DateClosed: " + typicalDate.innerHTML + "<BR>","");
       else if (strField == "DateModified")
            divTypical.innerHTML = "DateModified: " + typicalDate.innerHTML + "<BR>" + divTypical.innerHTML.replace("DateModified: " + typicalDate.innerHTML + "<BR>","");
       else if (strField == "EADate")
            divTypical.innerHTML = "EADate: " + typicalDate.innerHTML + "<BR>" + divTypical.innerHTML.replace("EADate: " + typicalDate.innerHTML + "<BR>","");
       else if (strField == "TargetDate")
            divTypical.innerHTML = "TargetDate: " + typicalDate.innerHTML + "<BR>" + divTypical.innerHTML.replace("TargetDate: " + typicalDate.innerHTML + "<BR>","");
       else if (strField == "EarliestProductMilestone")
            divTypical.innerHTML = "EarliestProductMilestone: " + typicalDate.innerHTML + "<BR>" + divTypical.innerHTML.replace("EarliestProductMilestone: " + typicalDate.innerHTML + "<BR>","");
       else if (strField == "ShortDescription")
            divTypical.innerHTML = "ShortDescription: " + typicalFreeformString.innerHTML + "<BR>" + divTypical.innerHTML.replace("ShortDescription: " + typicalFreeformString.innerHTML + "<BR>","");
       else if (strField == "LongDescription")
            divTypical.innerHTML = "LongDescription: " + typicalFreeformString.innerHTML + "<BR>" + divTypical.innerHTML.replace("LongDescription: " + typicalFreeformString.innerHTML + "<BR>","");
       else if (strField == "Updates")
            divTypical.innerHTML = "Updates: " + typicalFreeformString.innerHTML + "<BR>" + divTypical.innerHTML.replace("Updates: " + typicalFreeformString.innerHTML + "<BR>","");
       else if (strField == "StepsToReproduce")
            divTypical.innerHTML = "StepsToReproduce: " + typicalFreeformString.innerHTML + "<BR>" + divTypical.innerHTML.replace("StepsToReproduce: " + typicalFreeformString.innerHTML + "<BR>","");
       else if (strField == "CustomerImpact")
            divTypical.innerHTML = "CustomerImpact: " + typicalFreeformString.innerHTML + "<BR>" + divTypical.innerHTML.replace("CustomerImpact: " + typicalFreeformString.innerHTML + "<BR>","");
       else if (strField == "Localization")
            divTypical.innerHTML = "Localization: " + typicalFreeformString.innerHTML + "<BR>" + divTypical.innerHTML.replace("Localization: " + typicalFreeformString.innerHTML + "<BR>","");
       else if (strField == "ClosedInVersion")
            divTypical.innerHTML = "ClosedInVersion: " + typicalString.innerHTML + "<BR>" + divTypical.innerHTML.replace("ClosedInVersion: " + typicalString.innerHTML + "<BR>","");
       else if (strField == "ComponentVersion")
            divTypical.innerHTML = "ComponentVersion: " + typicalString.innerHTML + "<BR>" + divTypical.innerHTML.replace("ComponentVersion: " + typicalString.innerHTML + "<BR>","");
       else if (strField == "EANumber")
            divTypical.innerHTML = "EANumber: " + typicalString.innerHTML + "<BR>" + divTypical.innerHTML.replace("EANumber: " + typicalString.innerHTML + "<BR>","");
       else if (strField == "SAPartNumber")
            divTypical.innerHTML = "SAPartNumber: " + typicalString.innerHTML + "<BR>" + divTypical.innerHTML.replace("SAPartNumber: " + typicalString.innerHTML + "<BR>","");
       else if (strField == "SupplierVersion")
            divTypical.innerHTML = "SupplierVersion: " + typicalString.innerHTML + "<BR>" + divTypical.innerHTML.replace("SupplierVersion: " + typicalString.innerHTML + "<BR>","");
       else if (strField == "TestProcedure")
            divTypical.innerHTML = "TestProcedure: " + typicalString.innerHTML + "<BR>" + divTypical.innerHTML.replace("TestProcedure: " + typicalString.innerHTML + "<BR>","");
       else if (strField == "ReferenceNumber")
            divTypical.innerHTML = "ReferenceNumber: " + typicalString.innerHTML + "<BR>" + divTypical.innerHTML.replace("ReferenceNumber: " + typicalString.innerHTML + "<BR>","");
       else if (strField == "ReleaseFixImplemented")
            divTypical.innerHTML = "ReleaseFixImplemented: " + typicalString.innerHTML + "<BR>" + divTypical.innerHTML.replace("ReleaseFixImplemented: " + typicalString.innerHTML + "<BR>","");
       else if (strField == "LastReleaseTested")
            divTypical.innerHTML = "LastReleaseTested: " + typicalString.innerHTML + "<BR>" + divTypical.innerHTML.replace("LastReleaseTested: " + typicalString.innerHTML + "<BR>","");
       else if (strField == "ObservationID")
            divTypical.innerHTML = "ObservationID: " + typicalNumber.innerHTML + "<BR>" + divTypical.innerHTML.replace("ObservationID: " + typicalNumber.innerHTML + "<BR>","");
       else if (strField == "SubSystem")
            divTypical.innerHTML = "SubSystem: " + typicalSubSystem.innerHTML + "<BR>" + divTypical.innerHTML.replace("SubSystem: " + typicalSubSystem.innerHTML + "<BR>","");
       else if (strField == "Feature")
            divTypical.innerHTML = "Feature: " + typicalFeature.innerHTML + "<BR>" + divTypical.innerHTML.replace("Feature: " + typicalFeature.innerHTML + "<BR>","");
       else if (strField == "Frequency")
            divTypical.innerHTML = "Frequency: " + typicalFrequency.innerHTML + "<BR>" + divTypical.innerHTML.replace("Frequency: " + typicalFrequency.innerHTML + "<BR>","");
       else if (strField == "Component")
            divTypical.innerHTML = "Component: " + typicalComponent.innerHTML + "<BR>" + divTypical.innerHTML.replace("Component: " + typicalComponent.innerHTML + "<BR>","");
       else if (strField == "CoreTeam")
            divTypical.innerHTML = "CoreTeam: " + typicalCoreTeam.innerHTML + "<BR>" + divTypical.innerHTML.replace("CoreTeam: " + typicalCoreTeam.innerHTML + "<BR>","");
    
    
    
    }

    function insertAtCursor(myField, myValue) {
        //IE support
        if (document.selection) {
            myField.focus();
            sel = document.selection.createRange();
            sel.text = myValue;
        }
        //MOZILLA/NETSCAPE support
        else if (myField.selectionStart || myField.selectionStart == '0') {
            var startPos = myField.selectionStart;
            var endPos = myField.selectionEnd;
            myField.value = myField.value.substring(0, startPos)
                        + myValue
                        + myField.value.substring(endPos, myField.value.length);
            myField.selectionStart = endPos + myField.value.length;
        } else {
            myField.value += myValue;
        }
    }


    function InsertExamples(ExampleID){
        if (ExampleID == 1)
            insertAtCursor(frmMain.txtAdvanced," DateOpened between \'1/12/2009\' and \'11/1/2010\' ");
        else if (ExampleID == 2)
            insertAtCursor(frmMain.txtAdvanced," Owner in ('my.name@hp.com', 'your.name@hp.com') ");
        else if (ExampleID == 3)
            insertAtCursor(frmMain.txtAdvanced," Priority in (1,2) ");
        else if (ExampleID == 4)
            insertAtCursor(frmMain.txtAdvanced," ComponentType <> 'HW' ");
        else if (ExampleID == 5)
            insertAtCursor(frmMain.txtAdvanced," GatingMilestone is null ");
        else if (ExampleID == 6)
            insertAtCursor(frmMain.txtAdvanced," ComponentName like '%Video%' ");
        else if (ExampleID == 7)
            insertAtCursor(frmMain.txtAdvanced," ProductFamily not like 'Func Tst%' ");
            
    
    }

    function LookupValues(strField){
        var strResult;

        strResult = window.showModalDialog("BuildSQLLookup.asp?txtField=" + strField , "", "dialogWidth:300px;dialogHeight:350px;edge: Raised;center:Yes; help: No;resizable: yes;status: No");
        if (typeof (strResult) != "undefined") {
            insertAtCursor(frmMain.txtAdvanced,strResult);
        }
    }

    function window_onload(){
        frmMain.txtAdvanced.focus();
        if (typeof(window.dialogArguments) != "undefined")
            frmMain.txtAdvanced.value = window.dialogArguments.frmMain.txtAdvanced.value;
        else if (typeof( window.parent.opener) != "undefined")
            frmMain.txtAdvanced.value = window.parent.opener.frmMain.txtAdvanced.value;
    }

    function cmdOK_onclick(){
        var AffectedProductFound = frmMain.txtAdvanced.value.indexOf("AffectedProduct");
        var AffectedStateFound = frmMain.txtAdvanced.value.indexOf("AffectedState");
        var blnOK = true;

        if (AffectedProductFound != -1 || AffectedStateFound != -1)
            blnOK = confirm("Warning: Specifying Affected Products or Affected States in this field can result in single observations being counted or displayed more than once in status reports, graphs, charts, and other observation reports.\r\rIf you would like to filter by Affected Product or Affected State without duplicate observations, please use the Affected Product and Affected State fields on the main screen instead of this field.\r\rAre you sure you want to save this filter?")

        if (blnOK)
            {
            frmMain.cmdOK.disabled=true;
            frmMain.cmdCancel.disabled=true;
            var i=0;
            var t;
            frmMain.action = "../ots/Report.asp";
            frmMain.submit();
            GetSyntaxStatus();
       
            SyntaxStatus = 0
            }
    }
        function GetSyntaxStatus(){
            if (frmMain.txtAdvanced.value=="")
                {
               // frmMain.txtAdvanced.value = frmMain.txtAdvanced.value.replace(/\"/g,"'");
  		        window.returnValue=frmMain.txtAdvanced.value;
                if(navigator.appVersion.indexOf("Chrome") > -1)
                    {
                    if (typeof( window.parent.opener) != "undefined")
                        window.parent.opener.frmMain.txtAdvanced.value = frmMain.txtAdvanced.value;
   		            }
                top.close();
                }
            else if( CheckSyntax.document.body.innerText.indexOf("SyntaxOK")==-1 && CheckSyntax.document.body.innerText.indexOf("SyntaxError:")==-1)
                setTimeout("GetSyntaxStatus()",1000);
            else if (CheckSyntax.document.body.innerText.indexOf("SyntaxOK")!=-1)
                {
                frmMain.txtAdvanced.value = frmMain.txtAdvanced.value.replace(/\"/g,"'");
  		        window.returnValue=frmMain.txtAdvanced.value;
                if(navigator.appVersion.indexOf("Chrome") > -1)
                    {
                    if (typeof( window.parent.opener) != "undefined")
                        window.parent.opener.frmMain.txtAdvanced.value = frmMain.txtAdvanced.value;
                    }
   		        top.close();
                }
            else
                {
                alert(CheckSyntax.document.body.innerText);
                frmMain.cmdOK.disabled=false;
                frmMain.cmdCancel.disabled=false;
                frmMain.txtAdvanced.focus();
                }
            CheckSyntax.document.body.innerText = "";
}
//-->
</SCRIPT>

<style>
    body
    {
        font-size: x-small;
        font-family: Verdana
    }
    textarea
    {
        font-size: xx-small;
        font-family: Verdana
    }
    td
    {
        font-size: xx-small;
        font-family: Verdana
    }
    divFieldName
    {
        font-weight:bold;
    }
A:visited
{
    COLOR: blue
}
A:hover
{
    COLOR: red
}    
</style>
</HEAD>
<BODY bgcolor=lavender onload="window_onload();">
<font size=2><b>Enter Other Criteria</b></font><br>
<BR>
<%


    dim cn
    dim rs
    dim p
    dim cm


    set cn = server.CreateObject("ADODB.Connection")
    cn.ConnectionString = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
    cn.Open

    set rs = server.CreateObject("ADODB.recordset")
%>
    <form id=frmMain target=CheckSyntax method=post>
    <table width=100% cellpadding=2 cellspacing=0>
        <tr>
            <td nowrap valign=top>
                <b>FieldNames</b><br>
                <div style="background-color: white;margin-bottom:4px;margin-top:2px; border: solid 1px steelblue;height:200px;overflow-y:scroll;"  >
                <%
                    dim strField 

                    for each strField in split("AffectedProduct,AffectedState,ApprovalCheck,Approver,ApproverGroup,ClosedInVersion,Component,ComponentPM,ComponentPMGroup,ComponentTestLead,ComponentTestLeadGroup,ComponentType,ComponentVersion,CoreTeam,CustomerImpact,DateOpened,Date Closed,Date Modified,Days Current Owner,Days In State,Days Open,Developer,Developer Group,Division,EA Date,EA Number,EA Status,Earliest Product Milestone,Failed Fixes,Feature,Frequency,Gating Milestone,Impacts,Implementation Check,Last Release Tested,Localization,Long Description,Observation ID,On Board,Originator,Originator Group,Owner,Owner Group,Priority,Primary Product,Product Family,Product PM,Product PM Group,Product Test Lead,Product Test Lead Group,Reference Number,Release Fix Implemented,SA Part Number,Severity,Short Description,State,Status,StepsToReproduce,Sub System,Supplier Version,Target Date,Test Escape,Test Procedure,Tester,Tester Group,Updates",",")
                        response.write "- <a href=""javascript: AddField('" & replace(strField," ","") & "');"">" & replace(strField," ","") & "</a><BR>"
                    next
                %>
                </div>
                <b>Examples</b><br>
                <div style="margin-top:2px;background-color: white; border: solid 1px steelblue;height:100px"  >
                - <a href="javascript: InsertExamples(1);">Opened between two dates</a><br>
                - <a href="javascript: InsertExamples(2);">Owned by list of people</a><br>
                - <a href="javascript: InsertExamples(3);">Priority 1 or 2</a><br>
                - <a href="javascript: InsertExamples(4);">Not hardware deliverables</a><br>
                - <a href="javascript: InsertExamples(5);">No GatingMilestone assigned</a><br>
                - <a href="javascript: InsertExamples(6);">"Video" in component name</a><br>
                - <a href="javascript: InsertExamples(7);">Not "Functional Test"</a><br>
                </div>
            </td>
            <td valign=top width=100%>
                <b>Filter Critera</b><br>
                <textarea id="txtAdvanced" name="txtAdvanced" style="margin-bottom:4px;border: solid 1px steelblue;height:200px;width:100%" rows=18></textarea><br>
                <b>Typical Values</b><br>
                <div id=divTypicalValues style="margin-top:2px;background-color: white; border: solid 1px steelblue;height:100px;overflow-y:scroll;"  >
                
                    
                        <div id=divTypical></div>
                </div>            
            </td>
        </tr>
    </table>

<!--Typical Values go here-->

<%
    dim strProductList
    dim strProductFamilyList

    strProductList = ""
    strProductFamilyList = ""
    strSQl = "Select top 5 Platform_Cycle_Name,Platform_Cycle_Version " & _
             "from HOUSIREPORT01.SIO.dbo.product with (NOLOCK) " & _
             "where Source_System_ID=0 " & _
             "and Platform_Cycle_Name not like '%-%' " & _
             "and RTM_Date is not null " & _
             "order by create_date desc"

    rs.open strSQL,cn,adOpenKeyset
    do while not rs.EOF
        if strProductFamilyList = "" then
            strProductFamilyList = "<a href=""javascript: AddValue('\'" & rs("Platform_Cycle_Name") & "\'');"">'" & rs("Platform_Cycle_Name") & "'</a>"
        else
            strProductFamilyList = strProductFamilyList & ", <a href=""javascript: AddValue('\'" & rs("Platform_Cycle_Name") & "\'');"">'" & rs("Platform_Cycle_Name") & "'</a>"
        end if
        if strProductList = "" then
            strProductList = "<a href=""javascript: AddValue('\'" & rs("Platform_Cycle_Version") & "\'');"">'" & rs("Platform_Cycle_Version") & "'</a>"
        else
            strProductList = strProductList & ", <a href=""javascript: AddValue('\'" & rs("Platform_Cycle_Version") & "\'');"">'" & rs("Platform_Cycle_Version") & "'</a>"
        end if

        rs.MoveNext
    loop
    rs.Close





    set rs=nothing
    cn.Close
    set cn = nothing

%>
<hr>
<table BORDER="0" CELLSPACING="1" CELLPADDING="1" align="right">
	<tr>
		<TD><INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()"></TD>
		<TD><INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel  LANGUAGE=javascript onclick="return window.parent.close();"  ></TD>
</TR></table>

<div style="display:none" id=typicalStatus><a href="javascript: AddValue('\'Open\'');">'Open'</a>, <a href="javascript: AddValue('\'Closed\'');">'Closed'</a>, <a href="javascript: AddValue('\'Pending EA\'');">'Pending EA'</a></div>
<div style="display:none" id=typicalPriority><a href="javascript: AddValue('0');">0</a>, <a href="javascript: AddValue('1');">1</a>, <a href="javascript: AddValue('2');">2</a>, <a href="javascript: AddValue('3');">3</a></div>
<div style="display:none" id=typicalFailedFixes><a href="javascript: AddValue('0');">0</a>, <a href="javascript: AddValue('1');">1</a>, <a href="javascript: AddValue('2');">2</a>, <a href="javascript: AddValue('3');">3</a></div>
<div style="display:none" id=typicalProduct><%=strProductList%>, <a href="javascript: LookupValues('Product')"><b>Lookup</b></a></div>
<div style="display:none" id=typicalProductFamily><%=strProductFamilyList%>, <a href="javascript: LookupValues('ProductFamily')"><b>Lookup</b></a></div>
<div style="display:none" id=typicalState><a href="javascript: AddValue('\'Fix Verified\'');">'Fix Verified'</a>, <a href="javascript: AddValue('\'Under Investigation\'');">'Under Investigation'</a>, <a href="javascript: AddValue('\'Will Not Fix\'');">'Will Not Fix'</a>, <a href="javascript: AddValue('\'Cannot Duplicate\'');">'Cannot Duplicate'</a>, <a href="javascript: LookupValues('State')"><b>Lookup</b></a></div>
<div style="display:none" id=typicalAffectedState><a href="javascript: AddValue('\'Untested*\'');">'Untested*'</a>, <a href="javascript: AddValue('\'Affected\'');">'Affected'</a>, <a href="javascript: AddValue('\'Fix Verified\'');">'Fix Verified'</a>, <a href="javascript: AddValue('\'Not Affected\'');">'Not Affected'</a>, <a href="javascript: AddValue('\'Will Not Fix\'');">'Will Not Fix'</a>, <a href="javascript: LookupValues('AffectedState')"><b>Lookup</b></a></div>
<div style="display:none" id=typicalEmail>An email address enclosed in single quotes.  Example: 'your.name@hp.com' <a href="javascript: LookupValues('Email')"><b>Lookup</b></a></div>
<div style="display:none" id=typicalMilestone><a href="javascript: AddValue('\'\'');">Not Specified</a>, <a href="javascript: AddValue('\'RTM 1\'');">'RTM 1'</a>, <a href="javascript: AddValue('\'RTM 2\'');">'RTM 2'</a>, <a href="javascript: AddValue('\'SI 1\'');">'SI 1'</a>, <a href="javascript: AddValue('\'Functional Test\'');">'Functional Test'</a>, <a href="javascript: AddValue('\'Web Release\'');">'Web Release'</a>, <a href="javascript: LookupValues('GatingMilestone')"><b>Lookup</b></a></div>
<div style="display:none" id=typicalDivision><a href="javascript: AddValue('\'Mobile\'');">'Mobile'</a>, <a href="javascript: AddValue('\'Desktops\'');">'Desktops'</a>, <a href="javascript: AddValue('\'Workstations\'');">'Workstations'</a></div>
<div style="display:none" id=typicalGroup><a href="javascript: AddValue('\'CNB_Software PM\'');">'CNB_Software PM'</a>, <a href="javascript: AddValue('\'BNB_Communications\'');">'BNB_Communications'</a>, <a href="javascript: AddValue('\'DSO_Thin_Clients\'');">'DSO_Thin_Clients'</a>, <a href="javascript: LookupValues('Group')"><b>Lookup</b></a></div>
<div style="display:none" id=typicalOnBoard><a href="javascript: AddValue('is null');">Not Assigned</a>, <a href="javascript: AddValue('1');">Yes</a>, <a href="javascript: AddValue('0');">No</a></div>
<div style="display:none" id=typicalApprovalCheck><a href="javascript: AddValue('is null');">Not Assigned</a>, <a href="javascript: AddValue('\'Author Approved\'');">'Author Approved'</a>, <a href="javascript: AddValue('\'PM Approved\'');">'PM Approved'</a>, <a href="javascript: AddValue('\'Specialist Approved\'');">'Specialist Approved'</a>, <a href="javascript: AddValue('\'Supplier Approved\'');">'Supplier Approved'</a></div>
<div style="display:none" id=typicalImplementationCheck><a href="javascript: AddValue('is null');">Not Assigned</a>, <a href="javascript: AddValue('\'Author Approved\'');">'Author Approved'</a>, <a href="javascript: AddValue('\'PM Approved\'');">'PM Approved'</a>, <a href="javascript: AddValue('\'Waiting For Author\'');">'Waiting For Author'</a></div>
<div style="display:none" id=typicalComponentType><a href="javascript: AddValue('\'SW\'');">'SW'</a>, <a href="javascript: AddValue('\'HW\'');">'HW'</a>, <a href="javascript: AddValue('\'FW\'');">'FW'</a>, <a href="javascript: AddValue('\'Image\'');">'Image'</a>, <a href="javascript: AddValue('\'CD/Docs\'');">'CD/Docs'</a>, <a href="javascript: AddValue('\'Certification\'');">'Certification'</a>, <a href="javascript: AddValue('\'Factory\'');">'Factory'</a>, <a href="javascript: AddValue('\'Softpaq\'');">'Softpaq'</a></div>
<div style="display:none" id=typicalImpacts><a href="javascript: AddValue('is null');">Not Assigned</a>, <a href="javascript: AddValue('\'Customer\'');">'Customer'</a>, <a href="javascript: AddValue('\'Factory\'');">'Factory'</a></div>
<div style="display:none" id=typicalSeverity><a href="javascript: AddValue('1');">1 - Critical</a>, <a href="javascript: AddValue('2');">2 - Serious</a>, <a href="javascript: AddValue('3');">3 - Medium</a>, <a href="javascript: AddValue('4');">4 - Low</a></div>
<div style="display:none" id=typicalEAStatus><a href="javascript: AddValue('is null');">Not Assigned</a>, <a href="javascript: AddValue('\'Not Required\'');">'Not Required'</a>, <a href="javascript: AddValue('\'Required\'');">'Required'</a>, <a href="javascript: AddValue('\'Complete\'');">'Complete'</a>, <a href="javascript: AddValue('\'Submitted\'');">'Submitted'</a></div>
<div style="display:none" id=typicalTestEscape><a href="javascript: AddValue('is null');">Not Assigned</a>, <a href="javascript: AddValue('\'Functional\'');">'Functional'</a>, <a href="javascript: AddValue('\'Integration\'');">'Integration'</a>, <a href="javascript: AddValue('\'Unit\'');">'Unit'</a></div>
<div style="display:none" id=typicalDays>A number of days. Examples:  <a href="javascript: AddValue('<7');">Less Than 7 Days</a>, <a href="javascript: AddValue('> 14');">More Than 14 Days</a>, <a href="javascript: AddValue('between 7 and 14');">Between 7 and 14 Days</a> </div>
<div style="display:none" id=typicalDate>A date or date range. Examples:  <a href="javascript: AddValue('is null');">Not Assigned</a>, <a href="javascript: AddValue('\'1/1/2010\'');">'1/1/2010'</a>, <a href="javascript: AddValue('Between \'1/1/2010\' and \'1/31/2010\'');">Between '1/1/2010' and '1/31/2010'</a> </div>
<div style="display:none" id=typicalFreeformString>Use the "Text Search" function on the main screen if possible. Otherwise, use <a href="javascript: AddValue('like \'%TEXTTOFIND%\'');">Like</a> to find a substring.</div>
<div style="display:none" id=typicalString>A text string enclosed in single quotes. Example: <a href="javascript: AddValue('\'TEXTTOFIND\'');">'TEXTTOFIND'</a>. Or, use <a href="javascript: AddValue('like \'%TEXTTOFIND%\'');">Like</a> to find a substring.</div>
<div style="display:none" id=typicalNumber>A number.</div>
<div style="display:none" id=typicalSubSystem><a href="javascript: AddValue('\'Application\'');">'Application'</a>, <a href="javascript: AddValue('\'Application - Common\'');">'Application - Common'</a>, <a href="javascript: AddValue('\'System BIOS\'');">'System BIOS'</a>, <a href="javascript: AddValue('\'Driver\'');">'Driver'</a>, <a href="javascript: AddValue('\'HW/System\'');">'HW/System'</a>, <a href="javascript: LookupValues('SubSystem')"><b>Lookup</b></a></div>
<div style="display:none" id=typicalFeature><a href="javascript: AddValue('\'\'');">None</a>, <a href="javascript: AddValue('\'Firmware\'');">'Firmware'</a>, <a href="javascript: AddValue('\'Application General\'');">'Application General'</a>, <a href="javascript: AddValue('\'Localization\'');">'Localization'</a>, <a href="javascript: AddValue('\'Networking\'');">'Networking'</a>, <a href="javascript: AddValue('\'Installation\'');">'Installation'</a>, <a href="javascript: AddValue('\'Browser\'');">'Browser'</a>, <a href="javascript: LookupValues('Feature')"><b>Lookup</b></a></div>
<div style="display:none" id=typicalFrequency><a href="javascript: AddValue('\'Always: 100%\'');">'Always: 100%'</a>, <a href="javascript: AddValue('\'Intermittent: 25-99%\'');">'Intermittent: 25-99%'</a>, <a href="javascript: AddValue('\'Intermittent: 5-25%\'');">'Intermittent: 5-25%'</a>, <a href="javascript: AddValue('\'Seen Once\'');">'Seen Once'</a>, <a href="javascript: AddValue('\'Intermittent: 1-5%\'');">'Intermittent: 1-5%'</a>, <a href="javascript: LookupValues('Frequency')"><b>Lookup</b></a></div>
<div style="display:none" id=typicalComponent><a href="javascript: AddValue('\'Mechanical\'');">'Mechanical'</a>, <a href="javascript: AddValue('\'HP QuickWeb\'');">'HP QuickWeb'</a>, <a href="javascript: AddValue('\'System Board\'');">'System Board'</a>, <a href="javascript: AddValue('\'System BIOS\'');">'System BIOS'</a>, <a href="javascript: AddValue('\'Unknown\'');">'Unknown'</a>, <a href="javascript: AddValue('\'Image\'');">'Image'</a>, <a href="javascript: LookupValues('Component')"><b>Lookup</b></a></div>
<div style="display:none" id=typicalCoreTeam><a href="javascript: AddValue('\'None\'');">'None'</a>, <a href="javascript: AddValue('\'Multimedia Apps\'');">'Multimedia Apps'</a>, <a href="javascript: AddValue('\'BIOS\'');">'BIOS'</a>, <a href="javascript: AddValue('\'Value Add SW\'');">'Value Add SW'</a>, <a href="javascript: AddValue('\'Preinstall\'');">'Preinstall'</a>, <a href="javascript: AddValue('\'Security Apps\'');">'Security Apps'</a>, <a href="javascript: LookupValues('CoreTeam')"><b>Lookup</b></a></div>
    
    <input style="display:none" id="txtObservationID" name="txtObservationID" type="text" value="1"/>
    <input style="display:none" id="txtReportSections" name="txtReportSections" type="text" value="-1"/>
    </form>
    <iframe style="display:none" id=CheckSyntax name=CheckSyntax width=100% height=120px></iframe>

</BODY>
</HTML>
