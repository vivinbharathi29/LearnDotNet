<%@ Language=VBScript %>
<html>
<head>
<meta name="VI60_DefaultClientScript" content="JavaScript">

<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

<script ID="clientEventHandlersVBS" LANGUAGE="javascript">

    var SelectedAvCategory;
    var SelectedBusinessID;
    var SelectedPartNo;
    var SelectedRegions;
    var iStep;

    function cmdNext_onclick() {
        if (window.parent.frames["UpperWindow"].iStep == 1) { //AV Selection
            window.parent.frames["UpperWindow"].ValidateAvCategorySelection();
            if (window.parent.frames["UpperWindow"].SelectedAvCategory > "") {
                SelectedAvCategory = window.parent.frames["UpperWindow"].SelectedAvCategory;
                window.parent.frames["UpperWindow"].iStep = 2;
                iStep = 2;
                window.parent.frames["UpperWindow"].ProcessStep(); 
                return true; }
            else {
                alert("Please Select An AV Category");
                return false;
            }
        }
   
        if (window.parent.frames["UpperWindow"].iStep == 2) { //PartNo Selection
            window.parent.frames["UpperWindow"].ValidatePartNoSelection();
            if (window.parent.frames["UpperWindow"].SelectedPartNo > "") {
                SelectedPartNo = window.parent.frames["UpperWindow"].SelectedPartNo;
                SelectedBusinessID = window.parent.frames["UpperWindow"].SelectedBusinessID;
                window.parent.frames["UpperWindow"].location = "SubAssemblyMain.asp?SAType=<%=Request("SAType")%>&Step=3&FeatureCategoryID=" + SelectedAvCategory + "&SubassemblyID=" + SelectedPartNo + "&BusinessID=" + SelectedBusinessID;
                window.parent.frames["UpperWindow"].SelectedPartNo = SelectedPartNo;
                window.parent.frames["UpperWindow"].sSubassemblyID = SelectedPartNo;
                window.parent.frames["UpperWindow"].SelectedAvCategory = SelectedAvCategory;
                window.parent.frames["UpperWindow"].frmAddNew.sAvFeatureCategoryID = SelectedAvCategory;
                cmdPrevious.disabled = false;
                cmdNext.disabled = true;
                cmdFinish.disabled = false;
                iStep = 3;
                return true; }
            else {
                alert("Please Select A Part Number");
                return false;
            }
        }
    }

    function cmdPrevious_onclick() {
        if (iStep == 2) {  //PartNo Selection
            window.parent.frames["UpperWindow"].PopulateAvFeatureCategortySelection(SelectedAvCategory)
            window.parent.frames["UpperWindow"].iStep = 1;
            iStep = 1;
            window.parent.frames["UpperWindow"].ProcessStep();
            return true;
        }
        if  (iStep == 3){ //Region(s) Selection
        
            window.parent.frames["UpperWindow"].PopulatePartNoSelection(SelectedPartNo)
            window.parent.frames["UpperWindow"].iStep = 2;
            iStep = 2;
            window.parent.frames["UpperWindow"].ProcessStep();
            return true;
        }
    }

    function cmdCancel_onclick(pulsarplusDivId) {
        if (window.confirm ("Are you sure you want to exit this screen without saving your changes?") == true)
            if (pulsarplusDivId != undefined && pulsarplusDivId != "") {
                // For Closing current popup if Called from pulsarplus
                parent.window.parent.closeExternalPopup();
            }
            else{
                window.parent.close();
            }
    }

    function cmdFinish_onclick() {
        window.parent.frames["UpperWindow"].ValidateRegionsSelection();
        //if (window.parent.frames["UpperWindow"].SelectedRegions > "") {
        SelectedRegions = window.parent.frames["UpperWindow"].SelectedRegions;
        window.parent.frames["UpperWindow"].frmAddNew.hidFunction.value = "save";
        window.parent.frames["UpperWindow"].frmAddNew.submit();
        cmdPrevious.disabled = true;
        cmdNext.disabled = true;
        cmdFinish.disabled = true;
        return true;
        //}
        //else {
        //    alert("Please Select Region(s)");
        //    return false;
        //}
    }


</SCRIPT>
</head>
<body bgcolor="ivory" LANGUAGE="javascript">
<table BORDER="0" CELLSPACING="1" CELLPADDING="1" align="right">
	<tr>
	    <%If Request("Existing") <> "" Then %>
			<td><input type="button" value="Cancel" id="cmdCancel" name="cmdCancel" LANGUAGE="javascript" onclick="return cmdCancel_onclick('<%=Request("pulsarplusDivId")%>')"></td>
			<td width="10"></td>
			<td><input type="button" value="&lt;&lt; Previous" id="cmdPrevious" disabled name="cmdPrevious" LANGUAGE="javascript" onclick="return cmdPrevious_onclick()" disabled></td>
			<td><input type="button" value="Next&gt;&gt;" id="cmdNext" name="cmdNext" disabled LANGUAGE="javascript" onclick="return cmdNext_onclick()"></td>
			<td width="10"></td>
			<td><input type="button" value="Finish" id="cmdFinish" name="cmdFinish"  LANGUAGE="javascript" onclick="return cmdFinish_onclick()"></td>
	    <%Else%>
	        <td><input type="button" value="Cancel" id="cmdCancel" name="cmdCancel" LANGUAGE="javascript" onclick="return cmdCancel_onclick('<%=Request("pulsarplusDivId")%>')"></td>
			<td width="10"></td>
			<td><input type="button" value="&lt;&lt; Previous" id="cmdPrevious" name="cmdPrevious" LANGUAGE="javascript" onclick="return cmdPrevious_onclick()" disabled></td>
			<td><input type="button" value="Next&gt;&gt;" id="cmdNext" name="cmdNext" LANGUAGE="javascript" onclick="return cmdNext_onclick()"></td>
			<td width="10"></td>
			<td><input type="button" value="Finish" id="cmdFinish" name="cmdFinish" disabled LANGUAGE="javascript" onclick="return cmdFinish_onclick()"></td>
	    <%End If%>
	</tr>
</table>
</body>
</html>