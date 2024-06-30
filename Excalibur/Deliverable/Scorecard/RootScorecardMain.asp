<%@ Language=VBScript %>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

    <script language="javascript" type="text/javascript">
<!--

        function window_onload() {
            if (frmMain.txtActionRequest.value == "1")
                frmMain.txtAction.focus();
        }

        function Button_onclick(strButton, strButtonGroup) {
            var ButtonR;
            var ButtonY;
            var ButtonG;
            var txtStatus;
            var txtText;

            if (strButtonGroup == "1") {
                ButtonR = frmMain.Button1R;
                ButtonY = frmMain.Button1Y;
                ButtonG = frmMain.Button1G;
                txtStatus = frmMain.txtExecutiveSummaryStatus;
                txtText = frmMain.txtExecutiveSummary;
            }
            else if (strButtonGroup == "2") {
                ButtonR = frmMain.Button2R;
                ButtonY = frmMain.Button2Y;
                ButtonG = frmMain.Button2G;
                txtStatus = frmMain.txtHPPeopleProcessStatus;
                txtText = frmMain.txtHPPeopleProcess;
            }
            else if (strButtonGroup == "3") {
                ButtonR = frmMain.Button3R;
                ButtonY = frmMain.Button3Y;
                ButtonG = frmMain.Button3G;
                txtStatus = frmMain.txtHPEquipmentStatus;
                txtText = frmMain.txtHPEquipment;
            }
            else if (strButtonGroup == "4") {
                ButtonR = frmMain.Button4R;
                ButtonY = frmMain.Button4Y;
                ButtonG = frmMain.Button4G;
                txtStatus = frmMain.txtSupplierPeopleProcessStatus;
                txtText = frmMain.txtSupplierPeopleProcess;
            }
            else if (strButtonGroup == "5") {
                ButtonR = frmMain.Button5R;
                ButtonY = frmMain.Button5Y;
                ButtonG = frmMain.Button5G;
                txtStatus = frmMain.txtSupplierDeliverablesStatus;
                txtText = frmMain.txtSupplierDeliverables;
            }
            else if (strButtonGroup == "6") {
                ButtonR = frmMain.Button6R;
                ButtonY = frmMain.Button6Y;
                txtStatus = frmMain.txtActionStatus;
                txtText = frmMain.txtAction;
            }
            ButtonR.style.background = "gainsboro";
            ButtonR.style.color = "black"
            ButtonY.style.background = "gainsboro";
            ButtonY.style.color = "black"
            if (strButtonGroup != "6") {
                ButtonG.style.background = "gainsboro";
                ButtonG.style.color = "black"
            }
            if (strButton == "1") {
                if (txtStatus.value == "1")
                    txtStatus.value = "0"
                else
                    {
                        if (strButtonGroup == "6")
                            ButtonR.style.background = "#9999FF"; //"slateblue"; //"slateblue";
                        else
                            ButtonR.style.background = "#FF9999"; // "red"; //"slateblue";
                        ButtonR.style.color = "black"
                    txtStatus.value = "1"
                    }
            }
            else if (strButton == "2") {
                if (txtStatus.value == "2")
                    txtStatus.value = "0"
                else 
                    {
                        if (strButtonGroup == "6") 
                            {
                            ButtonY.style.background = "#99CC66";"green"; // "slateblue";
                            ButtonY.style.color = "black"
                        }
                        else 
                            {
                                ButtonY.style.background = "#FFFF66"; // "yellow" //"#FFDE00"; // "slateblue";
                            ButtonY.style.color = "black"
                        }
                        txtStatus.value = "2"
                    }
            }
            else 
            {
                if (txtStatus.value == "3")
                    txtStatus.value = "0"
                else 
                    {
                        ButtonG.style.background = "#99CC66"; "#86FF88"; //"green"; //"slateblue";
                        ButtonG.style.color = "black";// "white"
                        txtStatus.value = "3"
                    }
                }
                txtText.focus();
        }


        function Button6_onclick(strButton) {
            frmMain.Button6O.style.background = "gainsboro";
            frmMain.Button6O.style.color = "black"
            frmMain.Button6C.style.background = "gainsboro";
            frmMain.Button6C.style.color = "black"
            if (strButton == "1") {
                frmMain.Button6O.style.background = "slateblue";
                frmMain.Button6O.style.color = "white"
            }
            else {
                frmMain.Button6C.style.background = "slateblue";
                frmMain.Button6C.style.color = "white"
            }
        }


        function Text_onkeyup(strField, intLength,LostFocus) {
            var txtField;
            var txtFieldTemplate;

            //determine the field to update
            if (strField == 1) {
                txtField = frmMain.txtExecutiveSummary;
                txtFieldTemplate = frmMain.txtExecutiveSummaryTemplate;
            }
            else if (strField == 2) {
                txtField = frmMain.txtHPPeopleProcess;
                txtFieldTemplate = frmMain.txtHPPeopleProcessTemplate;
            }
            else if (strField == 3) {
                txtField = frmMain.txtHPEquipment;
                txtFieldTemplate = frmMain.txtHPEquipmentTemplate;
            }
            else if (strField == 4) {
                txtField = frmMain.txtSupplierPeopleProcess
                txtFieldTemplate = frmMain.txtSupplierPeopleProcessTemplate;
            }
            else if (strField == 5) {
                txtField = frmMain.txtSupplierDeliverables
                txtFieldTemplate = frmMain.txtSupplierDeliverablesTemplate;
            }
            else if (strField == 6) {
                txtField = frmMain.txtAction
                txtFieldTemplate = frmMain.txtActionTemplate;
            }
            

            
            //Check Size on Lost Focus
            //            if (txtField.value.length > intLength && LostFocus == 1) {
            //              alert("This field can only be " + intLength + " characters long.  Please remove at least " + (txtField.value.length - intLength) + " characters to continue.");
            //            }


            //Change Template formatting as needed
            if (txtField.value == txtFieldTemplate.value)
                txtField.className = "Template";
            else if (txtField.value == "" && LostFocus=="1") {
                txtField.value = txtFieldTemplate.value;
                txtField.className = "Template";
            }
            else if (txtField.value.length > intLength) {
                txtField.className = "TooLong";
            }
            else
                txtField.className = "";


            
        }
// -->
    </script>
</HEAD>
<STYLE>
TD{
	FONT-FAMILY:Verdana;
	FONT-SIZE:xx-small;
}
INPUT{
	FONT-FAMILY:Verdana;
	FONT-SIZE:xx-small;
}

.Template{
	FONT-FAMILY:Verdana;
	FONT-SIZE:xx-small;
	font-style:italic;
	color:Blue;
}

.TooLong{
	FONT-FAMILY:Verdana;
	FONT-SIZE:xx-small;
	color:Red;
}
</STYLE>
<BODY bgcolor=Ivory onload="window_onload();">
<%
    dim strExecutiveSummary
    dim strHPPeopleProcess
    dim strHPEquipment
    dim strSupplierPeopleProcess
    dim strSupplierDeliverables
    dim strAction
    dim strDeliverable 
    dim strExecutiveSummaryStatus
    dim strHPPeopleProcessStatus
    dim strHPEquipmentStatus
    dim strSupplierPeopleProcessStatus
    dim strSupplierDeliverablesStatus
    dim strActionStatus
    dim strClass
    dim strExecutiveSummaryTemplate
    dim strHPPeopleProcessTemplate
    dim strHPEquipmentTemplate
    dim strSupplierPeopleProcessTemplate
    dim strSupplierDeliverablesTemplate
    dim strActionTemplate
    dim strOnStatus

    strOnStatus = ""

%>
<form ID=frmMain method=post action=RootScorecardSave.asp>
<%
	if trim(request("ID")) = "" then
		Response.Write "Not enough information supplied to process this request."
        strDeliverable = ""
	else

		set cn = server.CreateObject("ADODB.Connection")
		cn.ConnectionString = Session("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Password=dino;Persist Security Info=True;User ID=pdpadmin;Initial Catalog=prs;Data Source=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=KB2;User Id=pdpadmin;PASSWORD=dino;"
		cn.Open
		set rs = server.CreateObject("ADODB.recordset")
        strDeliverable = ""
        rs.open "spGetDeliverableRootName " & clng(request("ID")),cn,adOpenForwardOnly
        if not (rs.eof and rs.bof) then
            strDeliverable = trim(rs("name") & "")
            strOnStatus = trim(rs("ShowOnStatus") & "")
            if cbool(strOnStatus) then
                strOnStatus = " checked "
            else
                strOnStatus = " "
            end if
        end if
        rs.Close

        if strDeliverable <> "" then
            'Define Templates
            strExecutiveSummaryTemplate = "Concise text describing risk to POR. Identify what, who, when, and make specific request for staff action.  Consider Scope, Schedule, & Resources and impact to POR commitments"
            strHPPeopleProcessTemplate = "Meeting effectiveness (right people available, agenda, minutes)" & vbcrlf & "HP documentation, specifications (on-time, sufficient)" & vbcrlf & "HP personnel generally ""on top of things""" 
            strHPEquipmentTemplate = "IURs: Enough quantity, on-time, healthy" & vbcrlf & "BIOS health sufficient" & vbcrlf & "Software image sufficient"
            strSupplierPeopleProcessTemplate = "Meeting effectiveness (right people available, agenda, minutes)" & vbcrlf & "Supplier documentation, specifications" & vbcrlf & "OTS updates on-time and sufficient" & vbcrlf & "Supplier personnel ""on top of things"""
            strSupplierDeliverablesTemplate = "On-time" & vbcrlf & "OTS trend" & vbcrlf & "Risk Assessment"
            strActionTemplate = "Free Flowing text describing deliverable action item assigned by CTL or CTL Execution staff, etc."

            'Save Templates
            response.write "<textarea style=""display:none"" id=""txtExecutiveSummaryTemplate"" name=""txtExecutiveSummaryTemplate"">" & strExecutiveSummaryTemplate & "</textarea>"
            response.write "<textarea style=""display:none"" id=""txtHPPeopleProcessTemplate"" name=""txtHPPeopleProcessTemplate"">" & strHPPeopleProcessTemplate & "</textarea>"
            response.write "<textarea style=""display:none"" id=""txtHPEquipmentTemplate"" name=""txtHPEquipmentTemplate"">" & strHPEquipmentTemplate & "</textarea>"
            response.write "<textarea style=""display:none"" id=""txtSupplierPeopleProcessTemplate"" name=""txtSupplierPeopleProcessTemplate"">" & strSupplierPeopleProcessTemplate & "</textarea>"
            response.write "<textarea style=""display:none"" id=""txtSupplierDeliverablesTemplate"" name=""txtSupplierDeliverablesTemplate"">" & strSupplierDeliverablesTemplate & "</textarea>"
            response.write "<textarea style=""display:none"" id=""txtActionTemplate"" name=""txtActionTemplate"">" & strActionTemplate & "</textarea>"


		    rs.Open "spGetRootScoreCard " & clng(request("ID")),cn,adOpenForwardOnly
            if rs.eof and rs.bof then
                strExecutiveSummary = ""
                strHPPeopleProcess = "" 
                strHPEquipment = ""
                strSupplierPeopleProcess = ""
                strSupplierDeliverables = ""
                strAction = ""
                strExecutiveSummaryStatus = "0"
                strHPPeopleProcessStatus = "0"
                strHPEquipmentStatus = "0"
                strSupplierPeopleProcessStatus = "0"
                strSupplierDeliverablesStatus = "0"
                strActionStatus = "0"
            else
                strExecutiveSummary = rs("ExecutiveSummary") & ""
                strHPPeopleProcess = rs("HPPeopleProcess") & ""
                strHPEquipment = rs("HPEquipment") & ""
                strSupplierPeopleProcess = rs("SupplierPeopleProcess") & ""
                strSupplierDeliverables = rs("SupplierDeliverables") & ""
                strAction = rs("Action") & ""
                strExecutiveSummaryStatus = rs("ExecutiveSummaryStatus") & ""
                strHPPeopleProcessStatus = rs("HPPeopleProcessStatus") & ""
                strHPEquipmentStatus = rs("HPEquipmentStatus") & ""
                strSupplierPeopleProcessStatus = rs("SupplierPeopleProcessStatus") & ""
                strSupplierDeliverablesStatus = rs("SupplierDeliverablesStatus") & ""
                strActionStatus = rs("ActionStatus") & ""
            end if
	        rs.close
        end if	    

		set rs = nothing
		cn.Close
		set cn = nothing
	end if

    dim strShowNonAction
    dim strShowAction
    dim ActionHeight

    if trim(request("Action")) = "1" then
        strShowNonAction = "none"
        strShowAction = ""
        ActionHeight = 10
    else
        strShowNonAction = ""
        strShowAction = ""
        ActionHeight=5
    end if

    if strDeliverable <> "" then
	    response.write "<font face=verdana size=2><b>" & strDeliverable & " Scorecard"
        if trim(request("Action")) = "1" then
            response.write " Actions"
        end if
        response.Write "</b><BR><BR></font>"
%>

    

    <table style="width: 100%" cellpadding=2 cellspacing=0 bordercolor=tan border=1>
        <tr style="display:<%=strShowNonAction%>">
            <td bgcolor=cornsilk valign=top><b>Executive<BR>Summary:</b><br><input id="txtExecutiveSummaryStatus" name="txtExecutiveSummaryStatus" style="width:15px" type="hidden" value="<%=strExecutiveSummaryStatus%>"></td>
            <td bgcolor=cornsilk valign=top>
            <% 
                strOption1Color = "background-color: gainsboro;"
                strOption2Color = "background-color: gainsboro;"
                strOption3Color = "background-color: gainsboro;"

                if trim(strExecutiveSummaryStatus) = "1" then
                    strOption1Color = "background-color: #FF9999; color=black"
                elseif trim(strExecutiveSummaryStatus) = "2" then
                    strOption2Color = "background-color: #FFFF66; color=black"
                elseif trim(strExecutiveSummaryStatus) = "3" then
                    strOption3Color = "background-color: #99CC66; color=black"
                end if            
                if trim(strExecutiveSummary) = "" then
                    strExecutiveSummary = strExecutiveSummaryTemplate
                    strClass = "Template"
                else
                    strClass = ""
                end if  

            %>
                <input id="Button1R" type="button" value="Red" style="width: 70px; <%=strOption1Color%>" onclick="return Button_onclick(1,1)"><br>
                <input id="Button1Y" type="button" value="Yellow" style="width: 70px; <%=strOption2Color%>" onclick="return Button_onclick(2,1)"><br>
                <input id="Button1G" type="button" value="Green" style="width: 70px; <%=strOption3Color%>" onclick="return Button_onclick(3,1)"><br>
            </td>
            <td style="width: 100%">
                <textarea id="txtExecutiveSummary" name="txtExecutiveSummary" style="width: 100%; font-family: Verdana; font-size: xx-small;" rows="5" class="<%=strClass%>" onkeydown="Text_onkeyup(1,400,0);" onkeyup="Text_onkeyup(1,400,0);" onchange="Text_onkeyup(1,400,1);"><%=strExecutiveSummary%></textarea>
            </td>
        </tr>
        <tr style="display:<%=strShowNonAction%>">
            <td bgcolor=cornsilk valign=top><b><u>HP</u><BR>People &<BR>Process:</b><br><input id="txtHPPeopleProcessStatus" name="txtHPPeopleProcessStatus" style="width:15px" type="hidden" value="<%=strHPPeopleProcessStatus%>"></td>
            <td bgcolor=cornsilk valign=top>
            <% 
                strOption1Color = "background-color: gainsboro;"
                strOption2Color = "background-color: gainsboro;"
                strOption3Color = "background-color: gainsboro;"

                if trim(strHPPeopleProcessStatus) = "1" then
                    strOption1Color = "background-color: #FF9999; color=black"
                elseif trim(strHPPeopleProcessStatus) = "2" then
                    strOption2Color = "background-color: #FFFF66; color=black"
                elseif trim(strHPPeopleProcessStatus) = "3" then
                    strOption3Color = "background-color: #99CC66; color=black"
                end if            
                if trim(strHPPeopleProcess) = "" then
                    strHPPeopleProcess = strHPPeopleProcessTemplate
                    strClass = "Template"
                else
                    strClass = ""
                end if  

            %>
                <input id="Button2R" type="button" value="Red" style="width: 70px; <%=strOption1Color%>" onclick="return Button_onclick(1,2)"><br>
                <input id="Button2Y" type="button" value="Yellow" style="width: 70px; <%=strOption2Color%>" onclick="return Button_onclick(2,2)"><br>
                <input id="Button2G" type="button" value="Green" style="width: 70px; <%=strOption3Color%>" onclick="return Button_onclick(3,2)"><br>
            </td>
            <td style="width: 100%">
                <textarea id="txtHPPeopleProcess" name="txtHPPeopleProcess" style="width: 100%; font-family: Verdana; font-size: xx-small;" rows="5" class="<%=strClass%>"  onkeydown="Text_onkeyup(2,400,0);" onkeyup="Text_onkeyup(2,400,0);" onchange="Text_onkeyup(2,400,1);"><%=strHPPeopleProcess%></textarea>
            </td>
        </tr>
        <tr style="display:<%=strShowNonAction%>">
            <td bgcolor=cornsilk valign=top><b><u>HP</u><BR>Equipment:</b><br><input id="txtHPEquipmentStatus" name="txtHPEquipmentStatus" style="width:15px" type="hidden" value="<%=strHPEquipmentStatus%>"></td>
            <td  bgcolor=cornsilk valign=top>
            <% 
                strOption1Color = "background-color: gainsboro;"
                strOption2Color = "background-color: gainsboro;"
                strOption3Color = "background-color: gainsboro;"

                if trim(strHPEquipmentStatus) = "1" then
                    strOption1Color = "background-color: #FF9999; color=black"
                elseif trim(strHPEquipmentStatus) = "2" then
                    strOption2Color = "background-color: #FFFF66; color=black"
                elseif trim(strHPEquipmentStatus) = "3" then
                    strOption3Color = "background-color: #99CC66; color=black"
                end if            
                if trim(strHPEquipment) = "" then
                    strHPEquipment = strHPEquipmentTemplate
                    strClass = "Template"
                else
                    strClass = ""
                end if  

            %>
                <input id="Button3R" type="button" value="Red" style="width: 70px; <%=strOption1Color%>" onclick="return Button_onclick(1,3)"><br>
                <input id="Button3Y" type="button" value="Yellow" style="width: 70px; <%=strOption2Color%>" onclick="return Button_onclick(2,3)"><br>
                <input id="Button3G" type="button" value="Green" style="width: 70px; <%=strOption3Color%>" onclick="return Button_onclick(3,3)"><br>
            </td>
            <td style="width: 100%">
                <textarea id="txtHPEquipment" name="txtHPEquipment" style="width: 100%; font-family: Verdana; font-size: xx-small;" rows="5" class="<%=strClass%>" onkeydown="Text_onkeyup(3,400,0);" onkeyup="Text_onkeyup(3,400,0);" onchange="Text_onkeyup(3,400,1);"><%=strHPEquipment%></textarea>
            </td>
        </tr>
        <tr style="display:<%=strShowNonAction%>">
            <td bgcolor=cornsilk valign=top><b><u>Supplier</u><BR>People &<BR>Process:</b><br><input id="txtSupplierPeopleProcessStatus" name="txtSupplierPeopleProcessStatus" style="width:15px" type="hidden" value="<%=strSupplierPeopleProcessStatus%>"></td>
            <td  bgcolor=cornsilk valign=top>
            <% 
                strOption1Color = "background-color: gainsboro;"
                strOption2Color = "background-color: gainsboro;"
                strOption3Color = "background-color: gainsboro;"

                if trim(strSupplierPeopleProcessStatus) = "1" then
                    strOption1Color = "background-color: #FF9999; color=black"
                elseif trim(strSupplierPeopleProcessStatus) = "2" then
                    strOption2Color = "background-color: #FFFF66; color=black"
                elseif trim(strSupplierPeopleProcessStatus) = "3" then
                    strOption3Color = "background-color: #99CC66; color=black"
                end if            
                if trim(strSupplierPeopleProcess) = "" then
                    strSupplierPeopleProcess = strSupplierPeopleProcessTemplate
                    strClass = "Template"
                else
                    strClass = ""
                end if  
            %>
                <input id="Button4R" type="button" value="Red" style="width: 70px; <%=strOption1Color%>" onclick="return Button_onclick(1,4)"><br>
                <input id="Button4Y" type="button" value="Yellow" style="width: 70px; <%=strOption2Color%>" onclick="return Button_onclick(2,4)"><br>
                <input id="Button4G" type="button" value="Green" style="width: 70px; <%=strOption3Color%>" onclick="return Button_onclick(3,4)"><br>
            </td>
            <td style="width: 100%">
                <textarea id="txtSupplierPeopleProcess" name="txtSupplierPeopleProcess" style="width: 100%; font-family: Verdana; font-size: xx-small;" rows="5" class="<%=strClass%>" onkeydown="Text_onkeyup(4,400,0);" onkeyup="Text_onkeyup(4,400,0);" onchange="Text_onkeyup(4,400,1);"><%=strSupplierPeopleProcess%></textarea>
            </td>
        </tr>
        <tr style="display:<%=strShowNonAction%>">
            <td bgcolor=cornsilk valign=top><b><u>Supplier</u><BR>Software:</b><br><input id="txtSupplierDeliverablesStatus" name="txtSupplierDeliverablesStatus" style="width:15px" type="hidden" value="<%=strSupplierDeliverablesStatus%>"></td>
            <td  bgcolor=cornsilk valign=top>
            <% 
                strOption1Color = "background-color: gainsboro;"
                strOption2Color = "background-color: gainsboro;"
                strOption3Color = "background-color: gainsboro;"

                if trim(strSupplierDeliverablesStatus) = "1" then
                    strOption1Color = "background-color: #FF9999; color=black"
                elseif trim(strSupplierDeliverablesStatus) = "2" then
                    strOption2Color = "background-color: #FFFF66; color=black"
                elseif trim(strSupplierDeliverablesStatus) = "3" then
                    strOption3Color = "background-color: #99CC66; color=black"
                end if            
                if trim(strSupplierDeliverables) = "" then
                    strSupplierDeliverables = strSupplierDeliverablesTemplate
                    strClass = "Template"
                else
                    strClass = ""
                end if  
            %>
                <input id="Button5R" type="button" value="Red" style="width: 70px; <%=strOption1Color%>" onclick="return Button_onclick(1,5)"><br>
                <input id="Button5Y" type="button" value="Yellow" style="width: 70px; <%=strOption2Color%>" onclick="return Button_onclick(2,5)"><br>
                <input id="Button5G" type="button" value="Green" style="width: 70px; <%=strOption3Color%>" onclick="return Button_onclick(3,5)"><br>
            </td>
            <td style="width: 100%">
                <textarea id="txtSuplierDeliverables" name="txtSupplierDeliverables" style="width: 100%; font-family: Verdana; font-size: xx-small;" rows="5" class="<%=strClass%>"  onkeydown="Text_onkeyup(5,400,0);" onkeyup="Text_onkeyup(5,400,0);" onchange="Text_onkeyup(5,400,1);"><%=strSupplierDeliverables%></textarea>
            </td>
        </tr>
        <tr style="display:<%=strShowAction%>">
            <td bgcolor=cornsilk valign=top><b>Actions:</b><br><input id="txtActionStatus" name="txtActionStatus" style="width:15px" type="hidden" value="<%=strActionStatus%>"></td>
            <td  bgcolor=cornsilk valign=top>
           <% 
                strOption1Color = "background-color: gainsboro;"
                strOption2Color = "background-color: gainsboro;"

                if trim(strActionStatus) = "1" then
                    strOption1Color = "background-color: #9999FF; color=black"
                elseif trim(strActionStatus) = "2" then
                    strOption2Color = "background-color: #99CC66; color=black"
                end if            
                if trim(strAction) = "" then
                    strAction = strActionTemplate
                    strClass = "Template"
                else
                    strClass = ""
                end if  
            %>
                <input id="Button6R" type="button" value="Open" style="width: 70px; <%=strOption1Color%>" onclick="return Button_onclick(1,6)"><br>
                <input id="Button6Y" type="button" value="Closed" style="width: 70px; <%=strOption2Color%>" onclick="return Button_onclick(2,6)"><br>
            </td>
            <td style="width: 100%">
                <textarea id="txtAction" name="txtAction" style="width: 100%; font-family: Verdana; font-size: xx-small;" rows="<%=ActionHeight%>" class="<%=strClass%>" onkeydown="Text_onkeyup(6,800,0);" onkeyup="Text_onkeyup(6,800,0);" onchange="Text_onkeyup(6,800,1);"><%=strAction%></textarea>
            </td>
        </tr>
        <tr style="display:<%=strShowNonAction%>">
            <td bgcolor=cornsilk valign=top><b>Status:</b></td>
            <td bgcolor=cornsilk valign=top colspan=2>
                <input id="chkStatusReport" name="chkStatusReport" <%=strOnStatus%> type="checkbox" value="1" /> Include this deliverable on Core Team status reports.
            </td>
        </tr>
    </table>

    <%end if%>
    <input id="txtID" name="txtID" type="hidden" value="<%=trim(request("ID"))%>">
    <input id="txtActionRequest" name="txtActionRequest" type="hidden" value="<%=trim(request("Action"))%>">
</FORM>
</BODY>
</HTML>
