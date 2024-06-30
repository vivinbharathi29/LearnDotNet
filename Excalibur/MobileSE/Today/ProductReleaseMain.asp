<%@ Language=VBScript %>
<%	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"	
    
  Dim AppRoot
  AppRoot = Session("ApplicationRoot")
      
%>	
<HTML>
<HEAD>
    <TITLE>Select Product (Release)</TITLE>
    <META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
    <!-- #include file="../../includes/bundleConfig.inc" -->
    <SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

function cmdCancel_onclick() {
    if (parent.window.parent.document.getElementById('modal_dialog')) {
        parent.window.parent.modalDialog.cancel();
    } else {
        window.parent.close();
    }
}

function cmdOK_onclick() {
    if (Reassign.txtNotes.value == "" && Reassign.chkUrgent.checked) {
        alert("You must enter notes for Urgent requests.");
        Reassign.txtNotes.focus();
        return;
    }
	Reassign.txtName.value = Reassign.cboOwner.options[Reassign.cboOwner.selectedIndex].text;
	Reassign.submit();
}

function AddRelease(ID, TypeID, BSID) {	
    var sWidth = $(window).width() * 90 / 100;
    var sHeight = $(window).height() * 90 / 100;
    modalDialog.open({ dialogTitle: 'Add New Product Release', dialogURL: '../../Release/Release.asp?ID=' + ID + '&ProductTypeID=' + TypeID + '&BusinessSegmentID=' + BSID, dialogHeight: sHeight, dialogWidth: sWidth, dialogResizable: true, dialogDraggable: true });
}

function EditDisclaimerNotes(ReleaseID) {
    var sWidth = $(window).width() * 90 / 100;
    var sHeight = $(window).height() * 90 / 100;
    modalDialog.open({ dialogTitle: 'Cycle Disclaimer Notes', dialogURL: '../../Release/DisclaimerNotes.asp?ReleaseID=' + ReleaseID, dialogHeight: sHeight, dialogWidth: sWidth, dialogResizable: true, dialogDraggable: true });
}

// check if the release is used in an AV or PRL before removing it from the product
function ChangeRelease(me) {
    if ($("#isClone").val() == 1) {
        return false;
    }

    var ReleaseID = me.value.split('-')[0];
    var ProductVersionID = document.getElementById('ID').value;

    ajaxurl = "ProductReleaseCheckRelease.asp?ReleaseID=" + ReleaseID + "&ProductVersionID=" + ProductVersionID;
    $.ajax({
        url: ajaxurl,
        type: "POST",
        success: function (data) {
            if (data != "")
            {
                alert("Release cannot be removed from this Product since it is being used in " + data);
                me.checked = true;
            }
        },
        error: function (xhr, status, error) {
            alert("error is " + error);
        }
    });

}

function SelectLeadProductRelease(ReleaseID, LeadProductVersionReleaseID) {
    $("#txtSelectedRelease").val(ReleaseID);
    modalDialog.open({ dialogTitle: 'Select Lead Product (Release)', dialogURL: '../../../ipulsar/product/ProductReleaseSearch.aspx?LeadProductVersionReleaseID=' + LeadProductVersionReleaseID, dialogHeight: 400, dialogWidth: 400, dialogResizable: true, dialogDraggable: true });
}

function returnValues(value, desc) {
    var ReleaseID = $("#txtSelectedRelease").val();
    $("#chkRelease" + ReleaseID).val(ReleaseID + "-" + value);
    $("#spanLeadProductreleaseDesc" + ReleaseID).html("<a href='javascript: RemoveLeadProductRelease(" + ReleaseID + ");'>Remove</a> | <a href='javascript: SelectLeadProductRelease(" + ReleaseID + "," + value + ");'>Edit</a>  " + desc);
    modalDialog.cancel();
}

function RemoveLeadProductRelease(ReleaseID)
{
    $("#chkRelease" + ReleaseID).val(ReleaseID + "-0");
    $("#spanLeadProductreleaseDesc" + ReleaseID).html("<a href='javascript: SelectLeadProductRelease(" + ReleaseID + ",0);'>Add</a> ");
}

//*****************************************************************
//Description:  Code that runs when page loads
//Function:     window_onload();
//Modified By:  09/13/2016 - Harris, Valerie - PBI 23434/ Task 24367 - Change dialogs to JQuery dialogs     
//*****************************************************************
function window_onload() {
    //Add modal dialog code to body tag: ---
    modalDialog.load();
}

function filterDigital(e, pnumber) {

    if (!/^\d+$/.test(pnumber))     
    {        
        var newValue = /^\d+/.exec(e.value);        
        if (newValue != null)        
        {            
            e.value = newValue;       
        }     
        else    
        {         
            e.value = "";   
        }  
    } 
    return false;
}

$(function () {
    var location  = "";
    try  { 
        location = parent.window.getPageLocation();
    }
    catch(ex)
    { }

    if (location != "ProductProperties")
    {
        $('input[name="chkRelease"]').prop("disabled", true);
        $("#hlAddNew").hide();
    }
});
//-->
</SCRIPT>
</HEAD>
<BODY onload="window_onload();" bgcolor="ivory">
<%
if request("ID") = ""  then
	Response.Write "<BR>&nbsp;Not enough information supplied"
else
	dim cn
	dim rs
	dim cm
	dim p
	dim strID
	dim CurrentUser
	dim CurrentUserID
	
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
	end if
	rs.Close

    dim strLeadProductReleaseList
    strLeadProductReleaseList = "<select id='cboLeadProductRelease' name='cboLeadProductRelease' style='width: 160px;'><option value='0' selected></option>"
    rs.Open "usp_GetProductReleaseList",cn,adOpenForwardOnly
	do while not rs.EOF
		strLeadProductReleaseList = strLeadProductReleaseList & "<OPTION value=" & rs("ID") & ">" & rs("ProductRelease")  & "</OPTION>"
		rs.MoveNext
	loop
	rs.Close
	strLeadProductReleaseList = strLeadProductReleaseList & "</select>"
%>
<link rel="stylesheet" type="text/css" href="../Style/programoffice.css">
<font size=3 face=verdana><b> 
<%
    rs.open "spGetProductVersionName " & clng(request("ID")),cn,adOpenForwardOnly
    if rs.eof and rs.bof then
        response.write "Product"    
    else
        response.write rs("Name")  
    end if
    rs.close
%>
</b></font>
<form ID=frmMain method=post>
<table ID="tabAdd" WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
	<tr>
		<td width="100%"><a id="hlAddNew" href="javascript: AddRelease(<%=request("ID")%>, <%=request("ProductTypeID")%>, <%=request("BusinessSegmentID")%>);">Add New Product Release</a><br /><br />
            <table id="releaseList">
            <thead>
                <tr style="font-weight:bold">
                    <td>Releases</td>
                    <td style="padding-left: .5cm">Lead Product (Release)</td>
                    <td style="padding-left: .5cm">RTM WAVE</td>
                </tr>
            </thead>
        <%  
            dim strName
			dim writeBuffer
            rs.open "usp_ProductVersion_Release " & clng(request("ID")) & "," & clng(request("BusinessSegmentID")), cn, adOpenForwardOnly
            do while not rs.eof
                strname = rs("Name")                
                if trim(rs("OnProduct")) = "1" then
                    writeBuffer = "<input checked id='chkRelease" & rs("ID") & "' name='chkRelease' type='checkbox' value='" & rs("ID") & "-" & rs("LeadProductreleaseID") & "' ReleaseName='" & strname &  "' ReleaseID='" & rs("ID") & "' onclick='ChangeRelease(this);'> "
                else
                    writeBuffer = "<input id='chkRelease" & rs("ID") & "' name='chkRelease' type='checkbox' value='" & rs("ID") & "-0' ReleaseName='" & strName & "' ReleaseID='" & rs("ID") & "'> "
                end if
                               
                writeBuffer = "<tr><td>" & writeBuffer & strname &_ 
                              " <a href=""" & "javascript: EditDisclaimerNotes(" & rs("ID") & ");" & """>Cycle Disclaimer Notes</a>" & "<BR>" &_ 
                              "</td>" & "<td style='padding-left: .5cm'>" & "<span id='spanLeadProductreleaseDesc" & rs("ID") & "'>" 

                if rs("LeadProductreleaseID") = 0 then
                    writeBuffer =  writeBuffer & "<a href='javascript: SelectLeadProductRelease(" & rs("ID") & ",0);'>Add</a>"
                else 
                    writeBuffer =  writeBuffer & "<a href='javascript: RemoveLeadProductRelease(" & rs("ID") & ");'>Remove</a> | " &_ 
                                                 "<a href='javascript: SelectLeadProductRelease(" & rs("ID") & "," & rs("LeadProductreleaseID") & ");'>Edit</a> " & rs("LeadProductreleaseDesc") 
                end if

                writeBuffer = writeBuffer & "</span>" & "</td>" & "<td style='padding-left: .5cm'>"

                ' column Cyc == "00" is mean "NPI", Cyc is from 00 to 12
                if InStr(rs("Cyc"),"00") > 0 then
                    writeBuffer = writeBuffer &_
                                  "<input type='text' id='rtmwave" & rs("ID") & "' value='" & rs("RtmWave") & "' name='rtmwave' size='10' maxlength='5' onkeyup='return filterDigital(this,value)' />" &_
                                  "</td></tr>"
                else
                    writeBuffer = writeBuffer & "</td></tr>"
                end if

                response.write writeBuffer

                rs.movenext
            loop
            rs.close
           
        %>
            </table>
		</td>
	</tr>
</table>
    <input type="hidden" id="ID" name="ID" value="<%=request("ID")%>">
    <input type="hidden" id="isClone" value="<%=request("isClone")%>">
    <input id="txtSelectedRelease" name="txtSelectedRelease" type="hidden" value="0">
</form>
<%

	set rs = nothing
	set cn = nothing
end if


%>

</BODY>
</HTML>