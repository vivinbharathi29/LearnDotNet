<%@ Language=VBScript %>
<% Option Explicit %>
<!-- #include file="../../includes/DataWrapper.asp" -->
<%
		
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"

  Dim AppRoot, strRowBorderColor
  strRowBorderColor="Gainsboro"
  AppRoot = Session("ApplicationRoot")
    Dim strImageActionItemID : strImageActionItemID = Request("ImageActionItemID")
    Dim strActionType : strActionType = Request("ActionType")
    Dim strCurrentUser : strCurrentUser = Request("CurrentUser")
    Dim strProductVersionID : strProductVersionID = Request("ProductVersionID")

    Dim dw, cn, cmd, rs
    Set dw = New DataWrapper
    set rs = server.CreateObject("ADODB.recordset")
    Set cn = dw.CreateConnection("PDPIMS_ConnectionString")
    Set cmd = dw.CreateCommAndSP(cn, "usp_Image_ListImageDefinitionBrandsAll")	
    dw.CreateParameter cmd, "@p_intProductID", adInteger, adParamInput, 8, clng(request("ProductVersionID"))
    Set rs = dw.ExecuteCommAndReturnRS(cmd)
    
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<STYLE>
td{
	FONT-FAMILY:Verdana;
	FONT-SIZE:x-small;
}
</STYLE>
    <script src="<%= AppRoot %>/includes/client/jquery.min.js" type="text/javascript"></script>
    <script src="<%= AppRoot %>/includes/client/jquery-ui.min.js" type="text/javascript"></script>
    <script type="text/javascript" src="../../Scripts/PulsarPlus.js"></script>
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

function cmdCancel_onclick() {
   
    if (IsFromPulsarPlus()) {
        ClosePulsarPlusPopup();
    }
    else {
        window.parent.CloseImagePopUp(false, 0);
    }
}

function cmdOK_onclick() {
    var ActionType = document.getElementById("txtActionType").value;
    var ImageActionItemID = document.getElementById("txtImageActionItemID").value;
    var CurrentUser = document.getElementById("txtCurrentUser").value;
    var ProductBrandID = 0;
    var ProductVersionID = document.getElementById("txtProductVersionID").value;
    //0 add, 1 obso
    var parameters = "";
    var request;

    $('input:checkbox[name="chkSCM"]:checked').each(function () {
        ProductBrandID = $(this).val();
        var parameters = "Function=UpdateAV&ImageActionItemID=" + ImageActionItemID + "&ActionType=" + ActionType + "&CurrentUserName=" + CurrentUser + "&ProductVersionID=" + ProductVersionID + "&ProductBrandID=" + ProductBrandID;
        var request = null;
        //Initialize the AJAX variable.
        if (window.XMLHttpRequest) {// Are we working with mozilla
            request = new XMLHttpRequest(); //Yes -- this is mozilla.
        } else { //Not Mozilla, must be IE
            request = new ActiveXObject("Microsoft.XMLHTTP");
        } //End setup Ajax.
        request.open("POST", "ImageTabChangeActions.asp", false);
        request.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
        request.send(parameters);
    });

    parameters = "Function=UpdateAV&ImageActionItemID=" + ImageActionItemID + "&ActionType=3&CurrentUserName=" + CurrentUser + "&ProductVersionID=0&ProductBrandID=0";
    request = null;
    //Initialize the AJAX variable.
    if (window.XMLHttpRequest) {// Are we working with mozilla
        request = new XMLHttpRequest(); //Yes -- this is mozilla.
    } else { //Not Mozilla, must be IE
        request = new ActiveXObject("Microsoft.XMLHTTP");
    } //End setup Ajax.
    request.open("POST", "ImageTabChangeActions.asp", false);
    request.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
    request.send(parameters);
   
    if (IsFromPulsarPlus()) {
        window.parent.parent.parent.popupCallBack(ImageActionItemID);
        ClosePulsarPlusPopup();
    }
    else {
        window.parent.CloseImagePopUp(true, ImageActionItemID);
    }
}

function chkAllSCMs_onclick()
{
    var i;
    var checkBoxes = document.getElementsByTagName("input");
    var chkAllSCMs, chkBoxName;
    chkAllSCMs = document.getElementById("chkAllSCMs");
    chkBoxName = "chkSCM";
    for (i = 0; i < checkBoxes.length; i++) {
        if (checkBoxes[i].name == chkBoxName) {
            checkBoxes[i].checked = chkAllSCMs.checked;
        }
    }
}

//-->
</SCRIPT>
</HEAD>

<BODY bgcolor=Ivory LANGUAGE=javascript>

<form ID=frmImageTabChangesSCMList method=post>
<%
If Not (rs.Bof And rs.Eof) Then
%>
    <TABLE width=100%>
    <TR><TD style="font-family:Verdana; text-align:left"><b>Select required SCMs and click save</b></td></tr>    
    </table>
    <br />
    <table id="tblImageChangeSCMList" border="0" width="100%" cellspacing="0" cellpadding="2">
    <thead>
        <tr bgcolor="beige">
            <td valign="top" style="width: 8%;">
                <input id="chkAllSCMs" type="checkbox" style="height: 16px; width: 16px" onclick="javascript: chkAllSCMs_onclick();">
            </td>
            <td style="width: 92%;"><font size="1"><b>SCM</b></font></td>
        </tr>
    </thead>
    <%
	do while not rs.EOF  
    %>
        <tr bgcolor="ivory" id="trImageChangeSCM<%=rs("CombinedProductBrandId")%>">
            <td valign="top" style="width: 8%; BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK: KEEP-ALL">
                <input id="<%=rs("CombinedProductBrandId")%>" name="chkSCM" type="checkbox" style="height: 16px; width: 16px" value="<%=rs("CombinedProductBrandId")%>">
            </td>
            <td style="width: 92%; BORDER-TOP: <%=strRowBorderColor%> thin solid; WORD-BREAK: KEEP-ALL" class="cell" valign="top" nowrap>
                <font class="text" size="1"><%= rs("Brand")%></font>
            </td>
        </tr>
    <%
    rs.MoveNext
	loop	
    %>
    </table>
    <br />
    <br />
    <TABLE width=100%><TR><TD align=right><INPUT type="button" value="Cancel" id=cmdCancel name=cmdCancel LANGUAGE=javascript onclick="return cmdCancel_onclick()">&nbsp;<INPUT type="button" value="OK" id=cmdOK name=cmdOK LANGUAGE=javascript onclick="return cmdOK_onclick()"></td></tr></table>
    <input id="txtActionType" name="txtActionType" type="hidden" value="<%=strActionType%>" />
    <input id="txtImageActionItemID" name="txtImageActionItemID" type="hidden" value="<%=strImageActionItemID%>" />
    <input id="txtCurrentUser" name="txtCurrentUser" type="hidden" value="<%=strCurrentUser%>" />
    <input id="txtProductVersionID" name="txtProductVersionID" type="hidden" value="<%=strProductVersionID%>" />
<% 
End If 
rs.Close
%>
</form>
</BODY>
</HTML>
