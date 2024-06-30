<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<HTML>
<HEAD>
<script language="JavaScript" src="../../_ScriptLibrary/jsrsClient.js"></script>

<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">

<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--
    function window_onload() {

    }

    function cmdProdAdd_onclick() {
        var i;
        for (i = 0; i < frmJupiterXLR8Report.lstAvailableProd.length; i++) {
            if (frmJupiterXLR8Report.lstAvailableProd.options[i].selected) {
                frmJupiterXLR8Report.lstSelectedProd.options[frmJupiterXLR8Report.lstSelectedProd.length] = new Option(frmJupiterXLR8Report.lstAvailableProd.options[i].text, frmJupiterXLR8Report.lstAvailableProd.options[i].value);
            }
        }
        for (i = frmJupiterXLR8Report.lstAvailableProd.length - 1; i >= 0; i--) {
            if (frmJupiterXLR8Report.lstAvailableProd.options[i].selected)
                frmJupiterXLR8Report.lstAvailableProd.options[i] = null;
        }
    }

    function cmdProdAddAll_onclick() {
        var i;

        for (i = 0; i < frmJupiterXLR8Report.lstAvailableProd.length; i++) {
            frmJupiterXLR8Report.lstSelectedProd.options[frmJupiterXLR8Report.lstSelectedProd.length] = new Option(frmJupiterXLR8Report.lstAvailableProd.options[i].text, frmJupiterXLR8Report.lstAvailableProd.options[i].value);
        }
        for (i = frmJupiterXLR8Report.lstAvailableProd.length - 1; i >= 0; i--) {
            frmJupiterXLR8Report.lstAvailableProd.options[i] = null;
        }

    }

    function cmdProdRemove_onclick() {
        var i;

        for (i = 0; i < frmJupiterXLR8Report.lstSelectedProd.length; i++) {
            if (frmJupiterXLR8Report.lstSelectedProd.options[i].selected) {
                frmJupiterXLR8Report.lstAvailableProd.options[frmJupiterXLR8Report.lstAvailableProd.length] = new Option(frmJupiterXLR8Report.lstSelectedProd.options[i].text, frmJupiterXLR8Report.lstSelectedProd.options[i].value);
            }
        }
        for (i = frmJupiterXLR8Report.lstSelectedProd.length - 1; i >= 0; i--) {
            if (frmJupiterXLR8Report.lstSelectedProd.options[i].selected)
                frmJupiterXLR8Report.lstSelectedProd.options[i] = null;
        }

    }

    function cmdProdRemoveAll_onclick() {
        var i;

        for (i = 0; i < frmJupiterXLR8Report.lstSelectedProd.length; i++) {
            frmJupiterXLR8Report.lstAvailableProd.options[frmJupiterXLR8Report.lstAvailableProd.length] = new Option(frmJupiterXLR8Report.lstSelectedProd.options[i].text, frmJupiterXLR8Report.lstSelectedProd.options[i].value);
        }
        for (i = frmJupiterXLR8Report.lstSelectedProd.length - 1; i >= 0; i--)
            frmJupiterXLR8Report.lstSelectedProd.options[i] = null;

    }

    function lstAvailableProd_ondblclick() {
        if (frmJupiterXLR8Report.lstAvailableProd.options.selectedIndex > -1) {
            frmJupiterXLR8Report.lstSelectedProd.options[frmJupiterXLR8Report.lstSelectedProd.length] = new Option(frmJupiterXLR8Report.lstAvailableProd.options[frmJupiterXLR8Report.lstAvailableProd.options.selectedIndex].text, frmJupiterXLR8Report.lstAvailableProd.options[frmJupiterXLR8Report.lstAvailableProd.options.selectedIndex].value);
            frmJupiterXLR8Report.lstAvailableProd.options[frmJupiterXLR8Report.lstAvailableProd.options.selectedIndex] = null;
        }

    }

    function lstSelectedProd_ondblclick() {
        if (frmJupiterXLR8Report.lstSelectedProd.options.selectedIndex > -1) {
            frmJupiterXLR8Report.lstAvailableProd.options[frmJupiterXLR8Report.lstAvailableProd.length] = new Option(frmJupiterXLR8Report.lstSelectedProd.options[frmJupiterXLR8Report.lstSelectedProd.options.selectedIndex].text, frmJupiterXLR8Report.lstSelectedProd.options[frmJupiterXLR8Report.lstSelectedProd.options.selectedIndex].value);
            frmJupiterXLR8Report.lstSelectedProd.options[frmJupiterXLR8Report.lstSelectedProd.options.selectedIndex] = null;
        }
    }

    function GetSelectedProducts() {
        var strAddProducts;
        strAddProducts = "";

        for (i = 0; i < frmJupiterXLR8Report.lstAvailableProd.length; i++) {
            strAddProducts = strAddProducts + frmJupiterXLR8Report.lstAvailableProd.options(i).value + ",";
        }

        if (strAddProducts == "") {
            strAddProducts = ","
        }
        
        //alert(strAddProducts);
        var parameters = "function=UpdateJupiterXLR8ReportProducts&Products=" + strAddProducts;
        var request = null;
        //Initialize the AJAX variable.
        if (window.XMLHttpRequest) {// Are we working with mozilla
            request = new XMLHttpRequest(); //Yes -- this is mozilla.
        } else { //Not Mozilla, must be IE
            request = new ActiveXObject("Microsoft.XMLHTTP");
        } //End setup Ajax.
        request.open("POST", "UpdateJupiterXLR8ReportProducts.asp", false);
        request.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
        request.send(parameters);
        //alert(request.responseText);
        window.close();
    }

//-->
</SCRIPT>
</HEAD>

<BODY bgcolor=Ivory LANGUAGE=javascript onload="return window_onload()">
<link href="style/wizard%20style.css" type="text/css" rel="stylesheet">
<%
	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Application("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")

	'Load Supported Products
	rs.Open "usp_SelectJupiterXLR8ReportProducts",cn,adOpenForwardOnly
	strProductList = rs("Setting")
    rs.Close
%>

<form id="frmJupiterXLR8Report" method="post" action="JupiterXLR8ReportExclusionsSave.asp?">
<font size=3 face=verdana><b>Jupiter XLR8 Report Exclusions - Product List</b></font><BR><BR>
<table ID="ReportExclusions" WIDTH="550px" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
  <tr>
	<td colspan="10" width="100%">
        <table cellSpacing="1" cellPadding="1" width="460" border="0">
		<tr ID="ProdRow1" style="Display:<%=strDisplayProduct%>"><td><strong>Include</strong></td><td></td><td><strong>Exclude</strong></td><tr>       
        <tr ID="ProdRow2" style="Display:<%=strDisplayProduct%>">
          <td width="220"><select id="lstAvailableProd" style="WIDTH: 270px; HEIGHT: 650px" size="2" name="lstAvailableProd" LANGUAGE="javascript" ondblclick="return lstAvailableProd_ondblclick()" multiple> 
			<%
			strSQL = "usp_SelectProductsByBrand"
			rs.Open strSQL,cn,adOpenForwardOnly
			strUNSelectedProducts = ""
			do while not rs.EOF
				if instr("," & strProductList,"," & rs("ID") & ",") <> 0 then
					if rs("Division") & "" = "1" then
						Response.Write "<Option value=" & rs("ID") & ">" & rs("FullProdName") & "</OPTION>"
					end if
				else
					strUNSelectedProducts = strUNSelectedProducts &  "<Option value=" & rs("ID") & ">" & rs("FullProdName") & "</OPTION>"
					strProductsLoaded = strProductsLoaded & rs("ID") & ","
				end if
				rs.MoveNext
			loop
			rs.Close
			%>
            </select></td>
          <td width="20"><input id="cmdProdADD" style="WIDTH: 25px" type="button" width="25" value="&gt;" name="cmdProdADD" LANGUAGE="javascript" onclick="return cmdProdAdd_onclick()"><br><input <%=strProductButtonStatus%> id="cmdProdRemove" style="WIDTH: 25px" type="button" width="20" value="&lt;" name="cmdProdRemove" LANGUAGE="javascript" onclick="return cmdProdRemove_onclick()"><br><br><input id="cmdProdAddAll" type="button" value="&gt;&gt;" name="cmdProdAddAll" LANGUAGE="javascript" style="WIDTH: 25px" onclick="return cmdProdAddAll_onclick()"><br><input id="cmdProdRemoveAll" <%=strProductButtonStatus%> type="button" value="&lt;&lt;" name="cmdProdRemoveAll" LANGUAGE="javascript" style="WIDTH: 25px" onclick="return cmdProdRemoveAll_onclick()"></td>
          <td width="220"><select id="lstSelectedProd" style="WIDTH: 270px; HEIGHT: 650px" size="2" name="lstSelectedProd" multiple  LANGUAGE="javascript" ondblclick="return lstSelectedProd_ondblclick()" > 
              <%=strUNSelectedProducts%></select>
          </td>
      </tr>
      </table>
    </td>
  </tr>
</table>

<label id="lblProductsLoaded" style="Display:none"><%response.write strProductsLoaded%></label>

</form>
	<%
	set rs=nothing
	cn.Close
	set cn=nothing
	%>
</BODY>
</HTML>
