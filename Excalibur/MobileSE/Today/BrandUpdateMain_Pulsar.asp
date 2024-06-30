<%@ Language=VBScript %>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
<SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
<!--

var KeyString = "";

function combo_onkeypress() {
	if (event.keyCode == 13)
		{
		KeyString = "";
		frmEmployee.submit();
		}
	else
		{
		KeyString=KeyString+ String.fromCharCode(event.keyCode);
		event.keyCode = 0;
		var i;
		var regularexpression;
		
		for (i=0;i<event.srcElement.length;i++)
			{
				regularexpression = new RegExp("^" + KeyString,"i")
				if (regularexpression.exec(event.srcElement.options[i].text)!=null)
					{
					event.srcElement.selectedIndex = i;
					};
				
			}
		return false;
		}	
}

function combo_onfocus() {
	KeyString = "";
}

function combo_onclick() {
	KeyString = "";
}

function combo_onkeydown() {
	if (event.keyCode==8)
		{
		if (String(KeyString).length >0)
			KeyString= Left(KeyString,String(KeyString).length-1);
		return false;
		}
}

function Left(str, n)
    {
	if (n <= 0)     // Invalid bound, return blank string
		return "";
    else if (n > String(str).length)   // Invalid bound, return
        return str;                // entire string
    else // Valid bound, return appropriate substring
        return String(str).substring(0,n);
    }


function cmdCancel_onclick() {
    if (parent.window.parent.document.getElementById('modal_dialog')) {
        parent.window.parent.modalDialog.cancel();
    } else {
        window.close();
    }
}


function cmdOK_onclick() {
	frmMain.submit();
}

function window_onload() {
    if (frmMain.all.length !=0)
	    frmMain.cboNew.focus();
}

function PopulateSeriesBoxes(){

    if (typeof(event.srcElement.options[event.srcElement.selectedIndex].Series1) == "undefined")
        {
        frmMain.tagSeriesID1.value = "";
        frmMain.txtSeries1.value = "";
        }
    else
        {
        frmMain.txtSeries1.value = event.srcElement.options[event.srcElement.selectedIndex].Series1; 
        frmMain.tagSeriesID1.value = event.srcElement.options[event.srcElement.selectedIndex].SeriesID1; 
        }
    if (typeof(event.srcElement.options[event.srcElement.selectedIndex].Series2) == "undefined")
        {
        frmMain.tagSeriesID2.value = "";
        frmMain.txtSeries2.value = "";
        }
    else
        {
        frmMain.txtSeries2.value = event.srcElement.options[event.srcElement.selectedIndex].Series2; 
        frmMain.tagSeriesID2.value = event.srcElement.options[event.srcElement.selectedIndex].SeriesID2; 
        }

    if (typeof(event.srcElement.options[event.srcElement.selectedIndex].Series3) == "undefined")
        {
        frmMain.tagSeriesID3.value = "";
        frmMain.txtSeries3.value = "";
        }
    else
        {
        frmMain.txtSeries3.value = event.srcElement.options[event.srcElement.selectedIndex].Series3; 
        frmMain.tagSeriesID3.value = event.srcElement.options[event.srcElement.selectedIndex].SeriesID3; 
        }
    if (typeof(event.srcElement.options[event.srcElement.selectedIndex].Series4) == "undefined")

        {
        frmMain.tagSeriesID4.value = "";
        frmMain.txtSeries4.value = "";
        }
    else
        {
        frmMain.txtSeries4.value = event.srcElement.options[event.srcElement.selectedIndex].Series4; 
        frmMain.tagSeriesID4.value = event.srcElement.options[event.srcElement.selectedIndex].SeriesID4; 
        }
    if (typeof(event.srcElement.options[event.srcElement.selectedIndex].Series5) == "undefined")

        {
        frmMain.tagSeriesID5.value = "";
        frmMain.txtSeries5.value = "";
        }
    else
        {
        frmMain.txtSeries5.value = event.srcElement.options[event.srcElement.selectedIndex].Series5; 
        frmMain.tagSeriesID5.value = event.srcElement.options[event.srcElement.selectedIndex].SeriesID5; 
        }
    if (typeof(event.srcElement.options[event.srcElement.selectedIndex].Series6) == "undefined")

        {
        frmMain.tagSeriesID6.value = "";
        frmMain.txtSeries6.value = "";
        }
    else
        {
        frmMain.txtSeries6.value = event.srcElement.options[event.srcElement.selectedIndex].Series6; 
        frmMain.tagSeriesID6.value = event.srcElement.options[event.srcElement.selectedIndex].SeriesID6; 
        }
        
    frmMain.tagSeries1.value = frmMain.txtSeries1.value;
    frmMain.tagSeries2.value = frmMain.txtSeries2.value;
    frmMain.tagSeries3.value = frmMain.txtSeries3.value;
    frmMain.tagSeries4.value = frmMain.txtSeries4.value;
    frmMain.tagSeries5.value = frmMain.txtSeries5.value;
    frmMain.tagSeries6.value = frmMain.txtSeries6.value;
}

//-->
</SCRIPT>
</HEAD>
<STYLE>
td{
	FONT-FAMILY:Verdana;
	FONT-SIZE:x-small;
}
body{
	FONT-FAMILY:Verdana;
	FONT-SIZE:x-small;
}
</STYLE>
<BODY bgcolor=Ivory LANGUAGE=javascript onload="return window_onload()">
<form ID=frmMain action="BrandUpdateSave.asp" method=post>
<% 
    dim strSelected 
    dim i
    dim strExcludeBrands
    dim strProduct
    dim strBrand
    
    strSelected=""
    strAvailable=""
    strSeries = ""
    strBrand = ""
    strExcludeBrands = "," & replace(request("ExcludeIDList")," ","") & ","
    
    if request("ProductID") = "" or not isnumeric(request("ProductID")) or request("BrandID") = "" or not isnumeric(request("BrandID")) or trim(request("ExcludeIDList"))="" then
        response.Write "Unable to process request"
    else

    	set cn = server.CreateObject("ADODB.Connection")
    	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
    	cn.Open
	
    	set rs = server.CreateObject("ADODB.recordset")
    	
    	rs.open "spGetProductVersionName " & clng(request("ProductID")), cn,adOpenStatic
        if rs.eof and rs.bof then
            strProduct = ""
        else
            strProduct = rs("Name") & ""
        end if
    	rs.close

    	rs.open "spGetProductVersionName " & clng(request("ProductID")), cn,adOpenStatic
        if rs.eof and rs.bof then
            strProduct = ""
        else
            strProduct = rs("Name") & ""
        end if
    	rs.close

    	rs.open "spGetBrandbyID " & clng(request("BrandID")), cn,adOpenStatic
        if rs.eof and rs.bof then
            strBrand = ""
        else
            strBrand = rs("Name") & ""
        end if
    	rs.close

        rs.Open "spListBrands_Pulsar " & clng(request("BSegmentID")),cn,adOpenStatic
        do while not rs.EOF
            if rs("active") and instr(strExcludeBrands,"," & rs("ID") & ",") = 0 then
                 strAvailable = strAvailable & "<option value=""" & rs("ID") & """>" & rs("Name") & "</option>"
            end if
            rs.Movenext
        loop
        rs.Close
        
%>

<font size=3 face=verdana><b>Update Product Branding</b></font><br /><br />

<TABLE width=100% bgcolor=cornsilk border=1 bordercolor=tan cellspacing=0 cellpadding=1>
<TR>
	<TD><b>Product:</b>&nbsp;&nbsp;</TD>
	<TD><%=strProduct %>&nbsp;</TD>
</TR>
<TR>
	<TD><b>Old&nbsp;Brand:</b>&nbsp;&nbsp;</TD>
	<TD><%=strBrand%>&nbsp;</TD>
	
</TR>
<TR>
	<TD><b>New&nbsp;Brand:</b>&nbsp;<font color=red size=2><b>*</b></font>&nbsp;</TD>
	<TD>
        <SELECT style="width:250" id=cboNew name=cboNew LANGUAGE=javascript onkeypress="return combo_onkeypress()" onfocus="return combo_onfocus()" onclick="return combo_onclick()" onkeydown="return combo_onkeydown()">
			<OPTION value="0"></OPTION>
			<%=strAvailable %>
		</SELECT>
</TR>
</Table>

<input id="txtProductID" name="txtProductID" type="text" style="display:none" value="<%=clng(request("ProductID"))%>">
<input id="txtBrandID" name="txtBrandID" type="text" style="display:none" value="<%=clng(request("BrandID"))%>">
</form>
<%
    
        set rs = nothing
        cn.close
        set cn = nothing
    
    
    end if
    
    
%>
</BODY>
</HTML>
