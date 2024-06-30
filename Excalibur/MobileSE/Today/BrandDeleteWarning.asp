<%@ Language="VBScript" %>
<%Option Explicit%>
<!-- #include file="../../includes/no-cache.asp" -->
<!-- #include file = "../../includes/noaccess.inc" -->
<!-- #include file="../../includes/DataWrapper.asp" -->
<!-- #include file = "../../includes/Security.asp" --> 
<!-- #include file = "../../includes/lib_debug.inc" -->
<%	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
%>

<html>
<head>
    <title>Confirm Brand Deletion</title>

    <SCRIPT ID=clientEventHandlersJS LANGUAGE=javascript>
    <!--

    function cmdOK_onclick() {
	    window.returnValue = "1";
	    window.close();
    }
    
    function cmdCancel_onclick() {
	    window.returnValue = "0";
	    window.close();
    }    
    //-->
    </SCRIPT>
</head>
<style>
    body{
        FONT-Family: verdana;
        FONT-Size: x-small;
    }
</style>

<body bgcolor="ivory" leftmargin="20px" topmargin="20px">

    <%
        dim strBrand
        dim strProduct
        strbrand  = "the Brand """ & request("BrandName") & """"
        if strBrand = "" then
            strBrand = "this Brand"
        end if

        strProduct  = request("ProductName") 
        if strProduct = "" then
            strProduct = "this Product"
        end if

        Dim rs, dw, cn, cmd

        Set rs = Server.CreateObject("ADODB.RecordSet")
        Set cn = Server.CreateObject("ADODB.Connection")
        Set cmd = Server.CreateObject("ADODB.Command")
        Set dw = New DataWrapper
        Set cn = dw.CreateConnection("PDPIMS_ConnectionString")

        ' 07/26/2016 - ADao - Change IRS_Platform_Alias, IRS_Platform, IRS_Alias synonyms to use actual tables 
        rs.Open "select p.MktNameMaster from product_brand pb " &_
                "join productversion_platform pp on pp.ProductBrandID = pb.ID " &_
                "join Platform p on p.PlatformID = pp.PlatformID " &_
                "where pb.ProductVersionID = " & Trim(Request("PVID")) &_
                " And pb.BrandID =" & Trim(Request("BrandID")), cn, adOpenForwardOnly
        
        if rs.EOF then
                                     
    %>

    <font size=3><b>Delete Brand</b></font><br /><br />
    <font color=red><b>Removing this Brand will delete the associated Localizations, SCM, and Program Matrix.</b></font><br /><br />
    If you want to change the Brand associated with these items, please click the Brand name link on the previous page.
    <br /><br />

    Are you sure you want to remove <%=strBrand%> from <%=strProduct%>?<br /><br />
    <hr />
    <table width=100%>
    <tr><td align=right>
        <input id="cmdOk" style="width:75px" type="button" value="Yes" LANGUAGE=javascript onclick="return cmdOK_onclick()" />&nbsp;<input style="width:75px"  id="cmdCancel" type="button"  value="No" LANGUAGE=javascript onclick="return cmdCancel_onclick()"/></td></tr>
    </table>

    <% 
        Else

        strbrand  = request("BrandName") & " - " & request("ProductBrandID")
        
        Dim Platforms
        do while not rs.EOF
         
            if Platforms <> "" then
                Platforms = Platforms & ", " 
            end if

            Platforms = Platforms & rs("MktNameMaster")    
                                       
            rs.MoveNext
                                    
	    loop
    %>
    
    <font size=3><b>Delete Brand</b></font><br /><br />
    <div style="font-weight:bold; color:red">
        <p>This Brand (<%=strBrand%>) has been selected for the Base Unit Groups (<%= Platforms %>) in the Product (<%=strProduct%>).  You have to reassign Base Unit Groups to another Brand before you can remove this Brand (<%=strBrand%>) from this Product (<%=strProduct%>) in Base Unit Group Tab.</p>
        <p>You can also Update the Brand by selecting the Brand's name in the General tab and selecting a new Brand to replace the old Brand. When you save and close the Product Properties, the new Brand will be associated with the Base Unit Groups the old Brand was associated with."</p>
    </div>
    <div style="float:right">
        <input style="width:75px"  id="cmdClose" type="button"  value="Close" LANGUAGE=javascript onclick="return cmdCancel_onclick()"/>
    </div>   
    <%
        end if

        rs.Close
        cn.Close  %>
</body>
</html>

