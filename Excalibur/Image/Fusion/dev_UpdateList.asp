<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<html>
<head>

<meta name="GENERATOR" content="Microsoft Visual Studio 6.0" />
<title>IRS Component Updates</title>
<script type="text/javascript" id="clientEventHandlersJS" language="javascript">
<!--
    function ColumnSort(ID) {
        if (ID == 0)
            window.location.href = "IRSComponentUpdateList.asp?SortField=0";
        else if (ID == 1)
            window.location.href = "IRSComponentUpdateList.asp?SortField=1";
        else if (ID == 3)
            window.location.href = "IRSComponentUpdateList.asp?SortField=3";
        else
            window.location.href = "IRSComponentUpdateList.asp";
    }

    function column_onmouseover() {
        event.srcElement.style.color = "red";
        event.srcElement.style.cursor = "pointer";
    }

    function column_onmouseout() {
        event.srcElement.style.color = "black";
        event.srcElement.style.cursor = "default";
    }
//-->
</script>
<style>
    td
    {
         font-family: Verdana;
         font-size: xx-small;
    }
    body
    {
         font-family: Verdana;
         font-size: x-small;
    }

    a:visited
    {
        color: blue
    }
    a:hover
    {
        color: red
    }
    a
    {
        color: blue
    }
</style>
</head>

<body bgcolor="white">




<%
	dim cn
	dim rs
    dim ProductArray
    dim strProductList
    dim strRow
    dim RowCount
    dim LastID
    dim strAction
    dim OSArray
    dim OSSelectedArray
    dim ProductSelectedArray
    dim blnOK2Show
    dim RowsDisplayed
    dim ColorArray
    dim ColorID

    ColorArray = split("white,#E0E0E0",",")

    RowsDisplayed = 0

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Application("PDPIMS_ConnectionString") ' "Provider=SQLOLEDB.1;Persist Security Info=True;Initial Catalog=prs;Server=c.onspdp;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;User Id=pdpadmin;PASSWORD=dino;"
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
	set rs2 = server.CreateObject("ADODB.recordset")
    if trim(request("SortField")) = "" then
        rs.Open "spListComponentsToPutInIRSProductDrop ",cn,adOpenForwardOnly
    else
        rs.Open "spListComponentsToPutInIRSProductDrop " & clng(request("SortField")),cn,adOpenForwardOnly
    end if
    if rs.eof and rs.bof then
        response.write "No updates found."
        rs.close
    else

        'Get list of products to display
        'if request("lstProducts") = "" then
            strProductList = ""
            rs2.open "spListProductDropProductList",cn
            do while not rs2.eof
                strProductList = strProductList & "," & rs2("Product")
                rs2.movenext
            loop
            rs2.close
            if strProductList <> "" then
                ProductArray = split(mid(strProductList,2),",")
            else
                ProductArray = split("Product",",")
            end if
        'else
        '    ProductArray = split(replace(request("lstproducts"),", ",","),",")
        'end if

        
        response.write "<table width=""100%""><tr><td><font face=""verdana"" size=2><b>IRS Component Update List</b><br /></font></td><td align=right><a target=_blank href=""../IRSComponentUpdateList.asp"">View Old Report Format</a></td></tr></table>"
        response.write "<form style=""margin:0px 0px 0px 0px"" id=""frmMain"" method=""post"" action=""IRSComponentUpdateList.asp"">"
        response.write "<div style=""border:2px solid #B0B0B0 ;background-color:#E0E0E0;padding: 2px 2px 2px 2px;margin-top:6px;margin-bottom:6px""><font face=""verdana"" size=2><b><u>Report Filters</u></b><br> "
        response.write "<table><tr><td>OS:&nbsp;</td><td>"
        if instr(request("lstOS"),"Win7") > 0 or trim(request("lstOS")) = "" then
            response.write "<input id=chkWin7 name=""lstOS"" type=""checkbox"" value=""Win7"" checked/>&nbsp;Win7&nbsp;&nbsp;"
        else
            response.write "<input id=chkWin7 name=""lstOS"" type=""checkbox"" value=""Win7"" />&nbsp;Win7&nbsp;&nbsp;"
        end if
        
        if instr(request("lstOS"),"Win8") > 0 or trim(request("lstOS")) = "" then
            response.write "<input id=chkWin8 name=""lstOS"" type=""checkbox"" value=""Win8"" checked />&nbsp;Win8"
        else
            response.write "<input id=chkWin8 name=""lstOS"" type=""checkbox"" value=""Win8"" />&nbsp;Win8"
        end if
        response.write "</td></tr>"


        response.write "<tr><td>Products:&nbsp;</td><td>"
        for i = 0 to ubound(Productarray)
            ProductSelectedArray = split(request("lstProducts"),",")
            if inarray(ProductSelectedArray,ProductArray(i)) or trim(request("lstProducts")) = "" then
                strChecked = " checked "
            else
                strChecked = ""
            end if
            response.write "<input id=""chk" &  ProductArray(i) & """ name=""lstProducts"" type=""checkbox"" value=""" &  ProductArray(i) & """ " & strChecked & " />&nbsp;" &  ProductArray(i) & "&nbsp;&nbsp;"
        next
        response.write "</td></tr>"


        response.write "</table>"
        response.write "<input id=""cmdApply"" type=""submit"" value=""Apply"" />&nbsp;"
        response.write "<input id=""cmdReset"" type=""button"" value=""Reset"" onclick=""javascript:window.navigate('IRSComponentUpdateList.asp')"" />"




        response.write "</font></div>"

      '  dim strProductOSFamilyColumns
	   ' dim ProductOSFamilyArray
            
        dim strDeliverableOSFamilyList
        strDeliverableOSFamilyList = ""
    
'        set rs = server.CreateObject("ADODB.recordset")
 '       rs2.open "spListProductOSFamiliesPreinstalled " & rs("ProductID") & ",2",cn,adOpenStatic
	'    strProductOSFamilyColumns = ""
	'    do while not rs2.eof
	'        if rs2("Name") <> "FD" then
	'            strProductOSFamilyColumns = strProductOSFamilyColumns & "," & rs2("name")
	'        end if
	'        rs2.movenext
	'    loop
	'    rs2.close
	'    if strProductOSFamilyColumns <> "" then
	 '       strProductOSFamilyColumns = mid(strProductOSFamilyColumns,2)
	'    end if
        strProductOSFamilyColumns = "Win7,Win8"
	    ProductOSFamilyArray = split(strProductOSFamilyColumns,",")

     '   response.write strProductOSFamilyColumns
     response.write "</form>"
%>
	    <table id="tabParts" width="100%" bgcolor="white" border="1" cellspacing="0" cellpadding="2" style="border-color: Gray">
		<tr>
	    	<td onmouseover="javascript:column_onmouseover();" onmouseout="javascript:column_onmouseout();" onclick="javascript:ColumnSort(0)" style=" background-color:lightsteelblue; white-space:nowrap"><b>IRS Part Number</b></td>
            <%
	           ' for i = 0 to ubound(ProductOSFamilyArray)
	           '     response.write 	"<td width=""10"" style=""background-color:lightsteelblue;white-space:nowrap""><b>" & ProductOSFamilyArray(i) & "&nbsp;&nbsp;</b></td>"
	           ' next
	        %>

	    	<td onmouseover="javascript:column_onmouseover();" onmouseout="javascript:column_onmouseout();" onclick="javascript:ColumnSort(1)" style="background-color:lightsteelblue;white-space:nowrap"><b>Excalibur ID</b></td>
	    	<td onmouseover="javascript:column_onmouseover();" onmouseout="javascript:column_onmouseout();" onclick="javascript:ColumnSort(2)" style="background-color:lightsteelblue;white-space:nowrap"><b>Name</b></td>
	    	<td style="background-color:lightsteelblue;white-space:nowrap"><b>Version</b></td>
	    	<td style="background-color:lightsteelblue;white-space:nowrap"><b>Rev</b></td>
   	    	<td style="background-color:lightsteelblue;white-space:nowrap"><b>Pass</b></td>
   	    	<td onmouseover="javascript:column_onmouseover();" onmouseout="javascript:column_onmouseout();" onclick="javascript:ColumnSort(3)" style="background-color:lightsteelblue;white-space:nowrap"><b>Issues</b></td>

   	    	<td style="background-color:lightsteelblue;white-space:nowrap"><b>Product</b></td>
   	    	<td style="background-color:lightsteelblue;white-space:nowrap"><b>Action</b></td>
   	    	<td style="background-color:lightsteelblue;white-space:nowrap"><b>Replace</b></td>
   	    	<td style="background-color:lightsteelblue;white-space:nowrap"><b>Win7</b></td>
   	    	<td style="background-color:lightsteelblue;white-space:nowrap"><b>Win8</b></td>
   	    	<td style="background-color:lightsteelblue;white-space:nowrap"><b>Target Notes</b></td>
   	    	<td style="background-color:lightsteelblue;white-space:nowrap"><b>Image Summary</b></td>
		</tr>
        <% 

        strRow = ""
        LastID = 0 
        RowCount = 0
        ColorID=1
        do while not rs.eof

            strSQL = trim(rs("images") & "")
            if strSQl = "" then
                strDeliverableOSFamilyList = strProductOSFamilyColumns
            else
                if instr(strSQL, ":")> 0 then
                    strSQl = left(strSQl,instr(strSQL, ":")-1)
                end if
                if right(trim(strSQl),1) = "," then
                    strSql = strSQl & "0"
                end if
                strSQL = "Select distinct f.shortname as Name, f.id " & _
                            "from images i with (NOLOCK), ImageDefinitions d with (NOLOCK), oslookup o with (NOLOCK), osfamily f with (NOLOCK) " & _
                            "where d.id = i.imagedefinitionid " & _
	                        "and f.id = osfamilyid " & _
	                        "and d.osid = o.id " & _
                            "and i.id in (" & strSQL & ")"
                set rs2 = server.CreateObject("ADODB.recordset")
	            rs2.open strSQL,cn,adOpenStatic
	            if (rs2.eof and rs2.bof) then
    	            strDeliverableOSFamilyList = ""
                else
                    ' strDeliverableOSFamilyList = rs2("OSFamilyList") & ""
                    strDeliverableOSFamilyList = ""
                    do while not rs2.eof
                        strDeliverableOSFamilyList = strDeliverableOSFamilyList & "," & rs2("Name")
                        rs2.movenext
                    loop
                end if
	            rs2.close
	            set rs2 = nothing
                if strDeliverableOSFamilyList <> "" then
                    strDeliverableOSFamilyList = mid(strDeliverableOSFamilyList,2)
                end if
	        end if
	        strDeliverableOSFamilyList = "," & replace(strDeliverableOSFamilyList," ","") & "," 
            OSArray = split(strDeliverableOSFamilyList,",")


            blnOK2Show=true

            if trim(request("lstProducts")) <> "" then
                ProductSelectedArray = split(request("lstProducts"),",")
                if inarray(ProductSelectedArray,rs("Product")) = 0 then
                    blnOK2Show = false
                end if
            end if

            if trim(request("lstOS")) <> "" and blnOK2Show then
                OSSelectedArray = split(request("lstOS"),",")
                blnFound = false
                for i = 0 to ubound(OSArray)
                    if inarray(OSSelectedArray,OSArray(i)) then
                        blnFound = true
                        exit for
                    end if
                next
                if not blnFound then
                    blnOK2Show = false
                end if

            end if

            if blnOK2Show then
                RowsDisplayed = RowsDisplayed + 1
                if rs("PreviousVersionCount") = 0 then
                    strAction = "Add"
                elseif rs("PreviousVersionCount") = -1 then
                    strAction = "Remove"
                else
                    strAction = "Replace"
                end if
               if LastID <> rs("VersionID") and LastID <> 0 then
                   response.write replace(strRow,"!#ROWCOUNT#!",RowCount)
                   strRow = ""
                   RowCount = 0
               end if
               LastID = rs("VersionID") 
                RowCount=RowCount+1

                strRow = strRow & "<tr>"
                   if RowCount=1 then
                        if ColorID = 0 then
                            ColorID=1
                        else
                            ColorID=0
                        end if

                        strRow = strRow & "<td style=""background-color:" & ColorArray(ColorID) & ";border-top: 2 solid black"" valign=top rowspan=!#ROWCOUNT#!>" & rs("PartNUmber") & "&nbsp;</td>"
                        strRow = strRow & "<td style=""background-color:" & ColorArray(ColorID) & ";border-top: 2 solid black"" valign=top rowspan=!#ROWCOUNT#!>" & rs("VersionID") & "</td>"
                        strRow = strRow & "<td style=""background-color:" & ColorArray(ColorID) & ";border-top: 2 solid black"" valign=top rowspan=!#ROWCOUNT#!>" & rs("DeliverableName") & "</td>"
                        strRow = strRow & "<td style=""background-color:" & ColorArray(ColorID) & ";border-top: 2 solid black"" valign=top rowspan=!#ROWCOUNT#!>" & rs("Version") & "</td>"
                        strRow = strRow & "<td style=""background-color:" & ColorArray(ColorID) & ";border-top: 2 solid black"" valign=top rowspan=!#ROWCOUNT#!>" & rs("Revision") & "&nbsp;</td>"
                        strRow = strRow & "<td style=""background-color:" & ColorArray(ColorID) & ";border-top: 2 solid black"" valign=top rowspan=!#ROWCOUNT#!>" & rs("Pass") & "&nbsp;</td>"
                        strRow = strRow & "<td title=""" & rs("Transferdescription") & """ style=""background-color:" & ColorArray(ColorID) & ";border-top: 2 solid black"" valign=top rowspan=!#ROWCOUNT#!><font color=red>" & rs("TransferStatus") & "&nbsp;</font></td>"
                        strRow = strRow & "<td style=""background-color:" & ColorArray(ColorID) & ";border-top: 2 solid black"" valign=top nowrap>" & rs("Product") & "&nbsp;</td>"
                        strRow = strRow & "<td style=""background-color:" & ColorArray(ColorID) & ";border-top: 2 solid black"" valign=top>" & strAction & "&nbsp;</td>"
                        strRow = strRow & "<td nowrap style=""background-color:" & ColorArray(ColorID) & ";border-top: 2 solid black"" valign=top>" & replace(rs("PreviousVersionPartNumbers"),",","<br>") & "&nbsp;</td>"
                        for i = 0 to ubound(ProductOSFamilyArray)
                       	    if instr(lcase(strDeliverableOSFamilyList),"," & lcase(trim(ProductOSFamilyArray(i)))) > 0 then
	                               strRow = strRow & 	"<td style=""background-color:" & ColorArray(ColorID) & ";border-top: 2 solid black"" valign=top align=center>" & "&nbsp;X&nbsp;" & "&nbsp;</td>"
	                           else
	                               strRow = strRow & 	"<td style=""background-color:" & ColorArray(ColorID) & ";border-top: 2 solid black"">&nbsp;</td>"
                               end if
	                    next
                        strRow = strRow & "<td style=""background-color:" & ColorArray(ColorID) & ";border-top: 2 solid black"" valign=top>" & rs("TargetNotes") & "&nbsp;</td>"
                        strRow = strRow & "<td style=""background-color:" & ColorArray(ColorID) & ";border-top: 2 solid black"" valign=top>" & rs("ImageSummary") & "&nbsp;</td>"
                    else
                        strRow = strRow & "<td style=""background-color:" & ColorArray(ColorID) & """ valign=top nowrap>" & rs("Product") & "&nbsp;</td>"
                        strRow = strRow & "<td style=""background-color:" & ColorArray(ColorID) & """ valign=top>" & strAction & "&nbsp;</td>"
                        strRow = strRow & "<td nowrap style=""background-color:" & ColorArray(ColorID) & """ valign=top>" & replace(rs("PreviousVersionPartNumbers"),",","<br>") & "&nbsp;</td>"
                        
                        for i = 0 to ubound(ProductOSFamilyArray)
                       	    if instr(lcase(strDeliverableOSFamilyList),"," & lcase(trim(ProductOSFamilyArray(i)))) > 0 then
	                               strRow = strRow & 	"<td style=""background-color:" & ColorArray(ColorID) & """ valign=top align=center>" & "&nbsp;X&nbsp;" & "&nbsp;</td>"
	                           else
	                               strRow = strRow & 	"<td style=""background-color:" & ColorArray(ColorID) & """>&nbsp;</td>"
                               end if
	                    next
                        strRow = strRow & "<td style=""background-color:" & ColorArray(ColorID) & """ valign=top>" & rs("TargetNotes") & "&nbsp;</td>"
                        strRow = strRow & "<td style=""background-color:" & ColorArray(ColorID) & """ valign=top>" & rs("ImageSummary") & "&nbsp;</td>"
                end if
                strRow = strRow & "</tr>"
            end if


            rs.movenext
        loop
        rs.close

        if strRow <> "" then
            response.write replace(strRow,"!#ROWCOUNT#!",RowCount)
        end if

        if RowsDisplayed = 0 then
            response.write "<tr><td colspan=12>No deliverables match the selected criteria.</td></tr>"
        end if

        %>
    	</table>

<%
    response.write "<br><font size=1 face=verdana>Displayed: " & RowsDisplayed & "</font>"
    end if

	set rs=nothing
	set rs2=nothing
	cn.Close
	set cn=nothing


    function InArray(MyArray,strFind)
        dim strElement
	    dim blnFound
			
	    blnFound = false
	    for each strElement in MyArray
		    if trim(strElement) = trim(strFind) and trim(strElement) <> "" then
			    blnFound = true
			    exit for
		    end if
	    next
	    InArray = blnFound
    end function
%>

</BODY>
</HTML>
