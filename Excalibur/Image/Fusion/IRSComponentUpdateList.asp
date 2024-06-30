<%@ Language=VBScript %>

	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>


<%
    Server.ScriptTimeout = 360

    function InArray(MyArray,strFind)
        dim strElement
	    dim blnFound
	
	    blnFound = false
	    for each strElement in MyArray
	        if lcase(trim(strElement)) = lcase(trim(strFind)) and trim(strElement) <> "" then
	            blnFound = true
	            exit for
	        end if
	    next
	    InArray = blnFound
    end function
    
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

    function selectAllProduct() {
        //tdProductFilters
        var chkLstProducts = document.getElementById("tdProductFilters").getElementsByTagName("input");
        for (i = 0; i < chkLstProducts.length; i++) {
            chkLstProducts[i].checked= true;
        }
    }

    function unSelectAllProduct() {
        var chkLstProducts = document.getElementById("tdProductFilters").getElementsByTagName("input");
        for (i = 0; i < chkLstProducts.length; i++) {
            chkLstProducts[i].checked = false;
        }
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
    dim rs2
    dim rsOS 
    dim strSqlOs
    dim ProductArray
    dim ProductIDArray
    dim strProductList
    dim strProductIDList
    dim strRow
    dim RowCount
    dim LastID
    dim strAction
    dim ProductOSFamilyArray 'OS CheckBox List
    dim strProductOSFamilyColumns 'OS CheckBox List
    dim OSDeliverableArray 'OS List from each row of report
    dim strDeliverableOSFamilyList 'OS List from each row of report
    dim OSSelectedArray 'Selected OS
    dim ProductSelectedArray
    dim strProductSelected
    dim blnOK2Show
    dim RowsDisplayed
    dim ColorArray
    dim ColorID
    dim intSortField


    ColorArray = split("white,#E0E0E0",",")

    if request("lstProducts") <> "" then
        strProductSelected = trim(request("lstProducts"))
        ProductSelectedArray = split(request("lstProducts"),",")
    else
        strProductSelected = ""
        ProductSelectedArray = split("",",")
    end if


    if trim(request("SortField")) = "" then
        intSortField = 2
    else
        intSortField = clng(request("SortField"))
    end if

    RowsDisplayed = 0

	set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Application("PDPIMS_ConnectionString") 
    cn.CommandTimeout = 300
	cn.Open

    ''' The OS column was from the DB table 'OSFamily'.
    strProductOSFamilyColumns =""
    set rsOS = server.CreateObject("ADODB.recordset")
    strSqlOs = "select ID, FamilyName, ShortName, SortOrder from OSFamily where Active=1 and ShortName not in ('XP','VSTA','WinE','LNX','Chrome','Android') order by SortOrder ;"
    rsOS.Open strSqlOs,cn,adOpenForwardOnly
        do while not rsOS.eof
            strProductOSFamilyColumns = strProductOSFamilyColumns & trim(rsOS("ShortName")) & ","
            rsOS.movenext
        loop
    rsOS.close
    if right(strProductOSFamilyColumns,1) = "," then
        strProductOSFamilyColumns = left(strProductOSFamilyColumns, Len(strProductOSFamilyColumns)-1)
    end if
    ProductOSFamilyArray = split(strProductOSFamilyColumns,",")


	set rs2 = server.CreateObject("ADODB.recordset")
    
    
    strProductList = ""
    strProductIDList = ""
    rs2.open "spListProductDropProductList",cn
    do while not rs2.eof
        strProductList = strProductList & "," & rs2("Product")
        strProductIDList = strProductIDList & "," & rs2("ProductID")
        rs2.movenext
    loop
    rs2.close
    if strProductList <> "" and strProductList <> "" then
        ProductArray = split(mid(strProductList,2),",")
        ProductIDArray = split(mid(strProductIDList,2),",")
    else
        ProductArray = split("Product",",")
        ProductIDArray = split("0",",")
    end if
        
    response.write "<table width=""100%""><tr><td><font face=""verdana"" size=2><b>IRS Component Update List</b><br /></font></td><td align=right><a target=_blank href=""../IRSComponentUpdateList.asp"">View Old Report Format</a></td></tr></table>"
    response.write "<form style=""margin:0px 0px 0px 0px"" id=""frmMain"" method=""post"" action=""IRSComponentUpdateList.asp"">"
    response.write "<div style=""border:2px solid #B0B0B0 ;background-color:#E0E0E0;padding: 2px 2px 2px 2px;margin-top:6px;margin-bottom:6px""><font face=""verdana"" size=2><b><u>Report Filters</u></b><br> "
    response.write "<table><tr><td>OS:&nbsp;</td><td id=""tdOsFilters"">" & vbcrlf
    OSSelectedArray = split(request("lstOS"),",")

    for i = 0 to ubound(ProductOSFamilyArray) 
        if inarray(OSSelectedArray , ProductOSFamilyArray(i) ) or request("lstOS") = "" then
            response.write "<input id=chk" & ProductOSFamilyArray(i) & " name=""lstOS"" type=""checkbox"" value=""" & ProductOSFamilyArray(i) & """ checked/>&nbsp;" & ProductOSFamilyArray(i) & "&nbsp;&nbsp;" & vbcrlf
        else
            response.write "<input id=chk" & ProductOSFamilyArray(i) & " name=""lstOS"" type=""checkbox"" value=""" & ProductOSFamilyArray(i) & """ />&nbsp;" & ProductOSFamilyArray(i) & "&nbsp;&nbsp;" & vbcrlf
        end if
    next 
    if request("lstOS") = "" then
        OSSelectedArray = ProductOSFamilyArray
    end if

    response.write "</td></tr>"


    response.write "<tr><td>Products:&nbsp;</td><td id=""tdProductFilters"">"
    for i = 0 to ubound(Productarray)
            
        if inarray(ProductSelectedArray,ProductIDArray(i)) then
            strChecked = " checked "
        else
            strChecked = ""
        end if
        response.write "<input id=""chk" &  ProductIDArray(i) & """ name=""lstProducts"" type=""checkbox"" value=""" &  ProductIDArray(i) & """ " & strChecked & " />&nbsp;" &  ProductArray(i) & "&nbsp;&nbsp;"
    next
    response.write "</td></tr>"

    response.write "<tr><td>&nbsp;</td><td> <a href=""#"" onClick=""selectAllProduct();"">Select All</a>&nbsp;<a href=""#"" onClick=""unSelectAllProduct();"">Unselect All</a>   </td></tr>"


    response.write "</table>"
    response.write "<input id=""cmdApply"" type=""submit"" value=""Apply"" />&nbsp;"
    response.write "<input id=""cmdReset"" type=""button"" value=""Reset"" onclick=""javascript:window.navigate('IRSComponentUpdateList.asp')"" />"




    response.write "</font></div>"

    
    
    response.write "</form>"


    if strProductSelected = "" then
        response.write "No Product selected"
    else



        set rs = server.CreateObject("ADODB.recordset")

	    set cm = server.CreateObject("ADODB.Command")
	    Set cm.ActiveConnection = cn
        cm.CommandTimeout = 300
	    cm.CommandType = 4
	    cm.CommandText = "spListComponentsToPutInIRSProductDrop"
	
	    Set p = cm.CreateParameter("@SortField", 3, &H0001, 3)
	    p.Value = intSortField
	    cm.Parameters.Append p

	    Set p = cm.CreateParameter("@lstProductId", 200, &H0001, 8000)
	    p.Value = strProductSelected
	    cm.Parameters.Append p

	    rs.CursorType = adOpenForwardOnly
	    rs.LockType=AdLockReadOnly
	    Set rs = cm.Execute 

        if rs.eof and rs.bof then
            response.write "No updates found."
            rs.close
        else

%>
	    <table id="tabParts" width="100%" bgcolor="white" border="1" cellspacing="0" cellpadding="2" style="border-color: Gray">
		<tr>
	    	<td style=" background-color:lightsteelblue; white-space:nowrap"><b>IRS Part Number</b></td>

	    	<td style="background-color:lightsteelblue;white-space:nowrap"><b>Excalibur ID</b></td>
	    	<td style="background-color:lightsteelblue;white-space:nowrap"><b>Name</b></td>
	    	<td style="background-color:lightsteelblue;white-space:nowrap"><b>Version</b></td>
	    	<td style="background-color:lightsteelblue;white-space:nowrap"><b>Rev</b></td>
   	    	<td style="background-color:lightsteelblue;white-space:nowrap"><b>Pass</b></td>
   	    	<td style="background-color:lightsteelblue;white-space:nowrap"><b>Issues</b></td>

   	    	<td style="background-color:lightsteelblue;white-space:nowrap"><b>Product</b></td>
   	    	<td style="background-color:lightsteelblue;white-space:nowrap"><b>Action</b></td>
   	    	<td style="background-color:lightsteelblue;white-space:nowrap"><b>Replace</b></td>
            <% for i = 0 to ubound(ProductOSFamilyArray) %>
                <td style="background-color:lightsteelblue;white-space:nowrap"><b><%=ProductOSFamilyArray(i) %></b></td>
            <% next %>
   	    	<td style="background-color:lightsteelblue;white-space:nowrap"><b>Target Notes</b></td>
   	    	<td style="background-color:lightsteelblue;white-space:nowrap"><b>Image Summary</b></td>
       
		</tr>

        <% 

            strRow = ""
            LastID = 0 
            RowCount = 0
            ColorID=1
            
            do while not rs.eof

                if isnull(rs("OSList")) then
                    strDeliverableOSFamilyList = " "
                else
                    strDeliverableOSFamilyList = rs("OSList")
                end if
             
                OSDeliverableArray = split(strDeliverableOSFamilyList,",")

                blnOK2Show=true

                if trim(request("lstOS")) <> "" and blnOK2Show then
                    blnFound = false
                    for i = 0 to ubound(OSDeliverableArray)
                        if inarray(OSSelectedArray,OSDeliverableArray(i)) then
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

                    strRow = strRow  & vbcrlf & "<tr>"
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
                       	        if inarray(OSDeliverableArray,ProductOSFamilyArray(i)) then 
	                                strRow = strRow & 	"<td style=""background-color:" & ColorArray(ColorID) & ";border-top: 2 solid black"" valign=top align=center>" & "&nbsp;X&nbsp;" & "&nbsp;</td>"
	                            else
	                                strRow = strRow & 	"<td style=""background-color:" & ColorArray(ColorID) & ";border-top: 2 solid black"">&nbsp;</td>"
                                end if
	                        next

                            strRow = strRow & "<td style=""background-color:" & ColorArray(ColorID) & ";border-top: 2 solid black"" valign=top>" & trim(rs("TargetNotes")) & "&nbsp;</td>"
                            strRow = strRow & "<td style=""background-color:" & ColorArray(ColorID) & ";border-top: 2 solid black"" valign=top>" & trim(rs("ImageSummary")) & "&nbsp;</td>"
                    
                        else
                            strRow = strRow & "<td style=""background-color:" & ColorArray(ColorID) & """ valign=top nowrap>" & rs("Product") & "&nbsp;</td>"
                            strRow = strRow & "<td style=""background-color:" & ColorArray(ColorID) & """ valign=top>" & strAction & "&nbsp;</td>"
                            strRow = strRow & "<td nowrap style=""background-color:" & ColorArray(ColorID) & """ valign=top>" & replace(rs("PreviousVersionPartNumbers"),",","<br>") & "&nbsp;</td>"
                           
                            for i = 0 to ubound(ProductOSFamilyArray)
                       	        if inarray(OSDeliverableArray,ProductOSFamilyArray(i)) then 
	                                   strRow = strRow & 	"<td style=""background-color:" & ColorArray(ColorID) & """ valign=top align=center>" & "&nbsp;X&nbsp;" & "&nbsp;</td>"
	                               else
	                                   strRow = strRow & 	"<td style=""background-color:" & ColorArray(ColorID) & """>&nbsp;</td>"
                                   end if
	                        next
                            strRow = strRow & "<td style=""background-color:" & ColorArray(ColorID) & """ valign=top>" & rs("TargetNotes") & "&nbsp;</td>"
                            strRow = strRow & "<td style=""background-color:" & ColorArray(ColorID) & """ valign=top>" & rs("ImageSummary") & "&nbsp;</td>"
                        
                    end if
                    strRow = strRow & "</tr>"  & vbcrlf
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



    end if



	set rs=nothing
	set rs2=nothing	
    cn.Close
	set cn=nothing

%>

</BODY>
</HTML>
