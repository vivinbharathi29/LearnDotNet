<%@ Language="VBScript" %>
<%
    Response.Buffer = True
    Response.ExpiresAbsolute = Now() - 1
    Response.Expires = 0
    Response.CacheControl = "no-cache"
    
    Dim sProductName        'STRING
    Dim iProductID          'INTEGER
    Dim AppRoot             'STRING
    Dim cn                  'DB CONNECTION
    Dim rs                  'RECORDSET
    Dim followMktName: followMktName =0
    Dim Logo1 : Logo1=""
    Dim strLogoBadge1 : strLogoBadge1=""
    Dim isCMPermission : isCMPermission=0
    Dim strMasterLabel1: strMasterLabel1=""
    Dim strCTOModelNumber1:strCTOModelNumber1=""
    AppRoot = Session("ApplicationRoot")

    If Not IsNumeric(Request("ID")) Then
        iProductID = 0
    Else
        iProductID = CLng(Request("ID"))
    End if
    If Not (IsNull(Request("FollowMktName"))) then
        followMktName = Request("FollowMktName")
    end if

    If iProductID = 0 Then
        sProductName = ""
    Else
        set cn = server.CreateObject("ADODB.Connection")
	    cn.ConnectionString = Session("PDPIMS_ConnectionString")
	    cn.Open

	    set rs = server.CreateObject("ADODB.recordset")
	    rs.ActiveConnection = cn
        rs.open  "usp_Product_ValidateProductRTPDate " & iProductID ,cn,adOpenForwardOnly
          
        If rs.EOF Then
            sProductName = ""
        Else
            Do While Not rs.EOF
                sProductName = rs.Fields("productname").value + ","  + sProductName   
                rs.MoveNext()
            Loop
        End If
        rs.Close
    End If
    if(Request("isCM") ="True") then
        isCMPermission = 1
    end if
%>

<html>
<head>
<base target="_self" />
<title></title>
<meta name="GENERATOR" content="Microsoft Visual Studio 6.0" />
    <link href="<%= AppRoot %>/style/wizard style.css" type="text/css" rel="stylesheet" />
    <link href="<%= AppRoot %>/style/Excalibur.css" type="text/css" rel="stylesheet" />
    <link rel="stylesheet" type="text/css" href="<%= AppRoot %>/SupplyChain/style.css" />
    <!-- #include file="../../includes/bundleConfig.inc" -->
</head>
<body>


    <div style="margin-top:4px; height: 300px; width:100%; background-color: white" id="divPlatform">
        <% if request("Edit") <> "0" then %>
        <table id="tabPlatforms1" style="width:100%; border-collapse: collapse; border:solid 1px tan; background-color:cornsilk" border="1">
            <tr>
                <td class="smallest"><a href="javascript: AddPlatform(<%= request("ID")%>, <%=followMktName%>);">Add New Base Unit Group</a></td>
                <% if (followMktName = 0) then %>
                    <td class="smallest"><a href="javascript: AddExistingPlatform(<%= request("ID")%>);">Add Existing Base Unit Groups</a></td>          
                <% end if %>
                <td class="smallest"><a  href="javascript: RefreshPage();">Refresh</a></td>
            </tr>
        </table>        
        <% end if %>
        <table style = "width:100%" id="TablePlatform">
            <colgroup>
                <col style="width:10px" />
                <col style="width:50px" />
                <col style="width:20px" />
                <col style="width:150px" />
                <col style="width:150px" />
                <col style="width:250px" />
                <col style="width:150px" />
                <col style="width:150px" />
                <col style="width:150px" />
                <col style="width:200px" />
                <col style="width:150px" />
                <col style="width:30px" />
                <col style="width:30px" />
                <col style="width:30px" />
                <col style="width:30px" />
                <col style="width:80px" />
                <col style="width:80px" />
            </colgroup>
            <!--<thead>-->
                <tr style="position: relative; top: expression(document.getElementById('divPlatform').scrollTop-3);">
                    <td style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;
                        border-bottom: 1px outset; background-color:#c9ddff" class="smallest">
                        &nbsp;
                    </td>
                    <td style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;
                        border-bottom: 1px outset; background-color:#c9ddff" class="smallest">
                        Active&nbsp;Base&nbsp;Units&nbsp;(Total)
                    </td>
                    <td style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;
                        border-bottom: 1px outset;background-color:#c9ddff" class="smallest">
                        Year
                    </td>
                     <td style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;
                        border-bottom: 1px outset;   background-color:#c9ddff" class="smallest">
                        System Board
                    </td>
                    <% if (followMktName = 1) then %>
                        <td style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;
                            border-bottom: 1px outset; background-color:#c9ddff" class="smallest">
                            &nbsp;Marketing&nbsp;Name
                        </td>
                        <td  style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;
                            border-bottom: 1px outset;background-color:#c9ddff" class="smallest">
                            &nbsp;PHWeb&nbsp;Family Name
                         </td>
                    <% else %>
                        <td  style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;
                            border-bottom: 1px outset;background-color:#c9ddff" class="smallest">
                            &nbsp;Generic&nbsp;Name
                        </td>
                        <td style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;
                            border-bottom: 1px outset; background-color:#c9ddff" class="smallest">
                            &nbsp;Marketing&nbsp;Name
                        </td>
                    <% end if %>
                    
                    <% if (followMktName = 1) then %>
                        <td  style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;
                            border-bottom: 1px outset;background-color:#c9ddff" class="smallest">
                            Logo Badge C Cover
                        </td>
                        <td  style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;
                            border-bottom: 1px outset;background-color:#c9ddff" class="smallest">
                            Model Number (Service Tag down)
                        </td>
                        <td  style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;
                            border-bottom: 1px outset;background-color:#c9ddff;" class="smallest">
                            CTO Model Number
                        </td>
                    <% end if %>
                    <% if (followMktName <> 1) then %>
                        <td style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;
                            border-bottom: 1px outset; background-color:#c9ddff" class="smallest">
                            &nbsp;Deployment
                        </td>
                    <% end if %>
                    <td style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;
                        border-bottom: 1px outset; background-color:#c9ddff" class="smallest">
                        &nbsp;Chassis
                    </td>
                    <td style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;
                        border-bottom: 1px outset; background-color:#c9ddff" class="smallest">
                        &nbsp;Brand&nbsp;Name
                    </td>  
                    <td style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;
                        border-bottom: 1px outset; background-color:#c9ddff" class="smallest">
                        &nbsp;eMMC&nbsp;onboard
                    </td>
                    <td style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;
                        border-bottom: 1px outset; background-color:#c9ddff" class="smallest">
                        &nbsp;Memory&nbsp;onboard
                    </td>
                    <td style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;
                        border-bottom: 1px outset; background-color:#c9ddff" class="smallest">
                        &nbsp;PCA&nbsp;Graphic&nbsp;Type
                    </td>
                    <td style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;
                        border-bottom: 1px outset; background-color:#c9ddff" class="smallest">
                        &nbsp;Model&nbsp;Number
                    </td>
                    <td style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;
                        border-bottom: 1px outset; background-color:#c9ddff" class="smallest">
                        &nbsp;Graphic&nbsp;Capacity
                    </td>
                    <% if (followMktName <> 1) then %>
                        <td style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset;
                            border-bottom: 1px outset; background-color:#c9ddff" class="smallest">
                            &nbsp;Display
                        </td>                  
                    <% end if %>
                    
                </tr>
           <!-- </thead>-->
<%
	    set cn = server.CreateObject("ADODB.Connection")
	    cn.ConnectionString = Session("PDPIMS_ConnectionString")
	    cn.Open

	    set rs = server.CreateObject("ADODB.recordset")
	    rs.ActiveConnection = cn

        if request("ID") <> "" then

		    rs.open  "spListProductPlatforms " & clng(request("ID")) ,cn,adOpenForwardOnly
		    do while not rs.eof
                response.write "<tr>"
                response.write "<td class=""cell""><a href='javascript:void' onclick='return RemovePlatform(" & rs("ID") & "," & request("ID") & "," & rs("PlatformID") & "," & rs("ProductBrandID") & "," & followMktName &")'>Remove</a></td>"
                if request("Edit") <> "0" then
                    response.write "<td align=""middle"" class=""cell""><a href=""#;"" onclick=""return ShowPlatformBaseUnits(" & rs("PlatformID") & "," & request("ID") & ",'" & rs("GenericName") & "');"">" & (rs("ActiveBaseUnits")) & " (" & (rs("BaseUnitCount")) & ")" & "</a></td>"
                else
                    response.write "<td align=""middle"" class=""cell"">" & (rs("ActiveBaseUnits")) & " (" & (rs("BaseUnitCount")) & ")" & "</td>"
                end if
                response.write "<td class=""cell"" onmouseover=""return rows_onmouseover()"" onmouseout=""return rows_onmouseout()"" onclick=""return UpdatePlatform(" & rs("PlatformID") & "," & request("ID") & "," & followMktName &");"">" & rs("IntroYear") & "</td>"
                response.write "<td nowrap class=""cell"" onmouseover=""return rows_onmouseover()"" onmouseout=""return rows_onmouseout()"" onclick=""return UpdatePlatform(" & rs("PlatformID") & "," & request("ID") & "," & followMktName & ");"">" & rs("SystemBoardName") & "</td>"
                if (followMktName = 1) then
                    response.write "<td nowrap class=""cell"" onmouseover=""return rows_onmouseover()"" onmouseout=""return rows_onmouseout()"" onclick=""return UpdatePlatform(" & rs("PlatformID") & "," & request("ID") & "," & followMktName & ");"">" & rs("Marketingname")  & "</td>"
                    response.write "<td nowrap class=""cell"" onmouseover=""return rows_onmouseover()"" onmouseout=""return rows_onmouseout()"" onclick=""return UpdatePlatform(" & rs("PlatformID") & "," & request("ID") & "," & followMktName & ");"">" & rs("PHWebFamilyName") & "</td>"
                else
                    response.write "<td nowrap class=""cell"" onmouseover=""return rows_onmouseover()"" onmouseout=""return rows_onmouseout()"" onclick=""return UpdatePlatform(" & rs("PlatformID") & "," & request("ID") & "," & followMktName & ");"">" & rs("GenericName") & "</td>"
                    response.write "<td nowrap class=""cell"" onmouseover=""return rows_onmouseover()"" onmouseout=""return rows_onmouseout()"" onclick=""return UpdatePlatform(" & rs("PlatformID") & "," & request("ID") & "," & followMktName & ");"">" & rs("Marketingname")  & "</td>"
                end if
                
                
                if (followMktName = 1) then
                    strLogoBadge1=""
                    strMasterLabel1=""
                    strCTOModelNumber1=""
                    set rs2 = server.CreateObject("ADODB.recordset")
	                rs2.ActiveConnection = cn
                    rs2.open  "usp_GetBrands4Platform "& clng(request("ID")) &", " & rs("PlatformID") ,cn,adOpenForwardOnly
                    if not(rs2.EOF and rs2.bof) then
                        '====logo name start
                        if isCMPermission then
					    	Logo1 = rs2("StreetName3") & " "
                        end if
                        if rs2("ShowSeriesNumberInLogoBadge") then
                            if rs2("SplitSeriesForLogoAndBrand") then
                                if isCMPermission then
                                    Logo1 = Logo1 & val(rs2("SeriesName"))
                                end if
                            else
                                if isCMPermission then
					    			Logo1 = Logo1 & rs2("SeriesName")							
                                end if
                            end if
					    end if
                        if (rs2("LogoBadge") = "") then
                           	if isCMPermission then
					    		Logo1 = rs2("StreetName3") & " "
					    	else
					    		strLogoBadge1 =  strLogobadge1 &rs2("StreetName3") & " "
					    	end if						
                            if rs2("ShowSeriesNumberInLogoBadge") then
                                if rs2("SplitSeriesForLogoAndBrand") then
                                    if isCMPermission then
                                        Logo1 = Logo1 & val(rs2("SeriesName"))
                                    else
					    			    strLogoBadge1 = strLogobadge1 & val(rs2("SeriesName"))
					    			end if
                                else
                                    if isCMPermission then
					    			    Logo1 = Logo1 & rs2("SeriesName")
					    			else
					    			    strLogoBadge1 = strLogobadge1 & rs2("SeriesName")
					    			end if
                                end if

					    	end if                     
                        end if
                        if trim(rs2("LogoBadge") & "") <> "" and Logo1 <> "" then						
                            strLogoBadge1 = strLogoBadge1 & trim(rs2("LogoBadge"))
					    elseif Logo1 <> "" then 
                            strLogoBadge1 = strLogoBadge1 & trim(Logo1) 		
					    else 
					    	strLogoBadge1 = strLogoBadge1 & rs2("LogoBadge") 
					    end if
                        '====logo name end
                        '====Master Label start
                        if rs2("MasterLabel") <> "" and isCMPermission then
					    	    strMasterLabel1 = strMasterLabel1 & "<a href=""javascript:ShowMarketingNameDialog(" & rs2("ID") & ",'" & rs2("MasterLabel") & "',8," & rs2("ProductBrandID") & ",'" & rs2("SeriesID") & "'," & rs("PlatformID") & ")"">" & rs2("MasterLabel") & "</a>"
					    elseif isCMPermission then
					    	    strMasterLabel1 = strMasterLabel1 & "<a href=""javascript:ShowMarketingNameDialog(" & rs2("ID") & ",'',8," & rs2("ProductBrandID") & ",'" & rs2("SeriesID") & "'," & rs("PlatformID") & ")"">Add</a>"				
					    else
					    	strMasterLabel1 = strMasterLabel1 & rs2("MasterLabel")
                        end if
                        '====Master Label end
                        '====CTO Model number start
                        if trim(rs2("CTOModelNumber") & "") <> "" and isCMPermission then
					    	strCTOModelNumber1 = strCTOModelNumber1 & "<a href=""javascript:ShowMarketingNameDialog(" & rs2("ID") & ",'" & rs2("CTOModelNumber") & "',5," & rs2("ProductBrandID") & ",'" & rs2("SeriesID") & "'," & rs("PlatformID") & ")"">" & rs2("CTOModelNumber") & "</a>"
					    elseif isCMPermission then
					    	strCTOModelNumber1 = strCTOModelNumber1 & "<a href=""javascript:ShowMarketingNameDialog(" & rs2("ID") & ",'',5," & rs2("ProductBrandID") & ",'" & rs2("SeriesID") & "'," & rs("PlatformID") & ")"">Add</a>"
					    else
					    	strCTOModelNumber1 = strCTOModelNumber1 & rs2("CTOModelNumber")
					    end if
                    '====CTO Model number end
                    End if
                    rs2.Close
                    response.write "<td nowrap class=""cell"" onmouseover=""return rows_onmouseover()"" onmouseout=""return rows_onmouseout()"" onclick=""return UpdatePlatform(" & rs("PlatformID") & "," & request("ID") & "," & followMktName & ");"">" & strLogoBadge1  & "</td>"
                    response.write "<td nowrap class=""cell"" onmouseover=""return rows_onmouseover()"" onmouseout=""return rows_onmouseout()"" >" & strMasterLabel1  & "</td>"
                    response.write "<td nowrap class=""cell"" onmouseover=""return rows_onmouseover()"" onmouseout=""return rows_onmouseout()"" >" & strCTOModelNumber1  & "</td>"
                end if

                if (followMktName <> 1) then
                    response.write "<td nowrap class=""cell"" onmouseover=""return rows_onmouseover()"" onmouseout=""return rows_onmouseout()"" onclick=""return UpdatePlatform(" & rs("PlatformID") & "," & request("ID") & "," & followMktName & ");"">" & rs("Deployment") & "</td>"            
                end if
                response.write "<td nowrap class=""cell"" onmouseover=""return rows_onmouseover()"" onmouseout=""return rows_onmouseout()"" onclick=""return UpdatePlatform(" & rs("PlatformID") & "," & request("ID") & "," & followMktName & ");"">" & rs("Chassis") & "</td>"
                response.write "<td nowrap class=""cell"" onmouseover=""return rows_onmouseover()"" onmouseout=""return rows_onmouseout()"" onclick=""return UpdatePlatform(" & rs("PlatformID") & "," & request("ID") & "," & followMktName & ");"">" & rs("BrandName") & "</td>"
                response.write "<td class=""cell"" onmouseover=""return rows_onmouseover()"" onmouseout=""return rows_onmouseout()"" onclick=""return UpdatePlatform(" & rs("PlatformID") & "," & request("ID") & "," & followMktName & ");"">" & rs("eMMConboard") & "</td>"
                response.write "<td class=""cell"" onmouseover=""return rows_onmouseover()"" onmouseout=""return rows_onmouseout()"" onclick=""return UpdatePlatform(" & rs("PlatformID") & "," & request("ID") & "," & followMktName & ");"">" & rs("MemoryOnboard") & "</td>"
                response.write "<td class=""cell"" onmouseover=""return rows_onmouseover()"" onmouseout=""return rows_onmouseout()"" onclick=""return UpdatePlatform(" & rs("PlatformID") & "," & request("ID") & "," & followMktName & ");"">" & rs("PCAGraphicsType") & "</td>"
                response.write "<td class=""cell"" onmouseover=""return rows_onmouseover()"" onmouseout=""return rows_onmouseout()"" onclick=""return UpdatePlatform(" & rs("PlatformID") & "," & request("ID") & "," & followMktName & ");"">" & rs("ModelNumber")  & "</td>"
                response.write "<td class=""cell"" onmouseover=""return rows_onmouseover()"" onmouseout=""return rows_onmouseout()"" onclick=""return UpdatePlatform(" & rs("PlatformID") & "," & request("ID") & "," & followMktName & ");"">" & rs("GraphicCapacity") & "</td>"
                if (followMktName <> 1) then
                    response.write "<td nowrap class=""cell"" onmouseover=""return rows_onmouseover()"" onmouseout=""return rows_onmouseout()"" onclick=""return UpdatePlatform(" & rs("PlatformID") & "," & request("ID") & "," & followMktName & ");"">" & rs("Display") & "</td>"
                end if
                response.write "</tr>"
    		    rs.MoveNext
    	    loop
	        rs.Close
            
       end if

       function Val(strText)
        dim strOutput
        dim i

        strOutput = ""
        for i = 1 to len(trim(strText))
            if isnumeric(mid(strText,i,1)) then
                strOutput = strOutput & mid(trim(strText),i,1)
            else
                exit for
            end if
        next
        Val = strOutput
      end function
%>
        </table>
    </div>

    <div id="iframeDialog" style="display:none">
            <iframe frameborder="0" name="modalDialog" id="modalDialog"></iframe>
        </div>

    <div id="dialog-confirm" title="Remove Base Unit Group?" style="display:none">

        <div id="pQ"><p><span class="ui-icon ui-icon-alert" style="float:left; margin:0 7px 20px 0;"></span>Are you sure that you want to remove the Base Unit Group from this Product?</p></div>
        <div id="pP" style="display:none"><p><span class="ui-icon ui-icon-scissors" style="float:left; margin:0 7px 20px 0;"></span>Removing Base Unit Group from the Product...</p></div>

    </div>
    <div style="display: none;">
         <div id="divOpenMarketingNameUpdate" title="Coolbeans">
            <iframe frameborder="0" name="ifOpenMarketingNameUpdate" id="ifOpenMarketingNameUpdate"></iframe>
        </div>
    </div>

    <%
    set rs = nothing
    cn.close
    set cn= nothing

%>	
<script id="clientEventHandlersJS"  language="javascript" type="text/javascript">
<!--

    var oPopup = window.createPopup();

    function RemovePlatform(ID, PVID, PlatformID, PBID, followMKT) {
       
        $("#dialog-confirm").dialog({
            resizable: false,
            width: 400,
            minWidth: 400,
            height: 180,
            minHeight: 140,
            modal: true,            
            closeOnEscape: true,
            //open: function(event, ui) { $(".ui-dialog-titlebar-close", ui.dialog | ui).show(); },
            buttons: {                                
                "Cancel": function () {
                    $(this).dialog("close");
                },
                "Delete": function () {                                     

                    $(":button:contains('Cancel')").prop("disabled", true).addClass("ui-state-disabled");
                    $(":button:contains('Delete')").prop("disabled", true).addClass("ui-state-disabled");
                    $(".ui-dialog-titlebar-close").hide();
                    $("#modalDialog").attr("src", "<%=AppRoot %>/SupplyChain/platformDelete.asp?ID=" + ID + "&PVID=" + PVID + "&PlatformID=" + PlatformID + "&PBID=" + PBID + "&followMKT=" + followMKT);
                    $("#pP").show();
                    $("#pQ").hide();

                }
            }
        });

        return false;
        //if (confirm("Are you sure that you want to remove the Platform from this Product?")) {

            //var Completed = 0;
            //Completed = window.parent.showModalDialog("<%=AppRoot %>/SupplyChain/platformDelete.asp?ID=" + ID + "&PVID=" + PVID, "Remove Platform from Product", "dialogWidth:400px;dialogHeight:50px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");

            //if (Completed == 1) {
            //    RefreshPage();
            //}
        //}
    }

    function Completed() {
        $('#dialog-confirm').dialog('close');
        $("#pQ").show();
        RefreshPage();
        return false;
    }

    function rows_onmouseout() {
        if (typeof (oPopup) == "undefined")
            return;

        if (!oPopup.isOpen) {
            if (window.event.srcElement.className == "text")
                window.event.srcElement.parentElement.parentElement.style.color = "black";
            else if (window.event.srcElement.className == "cell")
                window.event.srcElement.parentElement.style.color = "black";
        }
    }
    function rows_onmouseover() {
        if (typeof (oPopup) == "undefined")
            return;
        if (!oPopup.isOpen) {
            if (window.event.srcElement.className == "cell") {
                window.event.srcElement.parentElement.style.color = "red";
                window.event.srcElement.parentElement.style.cursor = "hand";
            }
            else if (window.event.srcElement.className == "text") {
                window.event.srcElement.parentElement.parentElement.style.color = "red";
                window.event.srcElement.parentElement.parentElement.style.cursor = "hand";
            }

            if (typeof (SelectedRow) != "undefined")
                if (SelectedRow != null)
                    SelectedRow.style.color = "black";

        }
    }

    function AddPlatform(ProductVersionID, followMKTName) {
        try {

            if ($("#hidproductname").val() != "") {
                // remove FCS from all areas - task 20243
                alert("The Release to Production (RTP) and End of Manufacturing (EM) is required for every Product Release schedule before the Base Unit Group can be created.  Please go to the Product Schedule Tab and enter the required dates.");
            }
            else {
                window.parent.OpenDialog1("platform.asp?ProductVersionID=" + ProductVersionID + "&FollowMKTName=" + followMKTName, "Base Unit Group - Add", 0, 0, true, true, true);
           }
        }
        catch(e)
        {
            var oParentIframe = parent.document.getElementById('pmview_PlatformFrame');
            if (typeof (oParentIframe) != 'undefined' && oParentIframe != null) {
                if (window.event.srcElement.className == "cell") {
                    parent.AddPlatform(ProductVersionID, followMKTName);
                }
            } else {
                var sheight = $(window).height() * (90 / 100);
                if (sheight < 600 && screen.height < 600)
                    sheight = 600;
                else
                    sheight = screen.height * (70 / 100);

                var sWidth = $(window).width() * (50 / 100);
                if (sWidth < 800)
                    sWidth = 800;

                var returnValue = window.showModalDialog("platform.asp?ProductVersionID=" + ProductVersionID + "&FollowMKTName=" + followMKTName, "", "dialogwidth:" + sWidth + "px; dialogheight:" + sheight + "px");
                if (returnValue == 1)
                    RefreshPage();
            }
        }
    }    

    function UpdatePlatform(ID, ProductVersionID, followMKTName) {
        try {
            window.parent.OpenDialog1("platform.asp?ID=" + ID + "&ProductVersionID=" + ProductVersionID + "&FollowMKTName=" + followMKTName, "Base Unit Group - Update", 0, 0, true, true, true);
        }
        catch (e) {
            var oParentIframe = parent.document.getElementById('pmview_PlatformFrame');
            if (typeof (oParentIframe) != 'undefined' && oParentIframe != null) {
                if (window.event.srcElement.className == "cell") {
                    parent.UpdatePlatform(ID, ProductVersionID, followMKTName);
                }
            } else {
                if (window.event.srcElement.className == "cell") {
                    var sheight = $(window).height() * (90 / 100);
                    if (sheight < 600 && screen.height < 600)
                        sheight = 600;
                    else
                        sheight = screen.height * (70 / 100);

                    var sWidth = $(window).width() * (50 / 100);
                    if (sWidth < 900)
                        sWidth = 900;

                    var returnValue = window.showModalDialog("platform.asp?ID=" + ID + "&ProductVersionID=" + ProductVersionID + "&FollowMKTName=" + followMKTName, "", "dialogwidth:" + sWidth + "px; dialogheight:" + sheight + "px");
                    if (returnValue == 1)
                        RefreshPage();
                }
            }
        }        
    }

    function ShowPlatformBaseUnits(ID, ProductVersionID, PlatformName) {
        window.parent.ViewPlatformBaseUnits(ID, ProductVersionID, PlatformName);
        /*var returnValue = window.showModalDialog("platformBaseUnitList.asp?ID=" + ID + "&ProductVersionID=" + ProductVersionID + "&PlatformName=" + PlatformName);
        if (returnValue == 1)
            ShowPlatformBaseUnits(ID, ProductVersionID, PlatformName);
        else
            RefreshPage();*/
    }   

    function RefreshPage() {
        window.location.reload();
    }

    function AddExistingPlatform(ID) {



        if ($("#hidproductname").val() != "") {
            // remove FCS from all areas - task 20243
            alert("The Release to Production (RTP) and End of Manufacturing (EM) is required for every Product Release schedule before the Base Unit Group can be created.  Please go to the Product Schedule Tab and enter the required dates.");
        } else {
            window.parent.AddExistingBaseUnitGroup(ID);
        }
    }

    function ShowMarketingNameDialog(BrandID, ExistingName, NameType, ProductBrandID, Series, PlatFormID) {
        var DlgWidth = 650;
        var DlgHeight = 330;
        var DialogName = ""
        if ($(window).height() < DlgHeight) DlgHeight = $(window).height() + 100;
        if ($(window).width() < DlgWidth) DlgWidth = $(window).width();
        $("#divOpenMarketingNameUpdate").dialog({ width: DlgWidth, height: DlgHeight });
        $("#ifOpenMarketingNameUpdate").attr("width", "100%");
        $("#ifOpenMarketingNameUpdate").attr("height", "100%");
        $("#ifOpenMarketingNameUpdate").attr("src", '<%=AppRoot %>/UpdateMarketingNames.asp?BID=' + BrandID + '&PBID=' + ProductBrandID + '&Name=' + ExistingName + '&Type=' + NameType + '&Series=' + Series+'&PlatFormID='+PlatFormID);


        if (NameType == 5 && ExistingName == "") {
            DialogName = "Add CTO Model Number";
        }
        if (NameType == 5 && ExistingName != "") {
            DialogName = "Edit CTO Model Number";
        }

        if (NameType == 6 && ExistingName == "") {
            DialogName = "Add Short Name";
        }
        if (NameType == 6 && ExistingName != "") {
            DialogName = "Edit Short Name";
        }

        if (NameType == 7 && ExistingName == "") {
            DialogName = "Add HP Brand Name (Service Tag up)";
        }
        if (NameType == 7 && ExistingName != "") {
            DialogName = "Edit HP Brand Name (Service Tag up)";
        }

        if (NameType == 8 && ExistingName == "") {
            DialogName = "Add Model Number (Service Tag down)";
        }
        if (NameType == 8 && ExistingName != "") {
            DialogName = "Edit Model Number (Service Tag down)";
        }

        if (NameType == 9 && ExistingName == "") {
            DialogName = "Add BIOS Branding";
        }
        if (NameType == 9 && ExistingName != "") {
            DialogName = "Edit BIOS Branding";
        }

        $("#divOpenMarketingNameUpdate").dialog("option", "title", DialogName);
        $("#divOpenMarketingNameUpdate").dialog("open");
    }

    function CloseMarketingNameDialog() {
        $("#divOpenMarketingNameUpdate").dialog("close");
    }
//-->
</script>

        <input type="hidden" id="hidproductname" name="hidproductname" value="<%= sProductName%>" />

</body>
</html>
