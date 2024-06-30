<%@  language="VBScript" %>

<%
		
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"

  Dim AppRoot
  AppRoot = Session("ApplicationRoot")
%>

<html>
<head>
    <title>Base Units</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0" />
    <script src="../../includes/client/jquery-1.11.0.min.js"></script>
    <link href="<%= AppRoot %>/style/redmond/jquery-ui-1.8.7.custom.css" rel="stylesheet" type="text/css" />
    <link href="<%= AppRoot %>/style/shared.css" rel="stylesheet" />
    <script src="<%= AppRoot %>/includes/client/jquery.min.js" type="text/javascript"></script>
    <script src="<%= AppRoot %>/includes/client/jquery-ui.min.js" type="text/javascript"></script>
    <script src="<%= AppRoot %>/Scripts/shared_functions.js"></script>
    <script id="clientEventHandlersJS" language="javascript" type="text/javascript">
<!--        
        if (!window.createPopup) {
            window.createPopup = function () {
                var popup = document.createElement("iframe"), //must be iframe because existing functions are being called like parent.func()
                    isShown = false, popupClicked = false;
                popup.src = "about:blank";
                popup.style.position = "absolute";
                popup.style.border = "0px";
                popup.style.display = "none";
                popup.addEventListener("load", function (e) {
                    popup.document = (popup.contentWindow || popup.contentDocument);//this will allow us to set innerHTML in the old fashion.
                    if (popup.document.document) popup.document = popup.document.document;
                });
                document.body.appendChild(popup);
                var hidepopup = function (event) {
                    if (isShown)
                        setTimeout(function () {
                            if (!popupClicked) {
                                popup.hide();
                            }
                            popupClicked = false;
                        }, 150);//timeout will allow the click event to trigger inside the frame before closing.
                }

                popup.show = function (x, y, w, h, pElement) {
                    if (typeof (x) !== 'undefined') {
                        var elPos = [0, 0];
                        if (pElement) elPos = findPos(pElement);//maybe validate that this is a DOM node instead of just falsy
                        elPos[0] += y, elPos[1] += x;

                        if (isNaN(w)) w = popup.document.scrollWidth;
                        if (isNaN(h)) h = popup.document.scrollHeight;
                        if (elPos[0] + w > document.body.clientWidth) elPos[0] = document.body.clientWidth - w - 5;
                        if (elPos[1] + h > document.body.clientHeight) elPos[1] = document.body.clientHeight - h - 5;

                        popup.style.left = elPos[0] + "px";
                        popup.style.top = elPos[1] + "px";
                        popup.style.width = w + "px";
                        popup.style.height = h + "px";
                    }
                    popup.style.display = "block";
                    isShown = true;
                }

                popup.hide = function () {
                    isShown = false;
                    popup.style.display = "none";
                }

                window.addEventListener('click', hidepopup, true);
                window.addEventListener('blur', hidepopup, true);
                return popup;
            }
        }

        var oPopup = window.createPopup;

        function findPos(obj, foundScrollLeft, foundScrollTop) {
            var curleft = 0;
            var curtop = 0;
            if (obj.offsetLeft) curleft += parseInt(obj.offsetLeft);
            if (obj.offsetTop) curtop += parseInt(obj.offsetTop);
            if (obj.scrollTop && obj.scrollTop > 0) {
                curtop -= parseInt(obj.scrollTop);
                foundScrollTop = true;
            }
            if (obj.scrollLeft && obj.scrollLeft > 0) {
                curleft -= parseInt(obj.scrollLeft);
                foundScrollLeft = true;
            }
            if (obj.offsetParent) {
                var pos = findPos(obj.offsetParent, foundScrollLeft, foundScrollTop);
                curleft += pos[0];
                curtop += pos[1];
            } else if (obj.ownerDocument) {
                var thewindow = obj.ownerDocument.defaultView;
                if (!thewindow && obj.ownerDocument.parentWindow)
                    thewindow = obj.ownerDocument.parentWindow;
                if (thewindow) {
                    if (!foundScrollTop && thewindow.scrollY && thewindow.scrollY > 0) curtop -= parseInt(thewindow.scrollY);
                    if (!foundScrollLeft && thewindow.scrollX && thewindow.scrollX > 0) curleft -= parseInt(thewindow.scrollX);
                    if (thewindow.frameElement) {
                        var pos = findPos(thewindow.frameElement);
                        curleft += pos[0];
                        curtop += pos[1];
                    }
                }
            }
            return [curleft, curtop];
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

    function UpdateBaseUnit(ID, ProductVersionID) {
        var url = "/IPulsar/Features/FeatureProperties.aspx?FeatureID=" + ID + "&ProductVersionID=" + ProductVersionID + "&PlatformTab=Y";
        window.open(url, "_blank", "resizable=yes,menubar=no,scrollbars=no,toolbar=no,top=" + GetWindowSize('top') + ",left=" + GetWindowSize('left') + ",width=" + adjustWidth(60) + ",height=" + adjustHeight(80) + "");
        /*returnvalue = window.showModalDialog("/IPulsar/Features/FeatureProperties.aspx?FeatureID=" + ID + "&ProductVersionID=" + ProductVersionID + "&PlatformTab=Y", 'Update Base Unit', "dialogWidth: " + DlgWidth + "px;dialogHeight:" + DlgHeight + "px;edge: Sunken;center:Yes; help: No;maximize:no;resizable: no;status: No; scroll:no");
        if (returnvalue) {
            RefreshPage(ID);
        }*/
    }

    function AddPlatformBaseUnit(ID, ProductVersionID) {
        var url = "/IPulsar/Features/FeatureCreate.aspx?PlatformTab=Y&PlatformID=" + ID + "&ProductVersionID=" + ProductVersionID;
        modalDialog.open({ dialogTitle: 'Add New Base Unit', dialogURL: '' + url + '', dialogHeight: 400, dialogWidth: 600, dialogResizable: true, dialogDraggable: true });
        /*var returnvalue = window.showModalDialog("/IPulsar/Features/FeatureCreate.aspx?PlatformTab=Y&PlatformID=" + ID + "&ProductVersionID=" + ProductVersionID, "Add Base Unit", "dialogWidth:600px;dialogHeight:200px;edge: Sunken;center:Yes; help: No;maximize:no;resizable: no;status: No; scroll:no");*/
    }

    function AddPlatformBaseUnitResult(returnvalue){
        if (returnvalue) {
            window.open("/IPulsar/Features/" + returnvalue + "&PlatformTab=Y", "_blank", "resizable=yes,menubar=no,scrollbars=no,toolbar=no,top=" + GetWindowSize('top') + ",left=" + GetWindowSize('left') + ",width=" + adjustWidth(60) + ",height=" + adjustHeight(80) + "");
        }
    }

    function RefreshPage(ID) {
        window.returnValue = 1;
        window.close();
    }

    function adjustWidth(percent) {
        return screen.width * (percent / 100);
    }

    function adjustHeight(percent) {
        return screen.height * (percent / 100);
    }

    //*****************************************************************
    //Description:  Code that runs when page loads
    //Function:     window_onload();
    //Modified By:  10/12/2016 - Harris, Valerie     
    //*****************************************************************
    function window_onload() {
        //Add modal dialog code to body tag: ---
        modalDialog.load();


        //add beforeclose to modalparent 
        window.parent.$("#modal_dialog").dialog({
            beforeClose: function (ev, ui) {
                window.parent.ViewPlatformBaseUnits_return();
                return true;
                }
        });
    }
    //-->
    </script>
    <link rel="stylesheet" type="text/css" href="../../style/wizard%20style.css">
</head>
<body onload="window_onload();">

     <div style="margin:4px 0px 0px 10px; height: 300px; width:100%; background-color: white" id="divPlatform">              
        
         <div style="font-size:x-small; margin-bottom:.5em">
             <span style="font-weight:bold">Base Unit Group:&nbsp;</span><%=request("PlatformName")%>
         </div>
         <div style="font-size:x-small; margin-bottom:.25em; float:right">
             <a href="javascript: AddPlatformBaseUnit(<%=request("ID")%>, <%=request("ProductVersionID")%>);">Add New Base Unit</a>
         </div>
  
           <table width="100%" id="TablePlatform" style="border:1px solid #ccc; margin-top:5px;">
                        <thead>
                            <tr style="position: relative; top: expression(document.getElementById('divPlatform').scrollTop-3);">
                                <td width="10" style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset; border-bottom: 1px outset"
                                    bgcolor="#c9ddff">Base&nbsp;Unit/Feature&nbsp;ID
                                </td>
                                <td width="180" style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset; border-bottom: 1px outset"
                                    bgcolor="#c9ddff">Base&nbsp;Unit&nbsp;Name
                                </td>
                                <td width="100" style="border-right: 1px outset; border-top: 1px outset; border-left: 1px outset; border-bottom: 1px outset"
                                    bgcolor="#c9ddff">Status
                                </td>
                            </tr>
                        </thead>
                        <%
	    set cn = server.CreateObject("ADODB.Connection")
	    cn.ConnectionString = Session("PDPIMS_ConnectionString")
	    cn.Open

	    set rs = server.CreateObject("ADODB.recordset")
	    rs.ActiveConnection = cn

       if request("ID") <> "" then

		    rs.open  "usp_GetPlatformBaseUnits " & clng(request("ID")) ,cn,adOpenForwardOnly
            do while not rs.eof
                response.write "<tr onmouseover=""return rows_onmouseover()"" onmouseout=""return rows_onmouseout()"" onclick=""return UpdateBaseUnit(" & rs("FeatureID") & "," & request("ProductVersionID") & ");"" >"
                response.write "<td class=""cell"">" & rs("FeatureID") & "&nbsp;</td>"
                response.write "<td nowrap class=""cell"">" & rs("MktNameMaster")  & "&nbsp;</td>"
                response.write "<td nowrap class=""cell"">" & rs("FeatureStatus")  & "&nbsp;</td>"
                response.write "</tr>"
    		    rs.MoveNext
    	    loop
	        rs.Close
       end if
%>
        </table>
    </div>

    <%
    set rs = nothing
    cn.close
    set cn= nothing

%>	
</body>
</html>
