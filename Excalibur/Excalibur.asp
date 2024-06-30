<%   

  dim LoginUser
  dim cn
  dim rs
  dim cnString
  dim preferedLayoutInCookie
  dim pulsar2LayoutName
  preferedLayoutInCookie = Request.Cookies("PreferredLayout2")
  pulsar2LayoutName = "pulsar2"

  If Session("PDPIMS_ConnectionString") = "" Then
                Session.Abandon
                Response.Write "<html><body><h2>connection string not set.</h2></body></html>"
                Response.End
  End If
  
  cnString= Session("PDPIMS_ConnectionString") 
  set cn = server.CreateObject("ADODB.Connection")
  cn.ConnectionString = cnString
  cn.Open

  set rs = server.CreateObject("ADODB.recordset")
  
                dim CurrentUser
                dim CurrentUserID
                dim CurrentUserPartner
                dim LoginOK
                CurrentUser = lcase(Session("LoggedInUser"))

                LoginOK = true

                if instr(currentuser,"@") = 0 then
                                if instr(currentuser,"\") = 0 then 
                                                LoginOK = false
                                elseif  not (left(CurrentUser,13)= "houhpqexcal03" or left(CurrentUser,13)= "houhpqexcal05" or left(CurrentUser,13)="houphqexcal02" or left(CurrentUser,8)= "americas" or left(CurrentUser,7)= "atlanta" or left(CurrentUser,8)= "atlanta2" or left(CurrentUser,11)= "col-springs" or left(CurrentUser,4)= "emea" or left(CurrentUser,11)= "asiapacific" or left(CurrentUser,7)= "asiapac" or left(CurrentUser,5)= "boise" or left(CurrentUser,7)= "europe2" or left(CurrentUser,7)= "europe1" or left(CurrentUser,7)= "cpqprod" or left(CurrentUser,9)= "palo-alto" or left(CurrentUser,9)= "vancouver"  or left(CurrentUser,4)= "auth") then
                                                LoginOK = false
                                end if
                end if
                
                if LoginOK then 
                                                'Get User
                                                dim CurrentDomain
                                                CurrentUser = lcase(Session("LoggedInUser"))
                                                CurrentDomain = ""

                                                                if instr(currentuser,"\") > 0 then
                                                                                CurrentDomain = left(currentuser, instr(currentuser,"\") - 1)
                                                                                Currentuser = mid(currentuser,instr(currentuser,"\") + 1) 
                                                                end if
                                                                
                                                                set cm = server.CreateObject("ADODB.Command")
                                                                Set cm.ActiveConnection = cn
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
                                                                                CurrentUserPartner = rs("PartnerID") & ""
                                                                                
                                                                                'if orginal name and current name is not the same we display current user (impersonate)
                                                                                if rs("originalName") = rs("CurrentName") then
                                                                                                LoginUser = rs("originalName")
                                                                                else
                                                                                                LoginUser = rs("originalName") & " (" & rs("CurrentName") & ")"
                                                                                end if

                                                                                if trim(CurrentUserPartner) = "9" then
                                                                                                Response.Redirect "./mobilese/modusmain.asp"
                                                                                end if
                                                                end if
                                                rs.Close
                        if currentuserid = 0 then
                            Response.Redirect "/pulsar/user/loginfailed"
                        else
                                Dim landingPageUrl
                                landingPageUrl = "/pulsarplus"
    
                                if preferedLayoutInCookie = "" or preferedLayoutInCookie = pulsar2LayoutName then 
                                    landingPageUrl  = "/" & pulsar2LayoutName & "/"
                                End if                            

                                if (Request.QueryString("classictodayPage") = "" or Request.QueryString("classictodayPage") <> 1) then
                                       if Request.QueryString("path") = "" or Request.QueryString("path") = "/" then
                                                Response.Redirect landingPageUrl 
                                            else 
                                                if Request.QueryString("path")="/Excalibur/mobilese/today/today.asp" then
                                                    Response.Redirect landingPageUrl
                                                end if
                                            end if                            
								end if 
								
								
								
                                                if trim(CurrentUserPartner) = "" then
                                                                TreePageName = "tree_vendor.asp"
                                                elseif trim(CurrentUserPartner) = "1" then
                                                                TreePageName = "tree.asp"
                                                else
                                                                rs.open "spgetPartnerType " & CurrentUserPartner,cn
                                                                if rs.eof and rs.bof then
                                                                                TreePageName = "tree_vendor.asp"
                                                                elseif trim(rs("PartnerTypeID") & "") = "1" or trim(rs("PartnerTypeID") & "") = "2" then
                                                                                TreePageName = "tree.asp"
                                                                else
                                                                                TreePageName = "tree_vendor.asp"
                                                                end if                                     
                                                                rs.close
                                                end if
                                               
                                set rs = nothing     
        
                                cn.Close
                                set cn = nothing
                                                

%>
<!doctype html>
<html style="height: 100%;">
<head>
    <meta http-equiv="X-UA-Compatible" content="IE=8" />
    <style>  
        #overlay {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background-color: #000;
            filter: alpha(opacity=66);
            opacity: 0.66;
            z-index: 20000;
        }

            #overlay p {
                margin: auto;
                padding-top: 100px;
                width: 50%;
                color: #fff;
                font-size: 24px;
                opacity: 1;
                text-align: center;
                font-weight: bold;
            }
    </style>
    <title>Pulsar</title>
    <link rel="shortcut icon" href="favicon.ico">
    <link href="style/shared.css" type="text/css" rel="stylesheet" />
    <link href="includes/client/jqueryui/jquery-ui-1.11.4/start/jquery-ui.min.css" rel="stylesheet" />
    <script src="includes/client/jquery-1.11.0.min.js" type="text/javascript"></script>
    <script src="includes/client/jqueryui/jquery-ui-1.11.4/jquery-ui.min.js" type="text/javascript"></script>
    <script src="/IPulsar/library/scripts/HoustonTime.js"></script>
    <script src="/IPulsar/library/scripts/Popup.js"></script>
    <script src="Scripts/shared_functions.js"></script>
    <script type="text/javascript">
        var iframeResize = false;

        $(document).ready(function () {

            if (!(window.navigator.userAgent.indexOf("MSIE ") != -1 && (window.navigator.userAgent.indexOf("Trident/7.0") != -1 || window.navigator.userAgent.indexOf("Trident/8.0") != -1))) {
                $('body').append('<div id="overlay"><p>IE11 is the only browser supported by Pulsar.<br/>Please use IE11 in compatibility view.<br/>Thank you.</p></div>')
                return;
            }

            $("body").css("overflow", "hidden");

            var menuUrl = '/pulsarplus/appheader/appheader/GetExcaliburPageHeader'
            var isPulsar2 = $("#hdnPreferedLayout").val().toLowerCase() != "pulsarplus" ? true : false;
        
            if (isPulsar2) {
                menuUrl = '/pulsar2/home/DisplayLayout?isCalledFromExcalibur=true';
            }

            $.ajax({
                url: menuUrl,
                dataType: 'html',
                cache: false,
                success: function (data) {
                    if (isPulsar2) {
                        $("#Wrapper").prepend(data);
                    }
                    else {
                        $("#pageHeader").html(data);
                    }
                },
                error: function (xhr, ajaxOptions, thrownError) {
                    alert('Status: ' + xhr.status + '\nError: ' + thrownError + '\nOptions: ' + ajaxOptions);
                }
            });

            var windowHeight = $(window).height();
            var windowWidth = $(window).width();
            document.getElementById("tbWrapper").style.height = (windowHeight - 51) + "px";
            document.getElementById("LeftWindowWrapper").style.height = (windowHeight - 51) + "px";
            document.getElementById("RightWindowWrapper").style.height = (windowHeight - 51) + "px";
            $("#RightWindowWrapper").width((windowWidth - $("#LeftWindowWrapper").width() - 5) + "px");


            $(window).resize(function () {

                if (!iframeResize) {
                    var windowHeight = $(window).height();
                    var windowWidth = $(window).width();
                    document.getElementById("tbWrapper").style.height = (windowHeight - 51) + "px";
                    document.getElementById("LeftWindowWrapper").style.height = (windowHeight - 51) + "px";
                    document.getElementById("RightWindowWrapper").style.height = (windowHeight - 51) + "px";

                }

                $(".ui-dialog-content:visible").each(function () {
                    $(this).dialog("option", "position", $(this).dialog("option", "position"));
                });
            });

            var path = GetParameterValues('path');
            var location = GetParameterValues('location');
            if (path) {
                path = decodeURIComponent(path);
                if (location == "pop")
                    launchExcalJqueryModal(path, true, false)
                else
                    launchExcalLink(path);
            }
            else {
                document.getElementById("RightWindow").src = "/Excalibur/mobilese/today/today.asp";
            }

            function GetParameterValues(param) {
                var url = window.location.href.slice(window.location.href.indexOf('?') + 1).split('&');
                for (var i = 0; i < url.length; i++) {
                    var urlparam = url[i].split('=');
                    if (urlparam[0] == param) {
                        return urlparam[1];
                    }
                }
            }

        });

        function launchExcalLinkModal(pageUrl, height, width, showProduct, showComponent) {
            var Height = $(window).height() - 200;
            var Width = ($(window).width() * 60) / 100;

            var strID = window.showModalDialog(pageUrl, "", "dialogWidth:" + Width + "px;dialogHeight:" + Height + "px;edge: Sunken;center:Yes; help: No;resizable: Yes;status: No");

            if (typeof (strID) != "undefined") {
                if (showProduct) {
                    window.document.getElementById("RightWindow").src = "/Excalibur/pmview.asp?Class=1&ID=" + strID;
                } else if (showComponent) {
                    window.location = "/Pulsar/Root/Component/" + strID;
                }

            }
        }
        function launchExcalJqueryModal(pageUrl, showProduct, showComponent) {
            var windowHeight = window.document.body.clientHeight - 200;
            var windowWidth = (window.document.body.clientWidth * 60) / 100;
            determineLoadedPage(pageUrl);
            window.ShowPropertiesDialog(pageUrl, "Add New", windowWidth, windowHeight);
        }
        function determineLoadedPage(baseUrl) {

            $("ul.lmTop li.lmItem a").each(function () { $(this).parent().removeClass("lmitem-arrow"); });

            var rightWindow = window.document.getElementById('RightWindow');
            var url;

            if (rightWindow != undefined && rightWindow.src != "") {
                url = rightWindow.src.replace(window.top.location.href, "");
            } else {
                url = window.top.location.pathname;
            }

            if (window.top.location.search != null && window.top.location.search != "") { url += window.top.location.search; }

            if (baseUrl != null) { url = baseUrl; }

            $(document).ready(function () {
                $("ul.lmTop li.lmItem a[href='" + url + "']").addClass("lmitem-arrow");
                $("ul.lmTop li.lmItem a[onclick*='" + url + "']").addClass("lmitem-arrow");
            });
        }
        function launchExcalLink(pageUrl) {

            var rightWindow = window.document.getElementById("RightWindow");

            if (rightWindow != undefined) {

                rightWindow.src = pageUrl;
                try { window.history.pushState({}, "Excalibur", "/Excalibur/"); }
                catch (e) { }

            }
        }

        function ShowPropertiesDialog(QueryString, Title, DlgWidth, DlgHeight) {
            if ($(window).height() < DlgHeight) DlgHeight = $(window).height() + 100;
            if ($(window).width() < DlgWidth) DlgWidth = $(window).width();
            $("#iframeDialog").dialog({
                width: DlgWidth, height: DlgHeight, resizable: true, resize: 'auto'
            });
            $("#modalDialog").attr("width", "98%");
            $("#modalDialog").attr("height", "98%");
            $("#modalDialog").attr("src", QueryString);
            $("#iframeDialog").dialog("option", "title", Title);
            $("#iframeDialog").dialog("open");
            $("#iframeDialog").dialog("option", "position", "center");

            globalVariable.save('add', 'product_prop_view');
        }

        function ClosePropertiesDialog(strID, showProduct, showComponent) {
            $("#modalDialog").attr("src", "");
            $("#iframeDialog").dialog("close");

            if (typeof (strID) != "undefined") {
                if (showProduct) {
                    document.getElementById("RightWindow").src = "/Excalibur/pmview.asp?Class=1&ID=" + strID;
                } else if (showComponent) {
                    window.location = "/Pulsar/Root/Component/" + strID;
                } else if (strID == "RefreshLeftTree") {
                    var iframe = document.getElementById("LeftWindow");
                    iframe.src = iframe.src;
                }
            }
        }

    </script>

    <style>
        body {
            font-family: Verdana, arial;
        }

        span.ui-dialog-title {
            font-size: 12px !important;
        }

        .ui-dialog .ui-resizable-se {
            width: 14px;
            height: 14px;
            right: 3px;
            bottom: 3px;
            background-position: -80px -224px;
        }
    </style>
</head>
<body style="margin: 0; height: 100%; overflow-y: hidden;">
    <%if preferedLayoutInCookie = "" or preferedLayoutInCookie = pulsar2LayoutName then %>
        <div id="Wrapper" class="wrapper">
        <div id="content-wrapper" class="d-flex flex-column" style="overflow: hidden">
            <table id="tbWrapper" style="width: 100%; padding: 0px; margin: 0px;">
                <tr>
                    <td style="width: 0%; visibility: hidden;">
                        <div id="LeftWindowWrapper">
                        </div>
                    </td>
                    <td>
                        <div id="RightWindowWrapper">
                            <iframe id="RightWindow" frameborder="0" seamless="seamless" scrolling="auto" style="width: 100%; height: 100%; border: 0;"></iframe>
                        </div>
                    </td>
                </tr>
            </table>
        </div>
    </div>
    <%else %>
            <div id="pageHeader" style="height: 50px;"></div>
            <div style="overflow: hidden">
                <table id="tbWrapper" style="width: 100%; padding: 0px; margin: 0px;">
                    <tr>
                        <td style="width: 0%; visibility: hidden;">
                            <div id="LeftWindowWrapper">
                            </div>
                        </td>
                        <td>
                            <div id="RightWindowWrapper">
                                <iframe id="RightWindow" frameborder="0" seamless="seamless" scrolling="auto" style="width: 100%; height: 100%; border: 0;"></iframe>
                            </div>
                        </td>
                    </tr>
                </table>
            </div>
    <%end if %>
    <div style="display: none;">
        <div id="iframeDialog" title="ExtendTables">
            <iframe frameborder="0" name="modalDialog" id="modalDialog"></iframe>
        </div>
    </div>
    <input type="hidden" id="hdnPreferedLayout" value="<%=preferedLayoutInCookie%>" />
</body>
</html>
<%
  end if
else
    Response.Redirect "/pulsar/user/loginfailed"
end if
%>
