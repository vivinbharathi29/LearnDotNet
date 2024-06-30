<%@  language="VBScript" %>

<%
	
    Response.Buffer = True
    Response.ExpiresAbsolute = Now() - 1
    Response.Expires = 0
    Response.CacheControl = "no-cache"
	  
    dim cn
	dim rs
	dim blnFound 
	dim i
    dim j
	dim cm
	dim p

    dim bitFusion
    dim strVersionId
    dim strCycleIds, arrCycleIds
    dim strCycleNames, arrCycleNames
    dim strGroupProductIds, arrGroupProductIds
    dim strGroupProductNames, arrGroupProductNames
    dim strMsgNoData
    dim strCmdNextDisabled
    dim strMsgFusion
%>

<html>
<title>Multiple Products RTM</title>
<head>
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
    <base target='_self'>
    <script id="clientEventHandlersJS" language="javascript">
<!-- #include file = "../includes/Date.asp" -->
  
///////////////////////////////////////////////////////////////


        var strPriProdID;
        var strID;
        var strIDs;
        var arrIDs;
        var intMaxProd;
        var arrlistProductAll;


        strID = "";
        strIDs = "";
        intMaxProd = 5; // limite max 5 product selected

        function cmdCancel_onclick() {
            var pulsarplusDivId = document.getElementById("pulsarplusDivId");
            if (pulsarplusDivId.value != undefined && pulsarplusDivId.value != "") {
                // For Closing current popup
                parent.window.parent.closeExternalPopup();
            }
            else {
                if (window.parent.document.getElementById('modal_dialog')) {
                    window.parent.modalDialog.cancel(false);
                } else {
                    window.close();
                }
            }
        }

        function setProdId(id) {
            strPriProdID = id;
        }

        function getAllProducts() {
            arrlistProductAll = document.getElementById("tbGroupProducts").getElementsByTagName("input");
        }

        function HowManyInArray(str, arrStr) {
            var intH;
            intH = 0;
            for (var i = 0; i < arrStr.length; i++) {
                if (str == arrStr[i])
                    intH += 1;
            }
            return intH;
        }

        function onCheckProduct(elmChk) {

            getAllProducts();
           
            checkSameProducts(elmChk);

            var intChecked;
            intChecked = 0;

            arrIDs = new Array();
           
            var chkProd
            for (var i = 0; i < arrlistProductAll.length; i++) {
                chkProd = arrlistProductAll[i];
                strID = chkProd.value.toString();

                if (chkProd.checked && HowManyInArray(strID, arrIDs) == 0) {
                    arrIDs.push(strID);
                }
            }
            arrIDs.sort();
            strIDs = arrIDs.toString();
            intChecked = arrIDs.length;

            // limite intMaxProd products selected
            if (intChecked == intMaxProd) {
                setDisable(true);
            } else {
                setDisable(false);
            }

        }

        // once a product selected, the same product in other Programs will be selected.
        function checkSameProducts(elmChk) {
            for (var i = 0; i < arrlistProductAll.length; i++) {
                chkProd = arrlistProductAll[i];
                
                if (chkProd.value.toString() == elmChk.value.toString()) {
                    chkProd.checked = elmChk.checked;
                }
                
            }
        }

        function setDisable(boolOnOff) {
            for (var i = 0; i < arrlistProductAll.length; i++) {
                chkProd = arrlistProductAll[i];
                    if (!chkProd.checked) {
                        chkProd.disabled = boolOnOff;
                    }
            }

        }


        // send IDs by url parameter.
        function cmdNext_onclick() {
            
            if (strIDs.toString() == "") {
                strIDs = strPriProdID.toString();
            }
            
            var linkNext = document.getElementById("linkNext");
            linkNext.href = "MultipleRTM.asp?ID=" + strPriProdID + "&IDS=" + strIDs + "&pulsarplusDivId=" + pulsarplusDivId.value;
            //linkNext.href = "MultipleRTMTest.asp?ID=" + strPriProdID + "&IDS=" + strIDs;
            linkNext.click();

        }
 
    </script>
    <script>
        setProdId(<%=trim(request("ID")) %>);
    </script>

<style>
    A:visited {
        COLOR: blue;
    }

    A:hover {
        COLOR: red;
    }

    .EmbeddedTable TBODY TD {
        FONT-FAMILY: Verdana;
    }

    .EmbeddedTable TBODY TD {
        Font-Size: xx-small;
    }

    input {
        FONT-SIZE: 10pt;
        FONT-FAMILY: Verdana;
    }

    textarea {
        FONT-SIZE: 10pt;
        FONT-FAMILY: Verdana;
    }

    .ImageTable TBODY TD {
        BORDER-TOP: gray thin solid;
        FONT-SIZE: xx-small;
        FONT-FAMILY: verdana;
    }

    .ImageTable TH {
        FONT-SIZE: xx-small;
        FONT-FAMILY: verdana;
    }

    .imagerows TBODY TD {
        BORDER-TOP: none;
        FONT-SIZE: xx-small;
        FONT-FAMILY: verdana;
    }

    .imagerows THEAD TD {
        BORDER-TOP: none;
        FONT-SIZE: xx-small;
        FONT-FAMILY: verdana;
    }

    body {
        margin-top: 20px;
        margin-bottom: 20px;
        margin-right: 20px;
        margin-left: 10px;
    }


</style>
<link href="../style/wizard%20style.css" type="text/css" rel="stylesheet">
</head>
<body bgcolor="ivory">
    <input type="hidden" id="pulsarplusDivId" name="pulsarplusDivId" value="<%=request("pulsarplusDivId")%>">
    
    <font size="4" face="verdana"><b>Multiple Product RTM Wizard </b></font>
    <br>
    <font size="2" face="verdana"><b><label ID="lblTitle">Select Products:</label></b></font>

    <form id="formProduct" name="formProduct" method="post" action="MultipleRTM.asp">

        <div id="tabProgram" style="display: inline;">


            <%
    set cn = server.CreateObject("ADODB.Connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString") 
	cn.Open
	set rs = server.CreateObject("ADODB.recordset")
    rs.CursorType = adOpenForwardOnly
	rs.LockType=AdLockReadOnly

    strVersionId = trim(request("ID"))   

    dim strSQL
    
    strSQL = "spGetProductIsFusion " & clng(strVersionId)
                rs.open strSQL,cn         

    bitFusion = 0
    strMsgFusion=" Not Fusion "
    do while not rs.eof
        if rs("fusion") then 
            bitFusion = 1
            strMsgFusion =" Fusion "
        end if
        rs.movenext
    loop

    rs.close 
                
    strSQL = "spGetProgramsByProductVersionId " & clng(strVersionId)  
    'strSQL = "SELECT pg.ID, pg.FullName, pp.ProductVersionID FROM Program pg INNER JOIN Product_Program pp ON pg.ID = pp.ProgramID WHERE pp.ProductVersionID = " + strVersionId + " ORDER BY pg.FullName"                                     
    rs.open strSQL,cn        

    if not(rs.eof and rs.bof) then
        strMsgNoData = ""
        strCmdNextDisabled =""
    else
        strCmdNextDisabled =" disabled "
        strMsgNoData = "<table><tr><td width=""120"">&nbsp;</td><td><font size=2 color=red face=verdana>There is no productgroup on this product.</font></td></tr></table>"
    end if            
            
    if not(rs.eof and rs.bof) then
        do while not rs.eof
            strCycleIds = strCycleIds & cstr(rs("ID")) & ","
            strCycleNames = strCycleNames & rs("FullName") & ","
            rs.movenext
        loop
    end if

    rs.close            
      
    if  strCycleIds<>"" then
        strCycleIds = left(strCycleIds, len(strCycleIds)-1)
        arrCycleIds = split(strCycleIds,",")
    end if     
    if  strCycleNames<>"" then
        strCycleNames = left(strCycleNames, len(strCycleNames)-1)
        arrCycleNames = split(strCycleNames,",")
    end if    
            
            %>
<%
    if  strCycleIds<>"" then
     %>
           
            <table border="1" cellpadding="2" cellspacing="0" bgcolor="cornsilk" bordercolor="tan" width="100%" style="display: inline;" id="tbGroupProducts">
                <tr>
                    <td width="120">
                        <font size="2" face="verdana">
                        <b>Group</b>
                    </font>
                    </td>
                    <td>
                        <font size="2" face="verdana">
                        <b>Select Multiple RTM Products:</b> <!-- <%=strMsgFusion %>-->
                    </font>
                    </td>
                </tr>
                

                <% for i = lbound(arrCycleNames) to ubound(arrCycleNames) %>
                <tr>
                    <td width="120">
                        <span id="<%=arrCycleIds(i) %>"><%=arrCycleNames(i)  %></span>
                    </td>
                    <td>
                        <% 
                
                strSQL = "spGetProductVersionsByProgramId2RTMs " & clng(arrCycleIds(i)) & "," & bitFusion
                'strSQL = "SELECT pv.ID, pv.DOTSName,pv.Fusion, pp.ProgramID FROM Product_Program pp INNER JOIN ProductVersion pv ON pp.ProductVersionID = pv.ID WHERE pp.ProgramID = " + arrCycleIds(i) + " and pv.Fusion = " + CStr(bitFusion) + " ORDER BY pv.DOTSName"
                rs.open  strSQL,cn        

                if (rs.eof and rs.bof) then
                    response.write "No other product in this group."
                end if            
            
                do while not rs.eof
                    if trim(cstr(rs("ID"))) = strVersionId then
                        %>
                        <input id='lstProduct<%=cstr(rs("ID")) %>' type="checkbox" name="lstProduct" onclick="onCheckProduct(this)" productname="<%=rs("DOTSName") %>" value="<%=cstr(rs("ID")) %>" primary="primary" checked disabled><%=rs("DOTSName") %><br />
                        <%
                    else
                        %>
                        <input id='lstProduct<%=cstr(rs("ID")) %>' type="checkbox" name="lstProduct" onclick="onCheckProduct(this)" productname="<%=rs("DOTSName") %>" value="<%=cstr(rs("ID")) %>"><%=rs("DOTSName") %><br />
                        <%
                    end if
                    rs.movenext
                loop

                rs.close  
                        %>
                    </td>
                </tr>
                <% next %>
            </table>

<%
    else
        response.write strMsgNoData
    end if    
     %>

        </div>

    </form>
    
    <hr width="100%">

    <table border="0" cellspacing="1" cellpadding="1" align="right">
        <tr>

            <td>
                <input type="button" value="Cancel" id="cmdCancel" name="cmdCancel" onclick="cmdCancel_onclick()"></td>
            <td width="10"></td>
            <td></td>
            <td>
                <input type="button" value="Next&gt;&gt;" id="cmdNext" name="cmdNext" onclick="cmdNext_onclick()" <%=strCmdNextDisabled %>>
            </td>
            <td width="10">
                <a id="linkNext" style="visibility:hidden;" href="#">NEXT</a>

            </td>
            <td>&nbsp;</td>


        </tr>
    </table>

</body>
</html>
