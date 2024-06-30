<%@ LANGUAGE="VBSCRIPT" %>
<!--#include file="../includes/Security.asp"-->
<!--#include file="../includes/DataWrapper.asp"-->
<!--#include file="../includes/no-cache.asp"-->
<%

Server.ScriptTimeout=480


Dim AppRoot

'##############################################################################	
'
' Create Security Object to get User Info
'
Dim regEx
Set regEx = New RegExp
regEx.Global = True

' COULD USE JUST Request("name"), but for now...
regEx.Pattern = "[^0-9]"
Dim strPVID 
If Request.Form("Action") = "save" Then
    strPVID=regEx.Replace(Request.Form("PVID"),"")
Else
    strPVID=regEx.Replace(Request.QueryString("PVID"),"")
End If

regEx.Pattern = "[^0-9-]"
Dim strSFPN
If Request.Form("Action") = "save" Then
    strSFPN=trim(Request.Form("SFPN"))
Else
    strSFPN=trim(Request.QueryString("SFPN"))
End If

regEx.Pattern = "[^0-9,]"
Dim strSKIDs
If Request.Form("Action") = "save" Then
    strSKIDs=regEx.Replace(Request.Form("SKIDS"),"")
Else
    strSKIDs=regEx.Replace(Request.QueryString("SKIDS"),"")
End If

Dim arySKIDs
Dim intNumSels: intNumSels=0

IF(Len(Trim(strSKIDs))>0) THEN
    arySKIDs=Split(strSKIDs,",")
    intNumSels=UBound(arySKIDs)+1
END IF


Dim csrLevelOptions
Dim dispositionOptions
Dim warrantyOptions
Dim stockAdviceOptions
Dim fsDate: fsDate=""


Dim m_IsSysAdmin
Dim m_IsGplm
Dim m_IsBomAnalyst
Dim m_EditModeOn
Dim m_UserFullName
	
	m_EditModeOn = False
	
	Dim Security
		
	Set Security = New ExcaliburSecurity
	
	m_IsSysAdmin = Security.IsSysAdmin()
    m_IsGplm = Security.UserInRole(strPVID, "GPLM")
    m_IsBomAnalyst = Security.UserInRole(strPVID, "SBA")

	m_UserFullName = Security.CurrentUserFullName()
	
	If m_IsSysAdmin Or m_IsGplm Or m_IsBomAnalyst Then
		m_EditModeOn = True
	End If
	
	If Not m_EditModeOn Then
		sMode = "view"
	End If

	Set Security = Nothing

'##############################################################################	

AppRoot = Session("ApplicationRoot")

If Request.Form("Action") = "save" Then
    Call UpdateSelectedRecords()
End If

Sub DisplaySelectedRecords

	Dim oDW: Set oDW=New DataWrapper
	Dim oConn: Set oConn=oDW.CreateConnection("PDPIMS_ConnectionString")
	Dim oComm: Set oComm=Server.CreateObject("ADODB.Command")
	Dim oRS: Set oRS=Server.CreateObject("ADODB.Recordset")
	Dim strDesc
    Dim strCatName


	oComm.ActiveConnection=oConn
    'Response.Write(strSKIDS & "<BR/>" & strSFPN & "<HR/>")
	oComm.CommandText="EXEC usp_SelectMultipleSpareKits '" & strSKIDS & "', '" & strSFPN & "'"
    'oComm.CommandText="EXEC usp_SelectMultipleSpareKitsORGONE '" & strSKIDS & "', '" & strSFPN & "'"
    'oRS.ActiveConnection=oConn
	Set oRS=oComm.Execute

	IF Not oRS.EOF THEN

		Do While Not oRS.EOF

		    strDesc=oRS("Description")

        	If(Len(Trim(strDesc))>40) Then 
                strDesc=mid(Trim(strDesc),1,40)
	        End If

            strCatName=oRS("CategoryName")

            'Response.Write("<div style='height: 25px'><span id='" & oRS("ID") & "' class='floatText'>" & strDesc & "</span></div>")

            Response.Write("<tr><td>"  & strDesc & "</td><td>" & strCatName & "</td></tr>") 

			oRS.MoveNext
		Loop

		oRS.Close
	END IF

	Set oRS=Nothing

	Set oComm=Nothing

	oConn.Close
	Set oConn=Nothing

	Set oDW=Nothing

End Sub


Sub InitializeOptions

    'CSR Level, Disposition, Warranty Labour Tier, and Local Stock Advice, and Default First Service Date

    Dim dw: Set dw=new DataWrapper
    Dim cn: Set cn=dw.CreateConnection("PDPIMS_ConnectionString")
    Dim cmd
    Dim rs
 
    '
    ' Get CSR Level List
    '
       
        Set cmd = dw.CreateCommandSP(cn, "usp_ListServiceCsrLevels")
        Set rs = dw.ExecuteCommandReturnRS(cmd)

        Do Until rs.Eof
            If CLng(currentKitCsrLevel) = CLng(rs("ID")) Then
                csrLevelOptions = csrLevelOptions & "<option value=""" & rs("ID") & """ selected>" & rs("CsrLevel") & " - " & rs("CsrDescription") & "</option>"
            Else
                csrLevelOptions = csrLevelOptions & "<option value=""" & rs("ID") & """>" & rs("CsrLevel") & " - " & rs("CsrDescription") & "</option>"
            End If
            rs.MoveNext
        Loop
        rs.close    

	    Set rs=Nothing
        Set cmd=Nothing


    '
    ' Get Disposition List
    ' 
        dispositionOptions = ""
        dispositionOptions = dispositionOptions & "<option value=""1"">1 - Disposable</option>"
        dispositionOptions = dispositionOptions & "<option value=""2"">2 - Repairable</option>"
        dispositionOptions = dispositionOptions & "<option value=""3"">3 - Return to Vendor</option>"
        dispositionOptions = dispositionOptions & "<option value=""4"">4 - Repair/Exhcange Only</option>"

    '
    ' Get Warranty Tier List
    ' 
        warrantyOptions = ""
        warrantyOptions = warrantyOptions & "<option value=""A"">A - 0 - 5 Minutes</option>"
        warrantyOptions = warrantyOptions & "<option value=""B"">B - 5 - 10 Minutes</option>"
        warrantyOptions = warrantyOptions & "<option value=""C"">C - 15 - 30 Minutes</option>"
        warrantyOptions = warrantyOptions & "<option value=""D"">D - No Reimbursment</option>"

    '
    ' Get Stock Advice List
    ' 
        stockAdviceOptions = ""
        stockAdviceOptions = stockAdviceOptions & "<option value=""1"">1 - Don't stock local (non SPOF and not likely to fail)</option>"
        stockAdviceOptions = stockAdviceOptions & "<option value=""2"">2 - Stock Strategically (non SPOF and likely to fail)</option>"
        stockAdviceOptions = stockAdviceOptions & "<option value=""3"">3 - Stock Local (SPOF and not likely to fail)</option>"
        stockAdviceOptions = stockAdviceOptions & "<option value=""4"">4 - Stock Local Critical (SPOF and likely to fail)</option>"

    '
    ' Get the Default First Service Date
    '
    fsDate = ""
    If TRIM(strPVID) <> "" Then
        Set cmd = dw.CreateCommandSp(cn, "usp_GetRSLFSDDefault")
        dw.CreateParameter cmd, "@p_ProductVersionId", adInteger, adParamInput, 0, CLng(strPVID)
        Set rs = dw.ExecuteCommandReturnRs(cmd)
        
        If Not rs.Eof Then
            fsDate = rs("Summary_Dt")
            If IsDate(fsDate) Then
		        fsDate = DateAdd("d", 14, fsDate)
            Else
                fsDate=""
            End If
            rs.close
        End If

        Set rs=Nothing
        Set cmd=Nothing

    End If

    cn.Close
    Set cn=Nothing

    Set dw=Nothing


End Sub


Function ConvertBoolean(strValue)
    Dim sRetValue

    Select Case LCase(Trim(strValue))
    Case "true", "on", "1": sRetValue="true"
    Case "false", "off", "0": sRetValue="false"
    Case Else: sRetValue="false"
    End Select

    ConvertBoolean=sRetValue

End Function


Sub UpdateSelectedRecords

    If Not m_EditModeOn Then
        Response.Write "<span style=""font:bold large verdana;color:darkred"">Insufficient Security Privileges</span>"
        Response.End
    End If

On Error Resume Next

    ' MAY NEED TO APPLY ESCAPE SEQUENCES TO PASSED DATA AND TO JAVASCRIPT

    ' CSR Level, Disposition, Warranty Labour Tier, Local Stock Advice, GEOS (4 checkboxes), First Service Dt., Notes, and RSL Comments.

    err.Clear()

    Dim oRegEx
    Set oRegEx = New RegExp
    oRegEx.Global = True

    Dim intErrorNumber: intErrorNumber=0
    Dim strErrorDesc:   strErrorDesc=""

    oRegEx.Pattern = "[^0-9-]"
    Dim strSFPN: strSFPN=trim(Request.Form("SFPN"))

    oRegEx.Pattern = "[^0-9,]"
    Dim strSKIDs: strSKIDs=oRegEx.Replace(Request.Form("SKIDS"),"")

    'Dim strUpdGEOS: strUpdGEOS=Request.Form("rdoUpdGEOS")
    Dim strAppNotes: strAppNotes=Request.Form("rdoNotes")
    Dim strAppComm: strAppComm=Request.Form("rdoComm")
    'Dim blnUpdGEOS: blnUpdGEOS=false
    Dim blnAppNotes: blnAppNotes=false
    Dim blnAppComm: blnAppComm=false

   ' IF strUpdGEOS="Y" THEN blnUpdGEOS=true
    IF strAppNotes="A" THEN blnAppNotes=true
    IF strAppComm="A" THEN blnAppComm=true

    Set oRegEx=Nothing

    Dim strUpdMatrix: strUpdMatrix=Request.Form("updateMatrix")

    Dim strSSKColumns: strSSKColumns="CsrLevelId|Disposition|WarrantyTier|LocalStockAdvice"
    Dim arySSKColumns: arySSKColumns=Split(strSSKColumns,"|")
    Dim strSSKValues: strSSKValues=Request.Form("spsCsrLevel") & "|" & Request.Form("spsDisposition") & "|" & Request.Form("spsWarranty") & "|" & Request.Form("spsLocalStockAdvice")
    Dim arySSKValues: arySSKValues=Split(strSSKValues,"|")

    Dim strSSKSFColumns: strSSKSFColumns="GeoNa|GeoLa|GeoApj|GeoEmea|FirstServiceDt|Notes|Comments"
    Dim arySSKSFColumns: arySSKSFColumns=Split(strSSKSFColumns,"|")
    Dim strSSKSFValues: strSSKSFValues=ConvertBoolean(Request.Form("spsGeosNa")) & "|" & ConvertBoolean(Request.Form("spsGeosLa")) & "|" & ConvertBoolean(Request.Form("spsGeosApj")) & "|" & ConvertBoolean(Request.Form("spsGeosEmea")) & "|" & Request.Form("spsFirstServiceDt") & "|" & Replace(Request.Form("spsNotes"),"'","''") & "|" & Replace(Request.Form("spsComments"),"'","''")
    Dim arySSKSFValues: arySSKSFValues=Split(strSSKSFValues,"|")

    Dim strSSKSQLPrefix: strSSKSQLPrefix="UPDATE ServiceSpareKit SET "
    Dim strSSKSQLCriteria: strSSKSQLCriteria=" WHERE ID IN(" & strSKIDs & ")"
    Dim strSSKSQLBody: strSSKSQLBody=""
    Dim strSSKSQLText: strSSKSQLText=""

    Dim strSSKSFSQLPrefix: strSSKSFSQLPrefix="UPDATE ServiceSpareKit_ServiceFamily SET "
    Dim strSSKSFSQLCriteria: strSSKSFSQLCriteria=" WHERE SpareKitId IN(" & strSKIDs & ") AND ServiceFamilyPn='" & strSFPN & "'"
    Dim strSSKSFSQLBody: strSSKSFSQLBody=""
    Dim strSSKSFSQLText: strSSKSFSQLText=""
    
    Dim intStartIdx
    Dim intCurrIdx
    Dim intIdx
    
    Dim strCurrFlag
    Dim strSSKUpdMatrix: strSSKUpdMatrix=Mid(strUpdMatrix,1,4)
    Dim strSSKSFUpdMatrix: strSSKSFUpdMatrix=Mid(strUpdMatrix,5,7)

    '**********************************************************************************************************************************************************************************************************
    ' Create the SQL statement to UPDATE the [ServiceSpareKit] table
    '**********************************************************************************************************************************************************************************************************
    intCurrIdx=0

    FOR intIdx=LBound(arySSKColumns) TO UBound(arySSKColumns)

        intCurrIdx=intCurrIdx+1
        strCurrFlag=Mid(strSSKUpdMatrix,intCurrIdx,1)

        IF strCurrFlag="1" THEN
            If LEN(strSSKSQLBody)=0 Then
                strSSKSQLBody = arySSKColumns(intIdx) & "='" & arySSKValues(intIdx) & "'"
            Else
                strSSKSQLBody=strSSKSQLBody & ", " & arySSKColumns(intIdx) & "='" & arySSKValues(intIdx) & "'"
            End If
        END IF

    NEXT

    IF(LEN(TRIM(strSSKSQLBody))>0) THEN
        strSSKSQLBody=strSSKSQLBody & ", LastUpdDate=GETDATE(), LastUpdUser='" & m_UserFullName & "'"
        strSSKSQLText=strSSKSQLPrefix & strSSKSQLBody & strSSKSQLCriteria
        'RESPONSE.WRITE("<HR>" & strSSKSQLText & "<HR>")
    ELSE
        'RESPONSE.WRITE("<HR> NO UPDATES TO ServiceSpareKit RECORDS<HR>")
    END IF
    '**********************************************************************************************************************************************************************************************************


    '**********************************************************************************************************************************************************************************************************
    ' Create the SQL statement to UPDATE the [ServiceSpareKit_ServiceFamily] table 
    '**********************************************************************************************************************************************************************************************************
    intCurrIdx=0

    FOR intIdx=LBound(arySSKSFColumns) TO UBound(arySSKSFColumns)

        intCurrIdx=intCurrIdx+1
        strCurrFlag=Mid(strSSKSFUpdMatrix,intCurrIdx,1)

        IF strCurrFlag="1" THEN

            ' ACCOMMODATE SPECIAL CASE WITH REGARD TO [Notes] & [Comments] (Append Or Overwrite)
            IF(intCurrIdx=6 AND blnAppNotes) THEN ' Notes
                If LEN(strSSKSFSQLBody)=0 Then
                    'strSSKSFSQLBody = arySSKSFColumns(intIdx) & "=COALESCE([" & arySSKSFColumns(intIdx) + "],'')+CHAR(13)+CHAR(10)+CONVERT(VARCHAR(30), CURRENT_TIMESTAMP)+' (CST) - '+'" & arySSKSFValues(intIdx) & "'"
                    strSSKSFSQLBody = arySSKSFColumns(intIdx) & "=COALESCE([" & arySSKSFColumns(intIdx) + "],'')+CHAR(13)+CHAR(10)+'" & arySSKSFValues(intIdx) & "'"
                Else
                    'strSSKSFSQLBody=strSSKSFSQLBody & ", " &  arySSKSFColumns(intIdx) & "=COALESCE([" & arySSKSFColumns(intIdx) + "],'')+CHAR(13)+CHAR(10)+CONVERT(VARCHAR(30), CURRENT_TIMESTAMP)+' (CST) - '+'" & arySSKSFValues(intIdx) & "'"
                    strSSKSFSQLBody=strSSKSFSQLBody & ", " &  arySSKSFColumns(intIdx) & "=COALESCE([" & arySSKSFColumns(intIdx) + "],'')+CHAR(13)+CHAR(10)+'" & arySSKSFValues(intIdx) & "'"
                End If                
            ELSEIF(intCurrIdx=7 AND blnAppComm) THEN ' Comments 
                If LEN(strSSKSFSQLBody)=0 Then
                    'strSSKSFSQLBody = arySSKSFColumns(intIdx) & "=COALESCE([" & arySSKSFColumns(intIdx) + "],'')+CHAR(13)+CHAR(10)+CONVERT(VARCHAR(30), CURRENT_TIMESTAMP)+' (CST) - '+'" & arySSKSFValues(intIdx) & "'"
                    strSSKSFSQLBody = arySSKSFColumns(intIdx) & "=COALESCE([" & arySSKSFColumns(intIdx) + "],'')+CHAR(13)+CHAR(10)+'" & arySSKSFValues(intIdx) & "'"
                Else
                    'strSSKSFSQLBody=strSSKSFSQLBody & ", " &  arySSKSFColumns(intIdx) & "=COALESCE([" & arySSKSFColumns(intIdx) + "],'')+CHAR(13)+CHAR(10)+CONVERT(VARCHAR(30), CURRENT_TIMESTAMP)+' (CST) - '+'" & arySSKSFValues(intIdx) & "'"
                    strSSKSFSQLBody=strSSKSFSQLBody & ", " &  arySSKSFColumns(intIdx) & "=COALESCE([" & arySSKSFColumns(intIdx) + "],'')+CHAR(13)+CHAR(10)+'" & arySSKSFValues(intIdx) & "'"
                End If                
            ELSE
                If LEN(strSSKSFSQLBody)=0 Then
                    strSSKSFSQLBody = arySSKSFColumns(intIdx) & "='" & arySSKSFValues(intIdx) & "'"
                Else
                    strSSKSFSQLBody=strSSKSFSQLBody & ", " & arySSKSFColumns(intIdx) & "='" & arySSKSFValues(intIdx) & "'"
                End If
            END IF

        END IF

    NEXT

    IF(LEN(TRIM(strSSKSFSQLBody))>0) THEN
        strSSKSFSQLBody=strSSKSFSQLBody & ", LastUpdDate=GETDATE(), LastUpdUser='" & m_UserFullName & "'"
        strSSKSFSQLText=strSSKSFSQLPrefix & strSSKSFSQLBody & strSSKSFSQLCriteria
        'RESPONSE.WRITE("<HR>" & strSSKSFSQLText & "<HR>")
    ELSE
        'RESPONSE.WRITE("<HR> NO UPDATES TO ServiceSpareKit_ServiceFamily RECORDS<HR>")
    END IF
    '**********************************************************************************************************************************************************************************************************

    Dim oDW: Set oDW=New DataWrapper
	Dim oConn: Set oConn=oDW.CreateConnection("PDPIMS_ConnectionString")
	Dim oComm: Set oComm=Server.CreateObject("ADODB.Command")
	Dim intAffectedSSKRecs
    Dim intAffectedSSKSFRecs
    Dim intTransCount

    ' Mark the Beginning of the Transaction
    intTransCount=oConn.BeginTrans()

	oComm.ActiveConnection=oConn
    
	IF(LEN(TRIM(strSSKSQLBody))>0) THEN
        oComm.CommandText=strSSKSQLText
	    intAffectedSSKRecs=oDW.ExecuteNonQuery(oComm)
    END IF
    
	IF(LEN(TRIM(strSSKSFSQLBody))>0) THEN
        oComm.CommandText=strSSKSFSQLText
	    intAffectedSSKSFRecs=oDW.ExecuteNonQuery(oComm)
    END IF
    
    If(err.number<>0) Then
        ' Error Occurred, save the error code and description to the respective variables
        intErrorNumber=err.number
        strErrorDesc=err.Description

        'Rollback the Transaction
        oConn.RollbackTrans()
    Else
	    ' Commit the Transaction	
	    oConn.CommitTrans()
    End If

	Set oComm=Nothing

	oConn.Close
	Set oConn=Nothing

	Set oDW=Nothing
   
   If(intErrorNumber<>0) Then
        Response.Write("ERROR OCCURRED: <BR/>ALL UPDATES HAVE BEEN ROLLED BACK<BR/>Number = " & Cstr(intErrorNumber) & "<BR/>Description = " & strErrorDesc & "<BR/>")
        Response.Write("<script language=""javascript"">var sDesc=""" & strErrorDesc & """; alert(""ERROR OCCURRED [ALL UPDATES HAVE BEEN ROLLED BACK]: ""+sDesc); window.returnValue=""cancel""; window.close();</script>")
    Else
        Response.Write("<script language=""javascript"">alert('Successfully updated selected record(s).'); window.returnValue=""save""; window.close();</script>")
    End If

End Sub


%>
<html>
<head>
<title>Spare Kit - Batch Update</title>

    <style type="text/css">
        body
        {
            font: xx-small verdana;
        }
        legend
        {
            font: bold x-small verdana;
            color: #004874;
        }
        .floatRight
        {
        	position: absolute;
            left: 250px;
            float: left;
            text-align: left;
        }
        .floatLeft
        {
            float: left;
            font: bold x-small verdana;
        }
        .floatText
        {
            position: absolute;
            left: 250px;
        }
        .display
        {
            text-align: left;
        }
        .link
        {
            font: bold x-small verdana;
            color: #004874;
            text-decoration: underline;
        }
        .link:Hover
        {
        	text-decoration: underline overline;
        	cursor: hand;
        }
        .linkEditFloatRight
        {
            float: right;
            text-align: right;
            font: bold x-small verdana;
            color: #004874;
            text-decoration: underline;
        }
        .linkEditFloatRight:Hover
        {
        	text-decoration: underline overline;
        	cursor: hand;
        }
        .inputBox
        {
            width: 350px;
        }
        .SelectedRecsTable
        {
            width: 100%;
            border-bottom: solid 1px black;
            border-collapse: collapse;
        }
        .SelectedRecsTable th
        {
            text-align: left;
            background: #004874;
            color: #ffffff;
        }
        .SelectedRecsTable td
        {
            text-align: left;
            border-bottom: solid 1px black;
        }
    </style>

    <script type="text/javascript" language="javascript">

        function cmdDate_onclick(target) {

            var strID;
            var txtDateField = document.getElementById(target);
            strID = window.showModalDialog("../MobileSE/Today/calDraw1.asp", txtDateField.value, "dialogWidth:300px;dialogHeight:220px;edge: Raised;center:Yes; help: No;resizable: No;status: No");
            if (typeof (strID) != "undefined") {
                txtDateField.value = strID;
            }
        }

        function setDefaultFSD(oChkBox, sDefValue) {
            var oFSD = document.getElementById("spsFirstServiceDt");
            
            if (oChkBox.checked) 
            {
                if (sDefValue.length > 0) 
                {
                    oFSD.value = sDefValue;
                }
                else 
                {
                    oChkBox.checked = false;
                    alert("No Default Value exists.");
                }
            }
            else 
            {
                oFSD.value = "";
            }
                        
            return;
        }

    </script>

</head>
<body>
<%
'****************************************************************************************************************************************************************************************************************************************************************************
If Request.Form("Action") <> "save" Then
    Call InitializeOptions()
%>
    <form id="frmMain" method="post" action="">
        <div style="width: 600px">
            <div id="KitDetails">
                <fieldset>
                    <legend>Batch Update - <%= m_UserFullName %></legend>
                    <div style="height: 25px">
                        <span class="floatLeft">CSR Level:</span>
                        <span id="spsCsrLevelEdit" class="floatRight">
                            <select class="inputBox" id="spsCsrLevel" name="spsCsrLevel">
                                <option value="0">-- Select CSR Level --</option>
                                <%= csrLevelOptions %>
                            </select>
                        </span> 
                    </div>
                    <div style="height:25px">
                        <span class="floatLeft">Disposition:</span>
                        <span id="spsDispositionEdit" class="floatRight">
                            <select class="inputBox" id="spsDisposition" name="spsDisposition">
                                <option value="0">-- Select Disposition --</option>
                                <%= dispositionOptions %>
                            </select>
                        </span>
                    </div>
                    <div style="height:25px">
                        <span class="floatLeft">Warranty Labour Tier:</span>
                        <span id="spsWarrantyEdit" class="floatRight">
                            <select class="inputBox" id="spsWarranty" name="spsWarranty">
                                <option value="0">-- Select Warranty Labour Tier --</option>
                                <%= warrantyOptions %>
                            </select>
                        </span>
                    </div>
                    <div style="height:25px">
                        <span class="floatLeft">Local Stock Advice:</span>
                        <span id="spsLocalStockAdviceEdit" class="floatRight">
                            <select class="inputBox" id="spsLocalStockAdvice" name="spsLocalStockAdvice">
                                <option value="0">-- Select Stock Advice --</option>
                                <%= stockAdviceOptions %>
                            </select>
                        </span>
                    </div>
                    <div style="height: 25px">
                        <span class="floatLeft">First Service Dt.:</span>
                        <span id="spsFirstServiceDtEdit" class="floatRight">
                            <input class="inputBox" style="width:150px" type="text" id="spsFirstServiceDt" name="spsFirstServiceDt" value="" />
                            <img src="../images/calendar.gif" alt="Calendar" onclick="cmdDate_onclick('spsFirstServiceDt')" />
                            <%
                            If(Len(Trim(fsDate))>0) Then
                            %>
                            &nbsp;
                            <input type="checkbox" id="chkbxUseDefFSD" name="chkbxUseDefFSD" onclick="setDefaultFSD(this,'<%=fsDate%>')"/>Use Default
                            <%
                            End If
                            %>
                        </span>
                    </div>
                    <div style="height:35px">
                        <span class="floatLeft">GEOS:</span>
                        <span class="floatRight"><b>Update</b>
                        <input type="radio" id="rdoUpdGEOSYes" name="rdoUpdGEOS" value="Y"/>Yes
                        <input type="radio" id="rdoUpdGEOSNo" name="rdoUpdGEOS" value="N" checked="true" />No
                        </span><br/><br />
                        <span id="spsGeosEdit" class="floatText">
                            <input type="checkbox" id="spsGeosNa" name="spsGeosNa" /> NA
                            <input type="checkbox" id="spsGeosLa" name="spsGeosLa" /> LA
                            <input type="checkbox" id="spsGeosApj" name="spsGeosApj" /> APJ
                            <input type="checkbox" id="spsGeosEmea" name="spsGeosEmea" /> EMEA
                        </span>
                    </div>
                    <br />
                    <div style="height: 100px">
                        <span class="floatLeft">Internal Notes:</span>
                        <span class="floatRight">
                        <input type="radio" id="rdoNotesA" name="rdoNotes" value="A" checked="true" />Append
                        <input type="radio" id="rdoNotesO" name="rdoNotes" value="O"/>Overwrite
                            <textarea class="inputBox" rows="4" id="spsNotes" name="spsNotes"></textarea>
                        </span>
                    </div>

                    <div style="height: 100px">
                        <span class="floatLeft">RSL Comments:</span><span class="floatRight">
                        <input type="radio" id="rdoCommA" name="rdoComm" value="A" checked="true" />Append
                        <input type="radio" id="rdoCommO" name="rdoComm" value="O"/>Overwrite
                            <textarea class="inputBox" rows="4" id="spsComments" name="spsComments"></textarea>
                        </span>
                    </div>
	            </fieldset>
            </div>
            <div id="SelectedRecordsInfo">
                <fieldset>
                    <legend>Selected Records (<%=CStr(intNumSels)%>)</legend>
                    <!--span style="LINE-HEIGHT: normal; FONT-VARIANT: normal; FONT-STYLE: normal; COLOR: #004874; FONT-SIZE: x-small; FONT-WEIGHT: bold"-->
                        <table class="SelectedRecsTable"  style="LINE-HEIGHT: normal; FONT-VARIANT: normal; FONT-STYLE: normal; COLOR: #004874; FONT-SIZE: x-small; FONT-WEIGHT: normal" id="SelectedRecords">
                        <tr><th>Description</th><th>Category</th></tr>
                        <%
                        If Request.Form("Action") = "" Then
                            Call DisplaySelectedRecords()
                        End If
                        %>                
                        </table>
                    <!--/span-->
                </fieldset>
            </div>
        </div>
        <input type="hidden" id="PVID" name="PVID" value="<%=strPVID%>" />
        <input type="hidden" id="SFPN" name="SFPN" value="<%=strSFPN%>" />
        <input type="hidden" id="SKIDs" name="SKIDS" value="<%=strSKIDs%>" />
        <input type="hidden" id="updateMatrix" name="updateMatrix" value="" />
        <input type="hidden" id="action" name="action" value="" />
    </form>
<%
End If
'****************************************************************************************************************************************************************************************************************************************************************************
%>
</body>
</html>