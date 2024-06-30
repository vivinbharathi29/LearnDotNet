<%@ Language=VBScript %>
	<%
	
  Response.Buffer = True
  Response.ExpiresAbsolute = Now() - 1
  Response.Expires = 0
  Response.CacheControl = "no-cache"
	  
	%>

<HTML>
<head>
    <meta name="GENERATOR" content="Microsoft Visual Studio 6.0">
    <script id="clientEventHandlersJS" language="javascript">
<!--


    function chkAll_onclick() {
        var i;

        if (typeof (frmChange.chkImage.length) == "undefined") {
            if (frmChange.chkImage.indeterminate) // && frmChange.chkAll.checked
            {
                frmChange.chkImage.indeterminate = 0;
                document.all("Lang" + frmChange.chkImage.value).innerText = document.all("Lang" + frmChange.chkImage.value).className;
                document.all("Row" + frmChange.chkImage.value).bgColor = "ivory";
            }
            frmChange.chkImage.checked = frmChange.chkAll.checked;
        }
        else {
            for (i = 0; i < frmChange.chkImage.length; i++) {
                if (frmChange.chkImage[i].indeterminate) //&& frmChange.chkAll.checked
                {
                    frmChange.chkImage[i].indeterminate = 0;
                    document.all("Lang" + frmChange.chkImage[i].value).innerText = document.all("Lang" + frmChange.chkImage[i].value).className;
                    document.all("Row" + frmChange.chkImage[i].value).bgColor = "ivory";
                }
                frmChange.chkImage[i].checked = frmChange.chkAll.checked;
                if (document.all("Base" + frmChange.chkImage[i].className).indeterminate) //&& frmChange.chkAll.checked
                    document.all("Base" + frmChange.chkImage[i].className).indeterminate = 0;
                document.all("Base" + frmChange.chkImage[i].className).checked = frmChange.chkAll.checked;
            }
        }


        //	for (i=0;i<document.all.length;i++)
        //		{
        //		if (document.all(i).className == "chkBase")
        //			{
        //			if (document.all(i).indeterminate) //&& frmChange.chkAll.checked
        //				document.all(i).indeterminate=0;
        //			document.all(i).checked = frmChange.chkAll.checked;
        //			}
        //		}

    }


    function chkAllRestore_onclick() {
        var i;

        if (typeof (chkImage.length) == "undefined") {
            if (chkImage.indeterminate) // && chkAll.checked
            {
                chkImage.indeterminate = 0;
                document.all("Lang" + chkImage.value).innerText = document.all("Lang" + chkImage.value).className;
                document.all("Row" + chkImage.value).bgColor = "ivory";
            }
            chkImage.checked = chkAllRestore.checked;
            if (document.all("Base" + chkImage.className).indeterminate) //&& chkAll.checked
                document.all("Base" + chkImage.className).indeterminate = 0;
            document.all("Base" + chkImage.className).checked = chkAllRestore.checked;
        }
        else {
            for (i = 0; i < chkImage.length; i++) {
                if (chkImage(i).indeterminate) //&& chkAll.checked
                {
                    chkImage(i).indeterminate = 0;
                    document.all("Lang" + chkImage(i).value).innerText = document.all("Lang" + chkImage(i).value).className;
                    document.all("Row" + chkImage(i).value).bgColor = "ivory";
                }
                chkImage(i).checked = chkAllRestore.checked;
                if (document.all("Base" + chkImage(i).className).indeterminate) //&& chkAll.checked
                    document.all("Base" + chkImage(i).className).indeterminate = 0;
                document.all("Base" + chkImage(i).className).checked = chkAllRestore.checked;
            }
        }

    }

    function chkBase_onclick() {
        var i;

        if (typeof (frmChange.chkImage.length) == "undefined") {
            if (frmChange.chkImage.indeterminate && window.event.srcElement.checked) {
                frmChange.chkImage.indeterminate = 0;
                document.all("Lang" + frmChange.chkImage.value).innerText = document.all("Lang" + frmChange.chkImage.value).className;
                document.all("Row" + frmChange.chkImage.value).bgColor = "ivory";
            }
            frmChange.chkImage.checked = window.event.srcElement.checked;
        }
        else {
            for (i = 0; i < frmChange.chkImage.length; i++) {
                if (frmChange.chkImage[i].className == window.event.srcElement.id) {
                    if (frmChange.chkImage[i].indeterminate && window.event.srcElement.checked) {
                        frmChange.chkImage(i).indeterminate = 0;
                        document.all("Lang" + frmChange.chkImage[i].value).innerText = document.all("Lang" + frmChange.chkImage[i].value).className;
                        document.all("Row" + frmChange.chkImage[i].value).bgColor = "ivory";
                    }
                    frmChange.chkImage[i].checked = window.event.srcElement.checked;
                }
            }
        }

    }


    function chkRestoreBase_onclick() {
        var i;

        if (typeof (chkImage.length) == "undefined") {
            if (chkImage.indeterminate && window.event.srcElement.checked) {
                chkImage.indeterminate = 0;
                document.all("Lang" + chkImage.value).innerText = document.all("Lang" + chkImage.value).className;
                document.all("Row" + chkImage.value).bgColor = "ivory";
            }
            chkImage.checked = window.event.srcElement.checked;
        }
        else {
            for (i = 0; i < chkImage.length; i++) {
                if (chkImage(i).className == window.event.srcElement.id) {
                    if (chkImage(i).indeterminate && window.event.srcElement.checked) {
                        chkImage(i).indeterminate = 0;
                        document.all("Lang" + chkImage(i).value).innerText = document.all("Lang" + chkImage(i).value).className;
                        document.all("Row" + chkImage(i).value).bgColor = "ivory";
                    }
                    chkImage(i).checked = window.event.srcElement.checked;
                }
            }
        }

    }

    function UpdateBase(chkClicked) {
        var i;
        var blnAllSame = true;
        var blnFound = false;

        for (i = 0; i < frmChange.chkImage.length; i++) {

            if (frmChange.chkImage(i).className != "")
                if (frmChange.chkImage(i).className == chkClicked.className) {
                    blnFound = true;
                    if ((frmChange.chkImage(i).checked != chkClicked.checked) || frmChange.chkImage(i).indeterminate) {
                        blnAllSame = false;
                    }
                }
        }
        if (typeof (chkImage) != "undefined" && !blnFound) {
            for (i = 0; i < chkImage.length; i++) {

                if (chkImage(i).className != "")
                    if (chkImage(i).className == chkClicked.className) {
                        blnFound = true;
                        if ((chkImage(i).checked != chkClicked.checked) || chkImage(i).indeterminate) {
                            blnAllSame = false;
                        }
                    }
            }
        }

        if (blnAllSame) {
            document.all("Base" + chkClicked.className).indeterminate = 0;
            document.all("Base" + chkClicked.className).checked = chkClicked.checked;
        }
        else
            document.all("Base" + chkClicked.className).indeterminate = -1;

        if (chkClicked.checked) {
            document.all("Lang" + chkClicked.value).innerText = document.all("Lang" + chkClicked.value).className;
            if (document.all("Row" + chkClicked.value) != null)
                document.all("Row" + chkClicked.value).bgColor = "ivory";
        }

    }


    function chkImage_onclick() {
        UpdateBase(window.event.srcElement);
    }


    /*
    function chkBase_onclick(){
        var i;
        
        
        for (i=0;i<document.all.length;i++)
            {
            
            if (document.all(i).className != "")
                if (document.all(i).className == window.event.srcElement.id)
                    {
                    document.all(i).checked = window.event.srcElement.checked;
                    }
            }
        
    
    } 
    
    
    function chkImage_onclick(){
        var i;
        var blnAllSame=true;
        
        
        for (i=0;i<document.all.length;i++)
            {
            
            if (document.all(i).className != "")
                if (document.all(i).className == window.event.srcElement.className)
                    {
                    if (document.all(i).checked != window.event.srcElement.checked)
                        {
                            blnAllSame = false;	
                        }
                    }
            }
        
        if (blnAllSame)
            {
            document.all("Base" + window.event.srcElement.className).indeterminate=0;
            document.all("Base" + window.event.srcElement.className).checked = window.event.srcElement.checked;
            }
        else
            document.all("Base" + window.event.srcElement.className).indeterminate=-1;
        
    } 
    */

    function BaseRow_onmouseover(strID) {
        if (window.event.srcElement.id == strID) {
            window.event.srcElement.style.cursor = "default";
            return;
        }

        window.event.srcElement.style.cursor = "hand";
        document.all("BaseRow" + strID).bgColor = "lightsteelblue";
    }

    function BaseRow_onmouseout(strID) {
        document.all("BaseRow" + strID).bgColor = "cornsilk";
    }

    function BaseRow_onclick(strID) {

        if (window.event.srcElement.id == strID)
            return;


        if (document.all("ImageRow" + strID).style.display == "")
            document.all("ImageRow" + strID).style.display = "none";
        else
            document.all("ImageRow" + strID).style.display = "";
    }


    function ChooseLanguages(strID, strAll) {
        var strResults;
        var i;

        strResults = window.showModalDialog("ChangeImageLanguage.asp?SelectedLangs=" + document.all("Lang" + strID).innerText + "&AllLangs=" + strAll, "", "dialogWidth:300px;dialogHeight:240px;edge: Sunken;center:Yes; maximize: Yes;help: No;resizable: Yes;status: No");
        if (typeof (strResults) != "undefined") {
            if (strResults != "") {
                if (strAll == strResults) {
                    document.all("Row" + strID).bgColor = "ivory";
                    document.all("chkImage" + strID).indeterminate = 0;
                    document.all("chkImage" + strID).checked = true;
                    UpdateBase(document.all("chkImage" + strID));
                }
                else {
                    document.all("Row" + strID).bgColor = "mistyrose";
                    document.all("chkImage" + strID).indeterminate = -1;
                    document.all(document.all("chkImage" + strID).className).indeterminate = -1;
                }
                document.all("Lang" + strID).innerText = strResults;
            }
        }
    }

    function GetRestoreLanguageChanges() {
        var i;
        var strOutput = "";

        if (typeof (chkImage) == "undefined") {
            frmChange.txtRestoreLangList.value = "";
            return;
        }

        if (typeof (chkImage.length) == "undefined") {
            if (chkImage.indeterminate)
                chkImage.checked = true;

            if ((document.all("Lang" + chkImage.value).className != document.all("Lang" + chkImage.value).innerText) && chkImage.checked)
                strOutput = strOutput + "(" + chkImage.value + "=" + document.all("Lang" + chkImage.value).innerText + ");"

        }
        else {

            for (i = 0; i < chkImage.length; i++) {
                if (chkImage(i).indeterminate)
                    chkImage(i).checked = true;

                if ((document.all("Lang" + chkImage(i).value).className != document.all("Lang" + chkImage(i).value).innerText) && chkImage(i).checked)
                    strOutput = strOutput + "(" + chkImage(i).value + "=" + document.all("Lang" + chkImage(i).value).innerText + ");"
            }
        }


        if (strOutput != "")
            frmChange.txtRestoreLangList.value = ":" + strOutput;
        else
            frmChange.txtRestoreLangList.value = "";

    }


    function GetLanguageChanges() {
        var i;
        var strOutput = "";

        if (typeof (frmChange.chkImage) == "undefined") {
            frmChange.txtLangList.value = "";
            return;
        }

        if (typeof (frmChange.chkImage.length) == "undefined") {
            if (frmChange.chkImage.indeterminate)
                frmChange.chkImage.checked = true;

            if ((document.all("Lang" + frmChange.chkImage.value).className != document.all("Lang" + frmChange.chkImage.value).innerText) && frmChange.chkImage.checked)
                strOutput = strOutput + "(" + frmChange.chkImage.value + "=" + document.all("Lang" + frmChange.chkImage.value).innerText + ");"

        }
        else {

            for (i = 0; i < frmChange.chkImage.length; i++) {
                if (frmChange.chkImage[i].indeterminate)
                    frmChange.chkImage[i].checked = true;

                if ((document.all("Lang" + frmChange.chkImage[i].value).className != document.all("Lang" + frmChange.chkImage[i].value).innerText) && frmChange.chkImage[i].checked)
                    strOutput = strOutput + "(" + frmChange.chkImage[i].value + "=" + document.all("Lang" + frmChange.chkImage[i].value).innerText + ");"
            }
        }


        if (strOutput != "")
            frmChange.txtLangList.value = ":" + strOutput;
        else
            frmChange.txtLangList.value = "";

    }


    function ChangeThis_onmouseover() {
        window.event.srcElement.style.cursor = "hand";
    }

    function ChangeDefault_onmouseover() {
        window.event.srcElement.style.cursor = "hand";

    }

    function ChangeThis_onclick() {
        frmChange.optThis.checked = true;
        frmChange.optFuture.checked = false;
    }

    function ChangeDefault_onclick() {
        frmChange.optThis.checked = false
        frmChange.optFuture.checked = true
    }




    //-->
</script>
    <link href="../style/wizard%20style.css" type="text/css" rel="stylesheet">
    <style>
        TH {
            FONT-SIZE: xx-small;
            FONT-FAMILY: Verdana;
        }

        TD {
            FONT-SIZE: xx-small;
            FONT-FAMILY: Verdana;
        }

        .ImageTable TBODY TD {
            BORDER-TOP: gray thin solid;
        }

        .imagerows TBODY TD {
            BORDER-TOP: none;
        }

        .imagerows THEAD TD {
            BORDER-TOP: none;
        }

        A:visited {
            COLOR: blue;
        }

        A:hover {
            COLOR: red;
        }
    </style>
</head>

<BODY  bgcolor=Ivory>

<%



function GetLanguages(strImages, strID)
	dim strTemp
	
	if instr(strImages,trim(strID) & "=") = 0 then
		GetLanguages = ""
	else
		strTemp = mid(strimages,instr(strImages,trim(strID) & "=")+ len(trim(strID) & "="))
		strTemp = mid(strTemp,1,instr(strTemp,")") -1) 'Strip off )...
		GetLanguages = strTemp
	end if
	
end function

	dim cn
	dim rs
	dim i
	dim blnLoadFailed
	dim strProdName
	dim strDeliverable
	dim strImageSummary
	dim strImages
	dim strAllImages
	dim LastDefinition
	dim strLanguageList
	dim strRows
	dim strBase
	dim YesCount
	dim NoCount
	dim MixedCount
	dim strSavedLanguages
	'dim strLoadedLanguages
	dim strCellColor
	dim blnImages
	dim strAllChecked
	dim TotalImageDefsChecked
	dim TotalRestoreImageDefsChecked
    dim RestoreRows
	dim strRestoreImages
    dim strRestoreImageLanguages
    dim blnPatch

	TotalImageDefsChecked = 0
	TotalRestoreImageDefsChecked = 0
	strAllChecked = "checked"
	AnyImagesChecked=false
	blnLoadFailed = false
	strProdName=""
	strDeliverable = ""
	strImageSummary = ""
	strImages = ""
    RestoreRows = ""
    blnPatch = false
	
	set cn = server.CreateObject("ADODB.connection")
	cn.ConnectionString = Session("PDPIMS_ConnectionString")
	cn.Open

	set rs = server.CreateObject("ADODB.recordset")


	if not blnLoadFailed then
		rs.Open "spGetProductVersionName " & clng(request("ProductID")),cn,adOpenForwardOnly
		if rs.EOF and rs.BOF then
			strProdName = ""
			blnLoadFailed = true
		else
			strprodName = rs("name") & ""
		end if
		
		rs.Close
	end if

	if not blnLoadFailed then
        if trim(request("VersionID")) = "" then
            'Load Root Delievrable info
		    rs.Open "spGetRootProperties4Version " & clng(request("RootID")),cn,adOpenForwardOnly
		    if rs.EOF and rs.BOF then
    			strDeliverable = ""
			    blnLoadFailed = true
		    else
    			strDeliverable = rs("name") & ""
		    end if
    		rs.Close
        else
            'Load Delievrable Version info
		    rs.Open "spGetVersionProperties4Web " & clng(request("VersionID")),cn,adOpenForwardOnly
		    if rs.EOF and rs.BOF then
			    strDeliverable = ""
			    blnLoadFailed = true
		    else
			    strDeliverable = rs("name") & "&nbsp;-&nbsp;" & rs("Version")
			    if trim(rs("Revision") & "") <> "" then
				    strDeliverable = strDeliverable & "," & rs("Revision")
			    end if
			    if trim(rs("Pass") & "") <> "" then
				    strDeliverable = strDeliverable & "," & rs("Pass")
			    end if
			    strDeliverable = strDeliverable &  "&nbsp;"
		    end if
		
    		rs.Close
	    end if
    end if
	if not blnLoadFailed then
        if trim(request("VersionID")) = "" then
       		rs.Open "spGetDistributionRoot " & clng(request("ProductID")) & ","  & clng(request("RootID")),cn,adOpenForwardOnly
        else
		    rs.Open "spGetDistributionVersion " & clng(request("ProductID")) & ","  & clng(request("VersionID")),cn,adOpenForwardOnly
		end if
        if rs.EOF and rs.BOF then
			blnLoadFailed = true
		else
			strImageSummary = rs("ImageSummary") & ""
			strImages = rs("Images") & ""
	        strRestoreImages = rs("RestoreImages") & ""
		    if trim(rs("patch")&"") <> "0" then
                blnPatch = true
            end if

			'if instr(strImages,":")>0 then
			''	strLoadedLanguages = mid(strImages,instr(strImages,":")+1)
			'else
			'	strLoadedLanguages = ""
			'end if			
		end if
		rs.Close

	end if



   blnImages = false
	if instr(strImages,":")>0 then
		if left(strImages,instr(strImages,":")-1)<>"" then
			blnImages = true
		end if
	elseif strImages <> "" then
		blnImages = true
	end if

    strRestoreImageLanguages = strRestoreImages

    if strRestoreImages <> "" then
        if instr(strRestoreImages,":")>0 then
            strRestoreImages = left(strRestoreImages,instr(strRestoreImages,":")-1)
        end if
    '    if strimages = "" then
    '        strimages = strRestoreImages
    '    else
    '        strimages = strimages & " " & strRestoreImages
    '    end if
    end if

 
	if trim(strImageSummary) = "" then
		strImageSummary = "ALL"
	end if

	if trim(strImages) <> "" then
		strImages = ", " & strImages & ","
	end if
	if trim(strRestoreImages) <> "" then
		strRestoreImages = ", " & strRestoreImages & ","
	end if
    

	strAllImages = ""
	rs.open "spListImagesForProductAll2 " & clng(request("ProductID")),cn,adOpenForwardOnly
	lastDefinition = ""
	if rs.EOF and rs.BOF then
		strAllImages = "<TR><TD colspan=10><FONT size=1 face=verdana>No images defined for this product.</font></td></tr>"
	else
		dim imagecount
        dim lastTypeID
		imagecount=0
        lastTypeID=0
		do while not rs.EOF
			if ( (rs("StatusID") = 1 and not blnPatch) or (rs("StatusID") = 2 and blnPatch) ) and isnumeric(rs("Priority")) then 'trim(rs("Priority")) = "1" or trim(rs("Priority")) = "2"  or trim(rs("Priority")) = "3"  or trim(rs("Priority")) = "4"  or trim(rs("Priority")) = "5"  or trim(rs("Priority")) = "6" then
				if rs("ImageTypeID") <> 2 then
                    imagecount = imagecount + 1
                end if
				if lastDefinition <> rs("DefinitionID") and lastDefinition <> "" then
                    if trim(lastTypeID) = "2" then
    					RestoreRows = RestoreRows & strBase
                        RestoreRows = RestoreRows &  "<TR style=""Display:none"" id=""ImageRow" & lastDefinition & """ bgcolor=cornsilk ><TD>&nbsp;</TD><TD colspan=7><BR><table cellspacing=0 cellpadding=2 border=1 bordercolor=gray class=imagerows border=0><THead bgcolor=wheat><TD>&nbsp;</TD><TD><b>Dash</b></TD><TD><b>Region</b></TD><TD><b>Code</b></TD><TD><b>Config</b></TD><TD><b>Lang</b></TD><TD><b>KBD</b></TD><TD><b>Cord</b></TD></THEAD>" & strRows & "</table><BR></TD></tr>"
					else
    					strAllImages = strAllImages & strBase
					    strAllImages = strAllImages &  "<TR style=""Display:none"" id=""ImageRow" & lastDefinition & """ bgcolor=cornsilk ><TD>&nbsp;</TD><TD colspan=7><BR><table cellspacing=0 cellpadding=2 border=1 bordercolor=gray class=imagerows border=0><THead bgcolor=wheat><TD>&nbsp;</TD><TD><b>Dash</b></TD><TD><b>Region</b></TD><TD><b>Code</b></TD><TD><b>Config</b></TD><TD><b>Lang</b></TD><TD><b>KBD</b></TD><TD><b>Cord</b></TD></THEAD>" & strRows & "</table><BR></TD></tr>"
                    end if
                    strRows = ""
					YesCount = 0
					NoCount = 0
					MixedCount=0
				end if
				lastdefinition = rs("DefinitionID")
                lastTypeID = rs("ImageTypeID")
				
				strLanguageList = rs("OSLanguage")
				if trim(rs("OtherLanguage") & "") <> "" then
					strLanguageList = strLanguageList & "," & rs("OtherLanguage")
				end if	
				
                if rs("ImageTypeID") = 2 then
                    strSavedLanguages = getLanguages(strRestoreImageLanguages,rs("ID"))
                else
				    strSavedLanguages = getLanguages(strImages,rs("ID"))
				end if
                if strSavedLanguages = "" then
					strSavedLanguages = strLanguageList
				end if

		    	if request("Type") = "1" then
                    strEnableRow = "disabled"
                else
                    strEnableRow = ""
                end if			
				strCellColor = "ivory"
				if (rs("ImageTypeID") <> 2 and instr(strImages,", " & rs("ID") & ",") > 0) or (rs("ImageTypeID") = 2 and instr(strRestoreImages,", " & rs("ID") & ",") > 0) or (rs("ImageTypeID") <> "2" and not blnImages )then
                    if strLanguageList <> strSavedLanguages then
						strCellColor = "mistyrose"
						MixedCount = MixedCount + 1
						strRows = strRows & "<TR bgcolor=ivory><TD><INPUT " & strEnableRow & " indeterminate=-1 class=""" & trim(rs("DefinitionID")) & """ type=""checkbox"" id=chkImage" & rs("ID") & " name=chkImage  LANGUAGE=javascript onclick=""return chkImage_onclick()"" value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""></TD>"
					else
						strCellColor = "ivory"
						YesCount = YesCount + 1
						strRows = strRows & "<TR bgcolor=ivory><TD><INPUT " & strEnableRow & " class=""" & trim(rs("DefinitionID")) & """ type=""checkbox"" checked id=chkImage" & rs("ID") & " name=chkImage  LANGUAGE=javascript onclick=""return chkImage_onclick()"" value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""></TD>"							
					end if
				else
					strRows = strRows & "<TR bgcolor=ivory><TD><INPUT " & strEnableRow & " class=""" & trim(rs("DefinitionID")) & """ type=""checkbox"" id=chkImage" & rs("ID") & " name=chkImage value=""" & rs("ID") & """ style=""WIDTH:16;HEIGHT:16""  LANGUAGE=javascript onclick=""return chkImage_onclick()""></TD>"
					NoCount = NoCount + 1
				end if
				if clng(request("ProductID")) = 100 and (rs("StatusID") = 3 or (rs("StatusID") = 2 and not blnPatch)) then
					RowStyle = "style=display:none"
				else
					rowStyle=""
				end if
		    	if request("Type") = "1" then
                    strEnableBase = "disabled"
                else
                    strEnableBase = ""
                end if
                if rs("ImageTypeID") = "2" then
                    strRestoreProc = "Restore"
                else
                    strRestoreProc = ""
                end if
				if YesCount = 0  and MixedCount=0 then
    			    strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><INPUT " & strEnableBase & " id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chk" & strRestoreProc & "Base_onclick()""></TD>" 
				elseif NoCount=0 and MixedCount=0  then
					strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><INPUT " & strEnableBase & " id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" checked  style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chk" & strRestoreProc & "Base_onclick()""></TD>" 
				    if rs("IMageTypeID") = 2 then
                        TotalRestoreImageDefsChecked= TotalRestoreImageDefsChecked + 1
                    else
                        TotalImageDefsChecked= TotalImageDefsChecked + 1
				    end if  
                else
					strBase = "<TR " & rowStyle & " id=BaseRow" & rs("DefinitionID") & " LANGUAGE=javascript onmouseover=""return BaseRow_onmouseover(" & rs("DefinitionID") & ")"" onmouseout=""return BaseRow_onmouseout(" & rs("DefinitionID") & ")"" onclick=""return BaseRow_onclick(" & rs("DefinitionID") & ")"" bgcolor=cornsilk><TD><INPUT " & strEnableBase & " id=""" & rs("DefinitionID") & """ name=""Base" & rs("DefinitionID") & """ type=""checkbox"" class=""chkBase"" checked style=""WIDTH:16;HEIGHT:16"" LANGUAGE=javascript onclick=""return chk" & strRestoreProc & "Base_onclick()"" indeterminate=-1></TD>"
				    if rs("IMageTypeID") = 2 then
                        TotalRestoreImageDefsChecked= TotalRestoreImageDefsChecked + 1
                    else
    				    TotalImageDefsChecked = TotalImageDefsChecked + 1
	                end if
    			end if
				strBase = strBase & "<TD>" & rs("DefinitionID") & "&nbsp;&nbsp;</TD>"
				strBase = strBase &  "<TD nowrap>" & rs("SKUNumber") & "&nbsp;&nbsp;</TD>"
				strBase = strBase &  "<TD nowrap>" & rs("brand") & "&nbsp;&nbsp;</TD>"
				strBase = strBase &  "<TD nowrap>" & rs("OS") & "&nbsp;&nbsp;</TD>"
				strBase = strBase &  "<TD nowrap>" & rs("SW") & "&nbsp;&nbsp;</TD>"
				strBase = strBase &  "<TD nowrap>" & rs("ImageType") & "&nbsp;&nbsp;</TD>"
				strBase = strBase &  "<TD>" & rs("DefinitionComments") & "&nbsp;</TD>"
				strBase = strBase &  "</tr>"
				
				strRows = strRows & "<TD>" & rs("Dash") & "</TD>"
				strRows = strRows & "<TD width=130>" & rs("Region") & "</TD>"
				strRows = strRows & "<TD width=50>" & rs("CountryCode") & "</TD>"
				strRows = strRows & "<TD>" & rs("OptionConfig") & "</TD>"
				if trim(rs("OtherLanguage") & "") <> "" then
				  if request("Type") = "1" then
				    strRows = strRows & "<TD width=70 class=""" & strLanguageList & """ id=""Lang" & rs("ID") & """>" & strLanguageList & "</TD>"
				  else
					strRows = strRows & "<TD ID=""Row" & rs("ID") & """ bgcolor=" & strCellColor & " width=70><a hidefocus class=""" & strLanguageList & """ href=""javascript:ChooseLanguages(" & rs("ID") & ",'" & strlanguageList & "');"" id=""Lang" & rs("ID") & """>" & strSavedLanguages & "</a></TD>"
				  end if
				else
					strRows = strRows & "<TD width=70 class=""" & strLanguageList & """ id=""Lang" & rs("ID") & """>" & strLanguageList & "</TD>"
				end if				
				'strRows = strRows & "<TD width=70>" & strLanguageList & "</TD>"
				strRows = strRows & "<TD width=65>" & rs("Keyboard") & "</TD>"
				strRows = strRows & "<TD width=75>" & rs("Powercord") & "</TD>"
				strRows = strRows & "</tr>"
	
			end if
			rs.MoveNext
		loop
		if imagecount = 0 then
            if blnPatch then
			    strAllImages = "<TR><TD colspan=10><FONT size=1 face=verdana>No images are currently in the factory for this product.</font></td></tr>"
            else
			    strAllImages = "<TR><TD colspan=10><FONT size=1 face=verdana>No active images defined for this product.</font></td></tr>"
		    end if
        end if
        
        if trim(lastTypeID) = "2" then
    		RestoreRows = RestoreRows & strBase
            RestoreRows = RestoreRows &  "<TR style=""Display:none"" id=""ImageRow" & lastDefinition & """ bgcolor=cornsilk ><TD>&nbsp;</TD><TD colspan=7><BR><table cellspacing=0 cellpadding=2 border=1 bordercolor=gray class=imagerows border=0><THead bgcolor=wheat><TD>&nbsp;</TD><TD><b>Dash</b></TD><TD><b>Region</b></TD><TD><b>Code</b></TD><TD><b>Config</b></TD><TD><b>Lang</b></TD><TD><b>KBD</b></TD><TD><b>Cord</b></TD></THEAD>" & strRows & "</table><BR></TD></tr>"
        else
		    strAllImages = strAllImages & strBase
		    strAllimages = strAllImages & "<TR style=""Display:none"" id=""ImageRow" & lastDefinition & """ bgcolor=cornsilk ><TD>&nbsp;</TD><TD colspan=7><BR><table cellspacing=0 cellpadding=2 border=1 bordercolor=gray class=imagerows border=0><THead bgcolor=wheat><TD>&nbsp;</TD><TD><b>Dash</b></TD><TD><b>Region</b></TD><TD><b>Code</b></TD><TD><b>Config</b></TD><TD><b>Lang</b></TD><TD><b>KBD</b></TD><TD><b>Cord</b></TD></THEAD>" & strRows & "</table><BR></TD></tr>"
        end if
		strRows = ""
		

	end if	
	rs.Close

	if TotalImageDefsChecked > 0 then
		strAllChecked="checked"
	else
		strAllChecked=""
	end if

	if TotalRestoreImageDefsChecked > 0 then
		strAllRestoreChecked="checked"
	else
		strAllRestoreChecked=""
	end if

%>


<form ID=frmChange action="ChangeImagesSave.asp?pulsarplusDivId=<%=Request("pulsarplusDivId")%>" method=post>

<font size=2 face=verdana><b>Select Supported Images</b></font><br><br>
<font size=2 face=verdana><%=strDeliverable & " (" & strProdName & ")"%></font><BR><BR>

<INPUT type="hidden" style="width:100%" id=txtLangList name=txtLangList value=""> 
<b><font size=2 face=verdana>General:</font></b>
<table WIDTH="100%" BGCOLOR="cornsilk" BORDER="1" CELLSPACING="0" CELLPADDING="2" bordercolor="tan">
	<TR>
		<TD valign=top width=120><font size=2 face=verdana><b>Image&nbsp;Strategy:&nbsp;<font color=red size=1 face=verdana>*</font></b></font></TD>
		<% if request("Type") = "1" then %>
			<TD valign=top width=500><INPUT disabled type="text" style="width=100%" id=txtSummary name=txtSummary maxlength=80 value="<%=strImageSummary%>"><BR>
		<% else %>
			<TD valign=top width=500><INPUT type="text" style="width=100%" id=txtSummary name=txtSummary maxlength=80 value="<%=strImageSummary%>"><BR>
		<% end if %>
		<font size=1 color=green face=verdana>Examples: ALL, ProBook Only, Win7 Pro Only, ALL (except -03,-04,-05), etc</font>
		
		</TD>
	</tr>
	
	<% if request("Type") <> "1" and trim(request("VersionID")) <> "" then %>
	<TR>
		<TD valign=top><font size=2 face=verdana><b>Scope:</b></font></TD>
		<TD valign=top>
		<INPUT type="radio" id=optThis name=optScope value="1">&nbsp;<font size=2 face=verdana ID="ChangeThis" LANGUAGE=javascript onclick="return ChangeThis_onclick()" onmouseover="return ChangeThis_onmouseover()">Change this version only</font><BR>
		<INPUT type="radio" id=optFuture name=optScope value="2" checked>&nbsp;<font size=2 face=verdana Id="ChangeDefault" LANGUAGE=javascript onclick="return ChangeDefault_onclick()" onmouseover="return ChangeDefault_onmouseover()">Change this version and all future versions</font><BR>
		</TD>
	</TR>

	<% end if %>

</table><BR>
<!--<font size=2 face=verdana color=green>Note: Inactive and Released images are not displayed on this page.</font><BR><BR>-->
<b><font size=2 face=verdana>Platform Images:</font><BR></b>

		<table class="ImageTable" width=100% border=0 cellspacing=0 cellpadding=1 >
			<THEAD bgcolor=Wheat><tr>
				<% if request("Type") = "1" then%>
				<TH align=left><INPUT disabled type="checkbox" id=chkAll name=chkAll style="WIDTH:16;HEIGHT:16" <%=strAllChecked%> Language=javascript onclick="return chkAll_onclick()"></TH>
				<%else%>
				<TH align=left><INPUT type="checkbox" id=chkAll name=chkAll style="WIDTH:16;HEIGHT:16" <%=strAllChecked%> Language=javascript onclick="return chkAll_onclick()"></TH>
                <%end if%>
				<TH align=left>ID</TH>
				<TH align=left>SKU</TH>
				<TH align=left>Brand</TH>
				<TH align=left>OS</TH>
				<TH align=left>Apps&nbsp;Bundle</TH>
				<TH align=left>BTO/CTO</TH>
				<TH align=left>Comments</TH></tr>
			</thead>
			<%=strAllImages%>

		</table>
<INPUT style="Display:none" type="checkbox" id=chkAllChecked name=chkAllChecked>
<INPUT type="hidden" id=txtProductID name=txtProductID value="<%=request("ProductID")%>">
<INPUT type="hidden" id=txtRootID name=txtRootID value="<%=request("RootID")%>">
<INPUT type="hidden" id=txtVersionID name=txtVersionID value="<%=request("VersionID")%>">
<INPUT type="hidden" id=chkRestoreImage name=chkRestoreImage value="">
<INPUT type="hidden" id=txtRestoreLangList name=txtRestoreLangList value="">
<%if blnPatch then
    response.write "<INPUT type=""hidden"" id=txtPatch name=txtPatch value=""1"">"
else
    response.write "<INPUT type=""hidden"" id=txtPatch name=txtPatch value=""0"">"
end if%>
</form>

        <%if RestoreRows<> "" then %>
        <b><font size=2 face=verdana>Restore Images:</font><BR></b>
		<table class="ImageTable" width=100% border=0 cellspacing=0 cellpadding=1 >
			<THEAD bgcolor=Wheat><tr>
				<% if request("Type") = "1" then%>
				<TH align=left><INPUT disabled type="checkbox" id=chkAllRestore name=chkAllRestore style="WIDTH:16;HEIGHT:16" <%=strAllRestoreChecked%> Language=javascript onclick="return chkAllRestore_onclick()"></TH>
				<%else%>
				<TH align=left><INPUT type="checkbox" id=chkAllRestore name=chkAllRestore style="WIDTH:16;HEIGHT:16" <%=strAllRestoreChecked%> Language=javascript onclick="return chkAllRestore_onclick()"></TH>
                <%end if%>
				<TH align=left>ID</TH>
				<TH align=left>SKU</TH>
				<TH align=left>Brand</TH>
				<TH align=left>OS</TH>
				<TH align=left>Apps&nbsp;Bundle</TH>
				<TH align=left>BTO/CTO</TH>
				<TH align=left>Comments</TH></tr>
			</thead>
			<%=RestoreRows%>

		</table>
        <%end if%>

<%
	set rs= nothing
	set cn=nothing
%>


</BODY>
</HTML>
