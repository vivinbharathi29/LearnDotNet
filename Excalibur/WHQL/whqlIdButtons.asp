<%@ Language=VBScript %>
<!-- #include file="../includes/no-cache.asp" -->
<!-- #include file = "../includes/noaccess.inc" -->

<html>
<head>
<title></title>
<script type="text/javascript" src="../includes/client/Common.js"></script>
<script type="text/javascript">
<!--

function VerifyStatus(step)
{
    var buChecked = false;
    var cpuChecked = false;

	with (window.parent.frames["UpperWindow"].frmMain)
	{
	    //Verify required fields
        switch (step) {
            case 1:
                break;
            
            case 2:
                if (!validateTextInput(submissionID, 'Submission ID')){ return false; }
		        if (!validateTextInput(submissionDt, "Submission Date")){ return false; }
		        if (!validateDateInput(submissionDt, "Submission Date")){ return false; }
                break;
            
            case 5:
		        if (!validateTextInput(BID, "Brand Selection")){ return false; }
                break;
                
            case 4:
                if (!validateTextInput(OSID, "OS Family")){ return false; }
                break;
            
            case 5:
                for (element in window.parent.frames["UpperWindow"].document.getElementsByTagName("INPUT"))
	                if (element.substr(0,5) == "cbxBu")
	                    if (window.parent.frames["UpperWindow"].document.all[element].checked)
	                        buChecked = true;
                
                if (!buChecked)
                    alert("Please select at least one Base Unit");
                    
                return buChecked
                break;
            
            case 6:
                for (element in window.parent.frames["UpperWindow"].document.getElementsByTagName("INPUT"))
	                if (element.substr(0,6) == "cbxCpu")
	                    if (window.parent.frames["UpperWindow"].document.all[element].checked)
	                        cpuChecked = true;
                
                if (!cpuChecked)
                    alert("Please select at least one CPU");
                    
                return cpuChecked
                break;
        }

	}
	return true;
}

function cmdCancel_onclick() 
{
	window.parent.close();
}

function cmdFinish_onclick() 
{
	var blnAll = true;
	var i;
	var sReturnValue;
	
	if (VerifyStatus(5))
	{
	    getCheckedBaseUnits();
		window.frmButtons.cmdFinish.disabled =true;
		window.frmButtons.cmdCancel.disabled =true;
		window.frmButtons.cmdPrevious.disabled =true;
        if (window.parent.frames["UpperWindow"].frmMain.hidFunction.value=="preview")
            window.parent.frames["UpperWindow"].frmMain.hidFunction.value="save";
        else
		    window.parent.frames["UpperWindow"].frmMain.hidFunction.value="preview";
		window.parent.frames["UpperWindow"].frmMain.submit();
	}
	
	return;
}

function document_OnLoad()
{
	window.frmButtons.cmdFinish.disabled = true;
    //changeStep(1);
            
}

function changeStep(step)
{
    if ((step <= 0) || (step > 5))
        return false;

    if (!VerifyStatus(step))
        return false;
    
    window.parent.frames["UpperWindow"].focus();
    
    hideAllSteps();
    window.parent.frames["UpperWindow"].document.all["step"+step].style.display = "";


    if (step == 1)
        window.frmButtons.cmdPrevious.disalbed = true;
    else
        window.frmButtons.cmdPrevious.disabled = false;
        
    if (step == 5){
        window.frmButtons.cmdNext.disabled = true;
        window.frmButtons.cmdFinish.disabled = false;    
        }
    else{
        window.frmButtons.cmdNext.disabled = false;
        window.frmButtons.cmdFinish.disabled = true;    
        }
    
    return true;
}

function hideAllSteps()
{
    for (element in window.parent.frames["UpperWindow"].document.getElementsByTagName("DIV"))
	    if (element.substr(0,4) == "step")
	        window.parent.frames["UpperWindow"].document.all[element].style.display = "none";
}

function cmdNext_onclick(){
    var curStep = window.parent.frames["UpperWindow"].document.all["hidCurrentStep"].value;
    curStep ++;
    if (changeStep(curStep))
        window.parent.frames["UpperWindow"].document.all["hidCurrentStep"].value = curStep;
}

function cmdPrevious_onclick(){
    var curStep = window.parent.frames["UpperWindow"].document.all["hidCurrentStep"].value;
    curStep --;
    if (changeStep(curStep))
        window.parent.frames["UpperWindow"].document.all["hidCurrentStep"].value = curStep;
}

function getCheckedBaseUnits()
{
    var uw = window.parent.frames["UpperWindow"]

    for (element in uw.document.getElementsByTagName("INPUT"))
        if (element.substr(0,3) == "cbx")
            if (uw.document.all[element].checked){
                if (element.substr(3,2) == "Bu"){
                    uw.document.all["hidBaseUnits"].value = uw.document.all["hidBaseUnits"].value + "," + element.substr(5, element.length - 5);
                }
                else if (element.substr(3,3) == "Cpu"){
                    uw.document.all["hidCpus"].value = uw.document.all["hidCpus"].value + "," + element.substr(6, element.length - 6);
                }
             }
}

//-->
</SCRIPT>
</head>
<body bgcolor="ivory" onload="document_OnLoad()">
<form id="frmButtons"  action="whqlIdButtons.asp" method="post">
<table BORDER="0" CELLSPACING="1" CELLPADDING="1" align="right">
	<tr>
		<TD><INPUT type="button" value="Cancel" id="cmdCancel" onclick="return cmdCancel_onclick()" />
		<INPUT type="button" value="<< Previous" id="cmdPrevious" onclick="return cmdPrevious_onclick()" />
		<INPUT type="button" value="Next >>" id="cmdNext" onclick="return cmdNext_onclick()" />
		<INPUT type="button" value="Finish" id="cmdFinish" onclick="return cmdFinish_onclick()" />
		</TD>
	</tr>
</table>
</FORM>
</body>
</html>