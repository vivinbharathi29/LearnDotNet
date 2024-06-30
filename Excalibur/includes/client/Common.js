/******************************************************************************
* Function Desc:  Validates that text was entered into input box
* Parameters:     text input object, error message 
* Returns:        boolean indicating if value was enter into text box
******************************************************************************/
function validateTextInput(inputText, errMessage) {
  if (inputText.parentElement.parentElement.style.display == "none")
	return true;

  if (inputText.value.length == 0 || trim(inputText.value) == "") {
    alert(errMessage + " required.");
    inputText.value = "";

    //if control is hidden, do not set focus
    if (inputText.name.substr(0, 3) != 'hid')
		inputText.focus();   
    return false;
  }
    
  return true;
}

/******************************************************************************
* Function Desc:  Validates that text was entered into input box on Tab control
* Parameters:     text input object, error message 
* Returns:        boolean indicating if value was enter into text box
******************************************************************************/
function validateTextInputOnTab(inputText, tabnum, errMessage) {
  if (inputText.value.length == 0 || trim(inputText.value) == "") {
    alert(errMessage + " required.");
    
    if (tabnum != 0) {
      SwitchTab( tabnum );
    }
    inputText.value = "";
    
    //if control is hidden, do not set focus
    if (inputText.name.substr(0, 3) != 'hid')
		inputText.focus();   
    
    return false;
  }
    
  return true;
}

/******************************************************************************
* Function Desc:  Validates that text entered into TextArea control
*                 did not exceed specified size.
* Parameters:     TextArea object, text length, message 
* Returns:        boolean indicating if value was enter into text box
******************************************************************************/
function validateTextAreaSize(textArea, textLength, controlName) {
  if (textArea.value.length > textLength) {
    alert(controlName + " exceeded " + textLength + " characters.");
    textArea.focus();
    
    return false;
  }    
  return true;
}

/******************************************************************************
* Function Desc:  Validates that text entered into TextArea control
*                 did not exceed specified size.
* Parameters:     TextArea object, text length, tab no, message 
* Returns:        boolean indicating if value was enter into text box
******************************************************************************/
function validateTextAreaSizeOnTab(textArea, textLength, tabNo, controlName) {
  if (textArea.value.length > textLength) {
    alert(controlName + " exceeded " + textLength + " characters.");
    SwitchTab(tabNo);
    textArea.focus();
    
    return false;
  }    
  return true;
}

/******************************************************************************
* Function Desc:  Opens a new window
* Parameters:     p_strURL = name of page that you want to open (i.e. sample.asp)
*				          p_strWinName = id of the window to open
*				          p_intHeight = height of the window that you want to open
*				          p_intWidth = width of the window that you want to open
*				          p_blnScroll = indicate "yes" to allow scrolling or "no" to disallow
* Returns:        N/A
******************************************************************************/
function openCenterWin(p_strURL, p_strWinName, p_intHeight, p_intWidth, p_blnScroll, p_blnToolbar) {
  window.open(p_strURL, p_strWinName, "height=" + p_intHeight + ",width=" + p_intWidth + 
    ",scrollbars=" + p_blnScroll + ",toolbar=" + p_blnToolbar + ",top=" + ((screen.availHeight-p_intHeight)/2) + 
    ",left=" + ((screen.availWidth-p_intWidth)/2));  
}

/******************************************************************************
* Function Desc:  Validates that only alpha numeric characters are allowed.
* Parameters:     text input object, error message 
* Returns:        boolean indicating if validation was successful
******************************************************************************/
function isValidChars(inputText, errMessage) {
  // search for invalid character (excluding blank spaces)
  if (inputText.value.search(/[^A-Za-z0-9-()/ ]/g) >= 0) {  
    alert(errMessage + " contains invalid characters.");
    inputText.focus();
    return false;
  }
 
  return true;
}

/******************************************************************************
* Function Desc:  Validates that only alpha numeric characters are allowed for control 
*                  on the Tab Control.
* Parameters:     text input object, error message 
* Returns:        boolean indicating if validation was successful
******************************************************************************/
function isValidCharsOnTab(inputText, tabNum, errMessage) {
  // search for invalid character (excluding blank spaces)
  if (inputText.value.search(/[^A-Za-z0-9-()/ ]/g) >= 0) {  
    alert(errMessage + " contains invalid characters.");
    
    if (tabNum != 0) {
      SwitchTab(tabNum);        
    }
    
    inputText.focus();
    
    return false;
  }
  return true;
}

/******************************************************************************
* Function Desc: Validates that only alpha numeric characters plus some special
*                      characters are allowed for control on the Tab Control.
* Parameters:     text input object, error message 
* Returns:        boolean indicating if validation was successful
******************************************************************************/
function isValidCharsPlusOnTab(inputText, tabNum, errMessage) {
  // search for invalid character (excluding blank spaces)
  if (inputText.value.search(/[^A-Za-z0-9. ,()/-:%"]/g) >= 0) {  
    alert(errMessage + " contains invalid characters.");
    
    if (tabNum != 0) {
        SwitchTab(tabNum);        
    }
    
    inputText.focus();
    
    return false;
  }
  return true;
}

/******************************************************************************
* Function Desc:  Validates that only numeric characters are allowed.
* Parameters:     text input object, error message 
* Returns:        boolean indicating if validation was successful
******************************************************************************/
function isNumeric(inputText, errMessage) {
  if (trim(inputText.value).length == 0) {
    return true;
  }

  // search for invalid character
  var re = /\,/g;
  if (isNaN(inputText.value.replace(re, ""))) {  
    alert(errMessage + " must be numeric.");
    inputText.focus();
    return false;
  }
  return true;
}

/******************************************************************************
* Function Desc:  Validates that only numeric characters are allowed.
* Parameters:     text input object, tab id, error message 
* Returns:        boolean indicating if validation was successful
******************************************************************************/
function isNumericOnTabs(inputText, tabNo, errMessage) {
  if (trim(inputText.value).length == 0) {
    return true;
  }

  // search for invalid character
  var re = /\,/g;
  if (isNaN(inputText.value.replace(re, ""))) {
    alert(errMessage + " must be numeric.");
    SwitchTab(tabNo);
    inputText.focus();
    return false;
  }
 
  return true;
}

/******************************************************************************
* Function Desc:  Validates that a valid latitude value is entered
* Parameters:     text input object
* Returns:        boolean indicating if validation was successful
******************************************************************************/
function validateLatitude(inputText) {
	if (!isNumeric(inputText, "Latitude")) {
		return false;
	}

	if (trim(inputText.value).length == 0) {
	  return true;
	}

  if (inputText.value.indexOf(".") > -1 && 
      inputText.value.length - inputText.value.indexOf(".") > 7) {
    alert("Please limit Latitude to a maximum of 6 decimal places.");
    return false;
  }
      
  // search for invalid character (excluding blank spaces)    
  if (isNaN(inputText.value) || parseInt(Math.abs(inputText.value)) > 90 || parseInt(inputText.value) < 0) {
    alert("Latitude must be less than 90.000000 and greater than 0.");
    return false;
  }
  return true;
}
/******************************************************************************
* Function Desc:  Validates that a valid longitude value is entered
* Parameters:     text input object
* Returns:        boolean indicating if validation was successful
******************************************************************************/
function validateLongitude(inputText) {
	if (!isNumeric(inputText, "Longitude")) {
		return false;
	}
	
	if (trim(inputText.value).length == 0) {
		return true;
	}

  if (inputText.value.indexOf(".") > -1 && 
      inputText.value.length - inputText.value.indexOf(".") > 7) {
    alert("Please limit Longitude to a maximum of 6 decimal places.");
    return false;
  }

	// search for invalid character (excluding blank spaces)  
	if ((inputText.value.search(/[0-9\-\. ]/g) < 0) || (parseInt(inputText.value) < -180) || (parseInt(inputText.value) > 0)) {  
	  alert("Longitude must be less than 0 and greater than -180.000000.");
	  return false;
	}
	return true;
}

/******************************************************************************
* Function Desc:  Validates that a valid Elevation value is entered
* Parameters:     text input object
* Returns:        boolean indicating if validation was successful
******************************************************************************/
function validateElevation(inputText) {
  if (trim(inputText.value).length == 0) {
      return true;
  }

  // search for invalid character (excluding blank spaces)
  var re = /\,/g;
  if (isNaN(inputText.value.replace(re, "")) || 
      (parseInt(Math.abs(inputText.value.replace(re, ""))) >= 30000)) {

    alert("Elevation must be less than 30,000 and greater than -30,000.");
    inputText.focus();
    return false;
  }
  return true;
}

/******************************************************************************
* Function Desc:  Validates that a valid Elevation value is entered
* Parameters:     text input object
* Returns:        boolean indicating if validation was successful
******************************************************************************/
function validateElevationOnTabs(inputText, tabNo) {
  if (trim(inputText.value).length == 0) {
    return true;
  }

  var re = /\,/g;
  if (isNaN(inputText.value.replace(re,"")) || 
      (parseInt(Math.abs(inputText.value.replace(re,""))) >= 30000)) {
    alert("Elevation must be less than 30,000 and greater than -30,000.");    
    SwitchTab(tabNo);
    inputText.focus();
    return false;
  }
  return true;
}

/******************************************************************************
* Function Desc:  Replaces @apos@ with apostrophe ('). 
* Parameters:     Input string
* Returns:        Input string with @apos@ substitution.
******************************************************************************/
function replaceApos(strIn) {
  var strRet = strIn.replace(/@apos@/g, "'");
  
  return strRet;
}

/***********************************************************************
* Function for transferring option tags
* and data between two adjacent list boxes.  This
* should be both Netscape and IE compliant.
************************************************************************/
function move(source,destination,blnSelectedOnly) {
   for (var i = 0; i < source.length; i++) {
   
     if (blnSelectedOnly) {
        if (source.options[i].selected) {
	       if (source.options[i].value != "default") {
		      destination.options[destination.length] = new Option(source.options[i].text)
		      destination.options[destination.length - 1].value = source.options[i].value
		      source.options[i] = null
		      i = i - 1
	       }
	     }
     }
     else
     {
       if (source.options[i].value != "default") {
		      destination.options[destination.length] = new Option(source.options[i].text)
		      destination.options[destination.length - 1].value = source.options[i].value
		      source.options[i] = null
		      i = i - 1
	       }
     }  // end of if selected only 
   }  // end of for-loop	
   return;
}

/******************************************************************************
* Function Desc:  Trims leading and trailing blanks from a string.
* Parameters:     String
* Returns:        String
******************************************************************************/
function trim(inputString) {
   // Removes leading and trailing spaces from the passed string. Also
   // removes consecutive spaces and replaces it with one space.
   var retValue = inputString;
   
   var ch = retValue.substring(0, 1);
   
   while (ch == " ") { // Check for spaces at the beginning of the string
      retValue = retValue.substring(1, retValue.length);
      ch = retValue.substring(0, 1);
   }
   
   ch = retValue.substring(retValue.length-1, retValue.length);
   
   while (ch == " ") { // Check for spaces at the end of the string
      retValue = retValue.substring(0, retValue.length-1);
      ch = retValue.substring(retValue.length-1, retValue.length);
   }
   
   while (retValue.indexOf("  ") != -1) { // Note that there are two spaces in the string - look for multiple spaces within the string
      retValue = retValue.substring(0, retValue.indexOf("  ")) + retValue.substring(retValue.indexOf("  ")+1, retValue.length); // Again, there are two spaces in each of the strings
   }
   
   return retValue; // Return the trimmed string back to the user
} // Ends the "trim" function

/******************************************************************************
* Function Desc:  Post form back to parent.
* Parameters:     Page name
* Returns:        n/a
******************************************************************************/
function backToParent(pageName) {
	with (document.frmMain) {
		action = pageName;
		submit();
	}
}

/******************************************************************************
* Function Desc:  Auto tab to next input field.
* Parameters:     Page name
* Returns:        n/a
******************************************************************************/
function autoTab(input,len, e) {
  var isNN = (navigator.appName.indexOf("Netscape")!=-1);
  
	var keyCode = (isNN) ? e.which : e.keyCode; 
	var filter = (isNN) ? [0,8,9] : [0,8,9,16,17,18,37,38,39,40,46];
	var i = 1;
	
	if(input.value.length >= len && !containsElement(filter,keyCode)) {
	  input.value = input.value.slice(0, len);

	  if(input.form[(getIndex(input)+i) % input.form.length].disabled == true){
		while(input.form[(getIndex(input)+i) % input.form.length].disabled == true){  
			i+=1;
		}		
	  }
	  input.form[(getIndex(input)+i) % input.form.length].focus();
	}
	
	function containsElement(arr, ele) {
		var found = false, index = 0;
		while(!found && index < arr.length)
			if(arr[index] == ele)
				found = true;
			else
				index++;
			return found;
	}
	
	function getIndex(input) {
		var index = -1, i = 0, found = false;
		while (i < input.form.length && index == -1)
			if (input.form[i] == input)index = i;
			else i++;
		return index;
	}
	return true;
}

/******************************************************************************
* Function Desc:  Auto tab to next input field.
* Parameters:     Page name
* Returns:        n/a
******************************************************************************/
function autoTab2(input,len, e, input2) {
  var isNN = (navigator.appName.indexOf("Netscape")!=-1);
  
	var keyCode = (isNN) ? e.which : e.keyCode; 
	var filter = (isNN) ? [0,8,9] : [0,8,9,16,17,18,37,38,39,40,46];
	
	if(input.value.length >= len && !containsElement(filter,keyCode)) {
	  input.value = input.value.slice(0, len);
	  input2.focus();
	}
	
	function containsElement(arr, ele) {
		var found = false, index = 0;
		while(!found && index < arr.length)
			if(arr[index] == ele)
				found = true;
			else
				index++;
			return found;
	}
	
	function getIndex(input) {
		var index = -1, i = 0, found = false;
		while (i < input.form.length && index == -1)
			if (input.form[i] == input)index = i;
			else i++;
		return index;
	}
	return true;
}


/******************************************************************************
* Function Desc:  Validates that only positive number are allowed.
* Parameters:     text input object, error message 
* Returns:        boolean indicating if validation was successful
******************************************************************************/
function isPositiveNumber(inputText,errMessage) {
	if (!isNumeric(inputText, errMessage)) {
		return false;
	}

	var dblTemp=parseFloat(inputText.value);

	if (dblTemp<=0) 
	{
		alert(errMessage + " must be greater than 0.");
		inputText.focus();
		return false;
	}
	return true;
}
/******************************************************************************
* Function Desc:  Validates that only positive number are allowed.
* Parameters:     text input object, error message 
* Returns:        boolean indicating if validation was successful
******************************************************************************/
function isPositiveNumberOnTab(inputText,tabNo,errMessage) {
	if (!isNumeric(inputText, errMessage)) {
		return false;
	}

	var dblTemp=parseFloat(inputText.value);

	if (dblTemp<=0) 
	{
		alert(errMessage + " must be greater than 0.");
		SwitchTab(tabNo);
		inputText.focus();
		return false;
	}
	return true;
}
/******************************************************************************
* Function Desc:  Validates that only non negative ( >= 0) number are allowed.
* Parameters:     text input object, error message 
* Returns:        boolean indicating if validation was successful
******************************************************************************/
function isNonNegativeNumber(inputText,errMessage) {
	if (!isNumeric(inputText, errMessage)) {
		return false;
	}

	var dblTemp=parseFloat(inputText.value);

	if (dblTemp<0) 
	{
		alert(errMessage + " must be greater than or equal to 0.");
		inputText.focus();
		return false;
	}
	return true;
}
/******************************************************************************
* Function Desc:  Format a number as a currency.
* Parameters:     text input object, error message 
* Returns:        boolean indicating if validation was successful
******************************************************************************/
function formatCurrency(num) {
  num = num.toString().replace(/\$|\,/g,'');
  
  if(isNaN(num))
    num = "0";
  
  sign = (num == (num = Math.abs(num)));
  num = Math.floor(num*100+0.50000000001);
  cents = num%100;
  num = Math.floor(num/100).toString();
  
  if(cents<10)
    cents = "0" + cents;
    
  for (var i = 0; i < Math.floor((num.length-(1+i))/3); i++)
    num = num.substring(0,num.length-(4*i+3))+','+
    num.substring(num.length-(4*i+3));
    return (((sign)?'':'-') + '$' + num + '.' + cents);
}

/******************************************************************************
* Function Desc:  Validates that text entered into TextBox control
*                 did not exceed specified size.
* Parameters:     TextBox object, text length, tab no, message 
* Returns:        boolean indicating if value was enter into text box
******************************************************************************/
function validateMaxValueOnTab(inputText, textLength, tabNo, errMessage) {

  if (inputText.value >= textLength) {

    alert(errMessage + " must be less than " + textLength + ".");
	if(tabNo != 0)
		SwitchTab(tabNo);
    inputText.focus();
    
    return false;
  }    
  return true;
}

/******************************************************************************
* Function Desc:  Validates that text was entered into input box
* Parameters:     text input object, error message 
* Returns:        boolean indicating if value was enter into text box
******************************************************************************/
function validateWBSLevel(inputText, level, validateType) {
var lowerleveltext;
var leveltext;

  if (inputText.value == level) { //If at the lowest level
	if (validateType == 'CREATE_LOWER_LEVEL'){
		if (level == 'PROJECT') {
			lowerleveltext = "Sub Projects";
			leveltext = "Project";
		}
		else {
			lowerleveltext = "Sub-Sub Projects";
			leveltext = "Sub Project";
		}
			
		alert(lowerleveltext + " not allowed. A WBS already exists for the " + leveltext + ". \n\n "
			+ "If you wish to change the hierarchy, you must first delete the WBS elements for the " 
			+ leveltext + ".");
	    return false;
	}
	else {
		if (validateType == 'CHECK_AFE_TYPE_VALUE' && document.frmMain.cboAFEType.value == '1COMBO'){
			alert("Funding Type invalid");
			return false;
		}
	}
  }
      
  return true;
}
/******************************************************************************
* Function Desc:  Validates privilege to change/delete record
*
* Parameters:     TextBox object, text length, tab no, message 
* Returns:        boolean indicating if value was enter into text box
******************************************************************************/
function validatePrivilege(inputText) {

  if (inputText.value.length == 0) {

    alert("Action unsuccessful, insufficient privilege");
    
    return false;
  }    
  return true;
}


/******************************************************************************
* Function Desc:  Formats a number to a given number of decimals
******************************************************************************/
function formatDecimal(argvalue, addzero, decimaln) {
  var numOfDecimal = (decimaln == null) ? 2 : decimaln;
  var number = 1;

  number = Math.pow(10, numOfDecimal);

  argvalue = Math.round(parseFloat(argvalue) * number) / number;
  // If you're using IE3.x, you will get error with the following line.
  // argvalue = argvalue.toString();
  // It works fine in IE4.
  argvalue = "" + argvalue;

  if (argvalue.indexOf(".") == 0)
    argvalue = "0" + argvalue;

  if (addzero == true) {
    if (argvalue.indexOf(".") == -1)
      argvalue = argvalue + ".";

    while ((argvalue.indexOf(".") + 1) > (argvalue.length - numOfDecimal))
      argvalue = argvalue + "0";
  }

  return argvalue;
}
/******************************************************************************
* Function Desc:  Validates a range of numeric values
* Parameters:     text input object, the low value, the high value, error message 
* Returns:        boolean indicating if value was inbetween the low & high values
******************************************************************************/
function validateRange(inputText, lowValue, highValue, errMessage) {
  if (inputText.parentElement.parentElement.style.display == "none")
	return true;
	
  if(inputText.value > highValue || inputText.value < lowValue){
	alert(errMessage + " must be between " + lowValue + " and " + highValue + ".");
	return false;
  }
  return true;
}

/******************************************************************************
* Function Desc:  Validates a correctly formated date was entered.
* Parameters:     text input object, error message 
* Returns:        boolean indicating if value was inbetween the low & high values
******************************************************************************/
function validateDateInput(inputText, errMessage){
  if (inputText.parentElement.parentElement.style.display == "none")
	return true;
	
  if (inputText.value == '')
    return true;

  if (isNaN(Date.parse(inputText.value))){
    alert(errMessage + " must be in mm/dd/yyyy format.");
	//inputText.value = "";

	//if control is hidden, do not set focus
	if (inputText.name.substr(0, 3) != 'hid')
	  inputText.focus();   
	return false;
  }
  return true;

}
