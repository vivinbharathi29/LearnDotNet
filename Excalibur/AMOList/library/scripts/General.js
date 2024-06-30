function ProcessCancelButton(RestartURL) {
	if (confirm("Are you sure you wish to cancel?"))
		window.location=RestartURL;
}

function CountCheck(ctrlName) {
	var chkBox, collection, nNumCheck=0;
	collection = document.all["cbSelected"];
	if (collection.length > 0)
		for (i=0; i<collection.length; i++) {
			chkBox = document.getElementById(collection[i].value);
			if (chkBox.checked)
			    nNumCheck += 1;
		}
	else {
		chkBox = document.getElementById(collection.value);
		if (chkBox.checked)
			nNumCheck += 1;
	}
	return nNumCheck;
}

function CheckOne(ctrlName) {
	var chkBox, collection;
	collection = document.all[ctrlName];
	if (collection.length > 0)
		for (i=0; i<collection.length; i++) {
			chkBox = document.getElementById(collection[i].value);
			if (chkBox.checked)
				return true;
		}		
	else {
		chkBox = document.getElementById(collection.value);
		if (chkBox.checked)
			return true;
	}
	return false;
}

function Cbx_AnySelected(sCtrlPrefix, nNumEntry) {
	var bSomeSelected = false;
	var oAllCol = document.all ? document.all : document.getElementsByTagName('*');
	var oObj, i;
	for (i=1; i <= nNumEntry; i++) {
		oObj = oAllCol[sCtrlPrefix + i];
		if (oObj) {
			if (oObj.checked) {
				bSomeSelected = true;
				break;
			}
		}
	}
	return bSomeSelected;
}

function Cbx_SelectAll(sCtrlPrefix, nNumEntry, bSelect) {
	var oAllCol = document.all ? document.all : document.getElementsByTagName('*');
	var oObj, i;
	for (i=1; i <= nNumEntry; i++) {
		oObj = oAllCol[sCtrlPrefix + i];
		if (oObj) {
			oObj.checked = bSelect;
		}
	}
	return true;
}

/*function ChangeReason() {
var sReason="";
  while (sReason == "") {
    sReason = window.prompt("Enter change reason for history tracking purposes and select OK to save the changes or Cancel to abort the changes:", "");
  }
  if (sReason != null)
    document.all.txtChangeReason.value = sReason;
  return sReason;
} */

function ChangeReason() {

var sReason="";

  while (sReason == "") 
  	{
   	 sReason = window.prompt("Enter change reason (max 200 chars) for history tracking purposes and select OK to save the changes or Cancel to abort the changes:", "");
 	 }
 
  sReason = ChangeReasonLength(sReason);
  
 if (sReason == null || sReason == "")
	{
		window.close();
	}
 	
 if (sReason != null && sReason.length <= 200)
	 {
		document.all.txtChangeReason.value = sReason;
	  }
  		
	return sReason;
}

function ChangeReasonLength(sReason)
{			
 	if (sReason != null && sReason.length > 200)
	{
		 var charsize = sReason.length;
		alert("The maximum number of chars for this field is 200, you entered " + charsize + " chars");
	    
 			sReason = window.prompt("Enter change reason (max 200 chars) for history tracking purposes and select OK to save the changes or Cancel to abort the changes:", sReason);
			ChangeReasonLength(sReason);
	}

	return sReason;

}

function FinishEscape(s) {
	// This handles doing an Escape on the characters that Escape can't handle
	var i;
	sNew = "";
	for (i = 0; i < s.length; i++) {   
		var c = s.charAt(i);
		if (c == "+") {
			sNew += "%2B"
		} else {
			if (c == "*") {
				sNew += "%2A"
			} else {
				if (c == "@") {
					sNew += "%40"
				} else {
					if (c == "-") {
						sNew += "%2D"
					} else {
						if (c == ".") {
							sNew += "%2E"
						} else {
							if (c == "/") {
								sNew += "%2F"
							} else {
								sNew += c
							}
						}
					}
				}
			}
		}
	}
	return sNew;
}

function strLTrim() {
	return this.replace(/^\s+/,'');
}

function strRTrim() {
	return this.replace(/\s+$/,'');
}

function strTrim() {
	return this.replace(/^\s+/,'').replace(/\s+$/,'');
}

String.prototype.ltrim = strLTrim;
String.prototype.rtrim = strRTrim;
String.prototype.trim = strTrim;

/*
Just an example to show the prototypes above work
	var strTest = '    MyString              '
	alert(strTest.ltrim()) gets 'MyString              '
	alert(strTest.rtrim()) gets '    MyString'
	alert(strTest.trim()) gets 'MyString'
*/

/* getElementsByClassName2
 * Curtis Harvey <curtis@curtisharvey.com>
 * http://curtisharvey.com/experiments/js/getelementsbyclassname/
 * inspired by the post and comments at
 * http://www.robertnyman.com/2005/11/07/the-ultimate-getelementsbyclassname/
 * 
 * USAGE:
 * var matches = getElementsByClassName(node, cls, [tag]);
 *     node (object)      : root node of search
 *     cls (array|string) : array, or space delimited string, of classnames
 *     tag (string)       : optional tagname used to limit search to specific elements
 * 
 *     returns an array of elements belonging to ALL classnames passed as cls
 */ 
function getElementsByClassName2(node, cls, tag) {
	var i, j, re, elm, elms, ismatch;
	var results = new Array();
	if (!document.getElementsByTagName) return results;
	// check type and validity of cls param
	if (typeof cls == 'string') cls = cls.split(' ');
	if (cls.length == 0) return results;
	// convert array of classnames to array of regexp objects matching classnames
	// doing this allows RegExp objects to be created just once
	for (i=0; i<cls.length; i++) {
		cls[i] = new RegExp('(^| )'+cls[i]+'( |$)');
	}
	// grab all the elements in node of type tag
	if (tag == null) tag = '*';
	elms = (tag == '*' && document.all) ? document.all : node.getElementsByTagName(tag); // IE5 does not like getElementsByTagName('*')
	// find matching elements
	for (i=0; elm=elms[i]; i++) {
		for (j=0; re=cls[j]; j++) {
			ismatch = true;
			if (!re.test(elm.className)) {
				ismatch = false;
				break;
			}
		}
		if (ismatch) results[results.length] = elm; // no Array.push in IE5
	}
	return results;
}

// From http://www.robertnyman.com/2005/11/07/the-ultimate-getelementsbyclassname/
// just a little slower than getElementsByClassName2
// oClassNames if more than one class to search for requires it to be entered as an array, i.e. [one],[two]
function getElementsByClassName(oElm, strTagName, oClassNames){
    var arrElements = (strTagName == "*" && oElm.all)? oElm.all : oElm.getElementsByTagName(strTagName);
    var arrReturnElements = new Array();
    var arrRegExpClassNames = new Array();
    if(typeof oClassNames == "object"){
        for(var i=0; i<oClassNames.length; i++){
            arrRegExpClassNames.push(new RegExp("(^|\\s)" + oClassNames[i].replace(/\-/g, "\\-") + "(\\s|$)"));
        }
    }
    else{
        arrRegExpClassNames.push(new RegExp("(^|\\s)" + oClassNames.replace(/\-/g, "\\-") + "(\\s|$)"));
    }
    var oElement;
    var bMatchesAll;
    for(var j=0; j<arrElements.length; j++){
        oElement = arrElements[j];
        bMatchesAll = true;
        for(var k=0; k<arrRegExpClassNames.length; k++){
            if(!arrRegExpClassNames[k].test(oElement.className)){
                bMatchesAll = false;
                break;                      
            }
        }
        if(bMatchesAll){
            arrReturnElements.push(oElement);
        }
    }
    return (arrReturnElements)
}

// return right n characters of the passed string
function Right(str, n) {
	if (n <= 0)
		return "";
	else if (n > String(str).length)
		return str;
	else {
		var iLen = String(str).length;
		return String(str).substring(iLen, iLen - n);
	}
}

// return left n characters of the passed string
function Left(str, n) {
	if (n <= 0)
		return "";
	else if (n > String(str).length)
		return str;
	else
	return String(str).substring(0, n);
} 