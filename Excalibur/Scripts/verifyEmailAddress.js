//*****************************************************************
//Description:  Veriy email address and multiple addresses. seperation as ";".
//
//Requirement:  jQuery required.
//Sample Code:   if (VerifyEmail(stringEmails) ){return true;} else{return false;};
//Created:      Herb, 11/16/2015, PBI 29472
//*****************************************************************

function VerifyEmail(src) {
    // if empty string then true
    var emailReg = "^\x20*([a-zA-Z0-9][a-zA-Z0-9_\+\.-]*[a-zA-Z0-9-]*)@([a-zA-Z0-9][a-zA-Z0-9\.-]*[a-zA-Z0-9-])\.([a-zA-Z]{2,})\x20*$";
    // the same conditions as the regular expression in Pulsar MVC MessageQueuedEmail.
    var regex = new RegExp(emailReg);
    var result = true;
    src = src.split(" ").join("").replace(/(\r\n|\n|\r)/gm, "");
    var maillist = src.split(";");
    for (var i = 0; i < maillist.length; i++) {
        regex.lastIndex = 0;
        if (maillist[i] != "") {
            if ((!regex.test(maillist[i])) || ((maillist[i].split("@").length-1) > 1 ) ) {
                result = false;
            }
        }

    }

    return result;
}
