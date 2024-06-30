//
//  msrsClient.js
//  A client for Remote Scripting supporting the MSRS (Microsoft) Remote Scripting
//  protocol.  This code is based on JSRS (by Brent Ashley - jsrs@megahuge.com)
//  and has been adapted to the MSRS protocol with his permission.
//
//  SYNOPSIS:  Make MSRS protocol asynchronous and synchronous Remote Scripting calls 
//             to the server side without an applet.  Can be used in .NET with 
//             Thycotic.Web.RemoteScripting
//             http://www.thycotic.com/dotnet_remotescripting.html
//
//  You can get the latest version of this library here:
//  http://www.thycotic.com/dotnet_remotescripting_client.html
//  Please report bugs and enhancement requests here too.
//
//  Author:  
//    Jonathan Cogley
//    Thycotic Software Ltd
//    http://www.thycotic.com
//
//  See msrsclient_license.txt for copyright and license information - this is carried 
//  forward from Brent Ashley's original license of JSRS.
// 
//  Changes:
//  0.43    04/27/2004 (Patch submitted by Alexander Garcia Sanabria) (14 test cases total_
//          Added support for RSGetASOject so that it works also for Asynchronous
//			calls. RSExecuteFromObject was added.
//  0.42    03/31/2004 (Patch submitted by Alexander Garcia Sanabria) (13 test cases total)
//          Added fix context full pool errror. Removed the use of contexts for 
//			synchronous calls since they are not needed
//  0.41    02/01/2004  (12 test cases total)
//          Added fix for container displaying in Mozilla.
//  0.40    12/30/2003  (12 test cases total)
//          Added support for RSGetASPObject to allow the use of a proxy
//          object when calling methods.  
//			See tests/testfixture_rsgetaspobject.html for examples.
//          Tested in Internet Explorer 6, Mozilla 1.5 and Netscape 7.1.
//  0.30    12/29/2003
//          Added support for synchronous calls for browsers that support the
//          use of an XML HTTP type object.  Synchronous calls only currently
//          support HTTP GET.
//          Added comprehensive set of tests using jsUnit (www.jsunit.net)
//          which is a spectacular product!
//          Tested in Internet Explorer 6, Mozilla 1.5 and Netscape 7.1.
//  0.23    12/23/2003
//          Fixed typo bug that prevented errorcallback function from being called.
//  0.22    12/18/2003
//          Fixed "_mtype" bug causing HTTP GET to fail in classic ASP.
//  0.21    02/25/2003
//          Fixed bug with window opening when using POST.
//  0.20    02/24/2003
//          Added support for HTTP POST.
//  0.10    02/10/2003
//          Initial release - only supports GET and asynchronous calls.
//

var msrsVersion = 0.43;

// callback pool needs global scope
var msrsContextPoolSize = 0;
var msrsContextMaxPool = 10;
var msrsContextPool = new Array();
var msrsBrowser = msrsBrowserSniff();
var msrsPOST = false;
var msrsVisibility = false;

// constructor for context object
function msrsContextObj( contextID ){
  // properties
  this.id = contextID;
  this.busy = true;
  this.callback = null;
  this.errorcallback = null;
  this.context = null;
  this.container = contextCreateContainer( contextID );
  this.request = null;  
  // methods
  this.GET = contextGET;
  this.POST = contextPOST;
  this.setVisibility = contextSetVisibility;
}

//  method functions are not privately scoped 
//  because Netscape's debugger chokes on private functions
function contextCreateContainer( containerName ){
  // creates hidden container to receive server data 
  var container;
  switch( msrsBrowser ) {
    case 'NS':
      container = new Layer(100);
      container.contextID = containerName;
      container.name = containerName;
      container.visibility = 'hidden';
      container.clip.width = 100;
      container.clip.height = 100;
      break;
    
    case 'IE':
      document.body.insertAdjacentHTML( "afterBegin", '<span id="SPAN' + containerName + '"></span>' );
      var span = document.all( "SPAN" + containerName );
      var html = '<iframe contextID="' + containerName + '" onload="contextLoaded(this,this.contentWindow.document.documentElement.innerHTML)" name="' + containerName + '" src=""></iframe>';
      span.innerHTML = html;
      span.style.display = 'none';
      container = window.frames[ containerName ];
      break;
      
    case 'MOZ':  
      var span = document.createElement('SPAN');
      span.id = "SPAN" + containerName;
      document.body.appendChild( span );
      var iframe = document.createElement('IFRAME');
      iframe.contextID = containerName;
      iframe.onload = function() { contextLoaded(this,this.contentWindow.document.documentElement.innerHTML); };
      iframe.name = containerName;
      span.appendChild( iframe );
      container = iframe;
      // fix for showing container in Mozilla     
      document.getElementById("SPAN" + containerName).style.visibility = 'hidden';
      container.width = 0;
      container.height = 0;
      // end fix
      break;
  }
  return container;
}

function contextPOST( rsPage, func, parms ){
  var d = new Date();
  var unique = d.getTime() + '' + Math.floor(1000 * Math.random());
  
  var doc = (msrsBrowser == "IE" ) ? this.container.document : this.container.contentDocument;
  this.container.inrequest = true;
  doc.open();
  doc.write('<html><body>');
  doc.write('<form name="msrsForm" method="post" target="" ');
  doc.write(' action="' + rsPage + '?U=' + unique + '">');
  doc.write('<input type="hidden" name="C" value="' + this.id + '">');

  // write the method to call and parameters as hidden form inputs
  if (func != null){
    doc.write('<input type="hidden" name="_method" value="' + func + '">');
    if (parms != null || parms.length == 0){
		// assume parms is array of strings
		for( var i=0; i < parms.length; i++ ){
			doc.write( '<input type="hidden" name="p' + i + '" '
                   + 'value="' + msrsEscapeQQ(parms[i]) + '">');
		}
		doc.write( '<input type="hidden" name="pcount" '
                   + 'value="' + parms.length + '">');
    } else {
		doc.write( '<input type="hidden" name="pcount" '
                   + 'value="0">');
    } // parms
  } // func

  doc.write('</form></body></html>');
  doc.close();
  doc.forms['msrsForm'].submit();
}

function getUrlforGET( rsPage, func, parms ){
  // build URL to call
  var URL = rsPage;
  // always send context
  URL += "?C=" + this.id;
  URL += "&_mtype=execute"; // for ASP compatibility only
  if (func != null){
	URL += "&_method=" + escape(func);
	if (parms != null || parms.length == 0){
		// assume parms is array of strings
		for( var i=0; i < parms.length; i++ ){
			URL += "&p" + i + "=" + escape(parms[i]+'') + "";
		}
		URL += "&pcount=" + parms.length;
    } else {
		URL += "&pcount=0";
    } // parms
  } // func  
  // unique string to defeat cache
  var d = new Date();
  URL += "&U=" + d.getTime();
  return URL;
}

function contextGET( rsPage, func, parms ){
  var URL = getUrlforGET(rsPage,func,parms);
  // make the call
  switch( msrsBrowser ) {
    case 'NS':
      this.container.src = URL;
      break;
    case 'IE':
      this.container.document.location.replace(URL);
      break;
    case 'MOZ':
      this.container.src = '';
      this.container.src = URL; 
      break;
  }  
}

function contextSetVisibility( vis ){
  switch( msrsBrowser ) {
    case 'NS':
      this.container.visibility = (vis)? 'show' : 'hidden';
      break;
    case 'IE':
      document.all("SPAN" + this.id ).style.display = (vis)? '' : 'none';
      break;
    case 'MOZ':
      document.getElementById("SPAN" + this.id).style.visibility = (vis)? '' : 'hidden';
      this.container.width = (vis)? 250 : 0;
      this.container.height = (vis)? 100 : 0;
      break;
  }  
}

// end of context constructor

function msrsGetContextID(){
  var contextObj;
  for (var i = 1; i <= msrsContextPoolSize; i++){
    contextObj = msrsContextPool[ 'msrs' + i ];
    if ( !contextObj.busy ){
      contextObj.busy = true;      
      return contextObj.id;
    }
  }
  // if we got here, there are no existing free contexts
  if ( msrsContextPoolSize <= msrsContextMaxPool ){
    // create new context
    var contextID = "msrs" + (msrsContextPoolSize + 1);
    msrsContextPool[ contextID ] = new msrsContextObj( contextID );
    msrsContextPoolSize++;
    return contextID;
  } else {
    alert( "msrs Error:  context pool full" );
    return null;
  }
}

function _RSExecuteSynchronously(rspage,method,parameters,context) {
	var url = getUrlforGET(rspage,method,parameters);
	var data = msrsDoSynchronousGet(url);
	var request = new RSCallObject();
	request.data = data;
	request.context = context;
	evalRequest(request);
	return request;
}

function RSExecute( rspage, method ) {
  // get parameters
  var parameters = new Array();	
  var callback = null;
  var errorcallback = null;
  var context = null;
  var synchronous = true;
  var finishedParameters = false;
  var length = RSExecute.arguments.length;
  var arg;
  for (var n=2; n < length; n++) {
    arg = RSExecute.arguments[n];
	if (typeof(arg) == 'function') {
	  synchronous = false;
	  finishedParameters = true;
	  if (callback == null) {
	    callback = arg;
	  } else {
	    errorcallback = arg; 
	    break;
	  }
	} else if (!finishedParameters) {
		parameters[parameters.length] = arg;
	} else {
		context = arg;
	}
  }
    
  if (synchronous) {
	if (!msrsSupportsSynchronousGets()) {
		alert("Your browser does not support synchronous calls.\nEither use a different browser such as Internet Explorer, Mozilla or Netscape OR use asynchronous calls.");
		return;
	}
	return _RSExecuteSynchronously(rspage,method,parameters,context);
  }
  // get pooling context only if this is an async call.
  var contextObj = msrsContextPool[ msrsGetContextID() ];
  // assign callbacks and context
  contextObj.callback = callback;
  contextObj.errorcallback = errorcallback;
  contextObj.context = context;

  // set visible if set
  contextObj.setVisibility( msrsVisibility );

  if (msrsPOST && ((msrsBrowser == 'IE') || (msrsBrowser == 'MOZ'))){
    contextObj.POST( rspage, method, parameters );
  } else {
    contextObj.GET( rspage, method, parameters );
  }  
}

function RSExecuteTest( rspage, method, args ) {
  // get parameters
  var parameters = new Array();	
  var callback = null;
  var errorcallback = null;
  var context = null;
  var finishedParameters = false;
  var synchronous = true;
	for (var i=0; i < args.length; i++)
	{
		if (typeof(args[i]) == 'function')
		{
			synchronous = false;
			finishedParameters = true;	// no more params
			if (callback == null)
				callback = args[i];
			else
				errorcallback = args[i];
		}
		else if (!finishedParameters)
		{
			parameters[parameters.length] = args[i];
		}
		else
			context = args[i];
	}
	      
  if (synchronous) {
	if (!msrsSupportsSynchronousGets()) {
		alert("Your browser does not support synchronous calls.\nEither use a different browser such as Internet Explorer, Mozilla or Netscape OR use asynchronous calls.");
		return;
	}
	return _RSExecuteSynchronously(rspage,method,parameters,context);
  }
  // get pooling context only if this is an async call.
  var contextObj = msrsContextPool[ msrsGetContextID() ];
  // assign callbacks and context
  contextObj.callback = callback;
  contextObj.errorcallback = errorcallback;
  contextObj.context = context;

  // set visible if set
  contextObj.setVisibility( msrsVisibility );

  if (msrsPOST && ((msrsBrowser == 'IE') || (msrsBrowser == 'MOZ'))){
    contextObj.POST( rspage, method, parameters );
  } else {
    contextObj.GET( rspage, method, parameters );
  }  
}

function msrsGetSynchronousHttpObject() {
  var remote = null;
  try {
    remote = new XMLHttpRequest();
  } catch (e) {
	try {
      remote = new ActiveXObject("Msxml2.XMLHTTP");
    } catch (e) {
      remote = new ActiveXObject("Microsoft.XMLHTTP");
    }
  }
  return remote;
}

function msrsSupportsSynchronousGets() {
  var remote = msrsGetSynchronousHttpObject();
  return remote != null;
}

function msrsDoSynchronousGet(url) {
  var remote = msrsGetSynchronousHttpObject();
  remote.open("GET",url,false);
  remote.send(null);
  return remote.responseText;
}

function msrsEscapeQQ( thing ){
  return thing.replace(/'"'/g, '\\"');
}

function msrsBrowserSniff(){
  if (document.layers) return "NS";
  if (document.all) return "IE";
  if (document.getElementById) return "MOZ";
  return "OTHER";
}

/////////////////////////////////////////////////
//
// user functions

function msrsDebugInfo(){
  // use for debugging by attaching to f1 (works with IE)
  // with onHelp = "return msrsDebugInfo();" in the body tag
  var doc = window.open().document;
  doc.open;
  doc.write( 'Pool Size: ' + msrsContextPoolSize + '<br><font face="arial" size="2"><b>' );
  for( var i in msrsContextPool ){
    var contextObj = msrsContextPool[i];
    doc.write( '<hr>' + contextObj.id + ' : ' + (contextObj.busy ? 'busy' : 'available') + '<br>');
    doc.write( contextObj.container.document.location.pathname + '<br>');
    doc.write( contextObj.container.document.location.search + '<br>');
    doc.write( '<table border="1"><tr><td>' + contextObj.container.document.body.innerHTML + '</td></tr></table>' );
  }
  doc.write('</table>');
  doc.close();
  return false;
}

//*****************************************************************
// Handle parsing the data when loading and creating the msrs return object
//*****************************************************************
function contextLoaded(container,data) {
	// get context object and invoke callback
	var contextObj = msrsContextPool[ container.contextID ];
	var request = new RSCallObject();
	request.data = data;
	request.context = contextObj.context;
	evalRequest(request);
	contextObj.request = request;
	if (request.status != MSRS_INVALID) {
		if (request.status == MSRS_FAIL)
		{	
			if (typeof(contextObj.errorcallback) == 'function')
			{
				contextObj.errorcallback(request);
			}
			else 
			{
				alert('Remote Scripting Error\n' + request.message);
			}
		}
		else
		{
			if (typeof(contextObj.callback) == 'function')
			{
				contextObj.callback(request);
			}
		}	
		// clean up and return context to pool
		contextObj.callback = null;
		contextObj.busy = false;
	}
}


//*****************************************************************
// Constants from rs.htm for MSRS
//*****************************************************************
var MSRS_FAIL = -1;
var MSRS_COMPLETED = 0;
var MSRS_PENDING = 1;
var MSRS_PARTIAL = 2;
var MSRS_INVALID = 3;
	
//*****************************************************************
// function RSGetASPObject(url)
//	This function returns a server object for an ASP file
//	described by its public_description.
//*****************************************************************
function RSGetASPObject(url)
{
	var cb, ecb, context;
	var params = new Array;
	var request = RSExecute(url,'GetServerProxy');
	if (request.status == MSRS_COMPLETED)
	{
		var server = request.return_value;
		if (typeof(Function) == 'function')
		{
			for (var name in server)
				server[name] = Function('return RSExecuteTest(this.location,"' +  name + '",this.' + name + '.arguments);');
		}
		else
		{	// JavaScript 1.0 does not support Function  ( IE3.0 )
			for (var name in server)
				server[name] = eval('function t(p0,p1,p2,p3,p4,p5,p6,p7,p8,p9,pA,pB,pC,pD,pE,pF) { return _RSExecuteSynchronously(this.location,"' + name + '",this.' + name + '.arguments);} t');
		}
		server.location = url;
		return server;
	}
	alert('Failed to create ASP object for : ' + url);
	return null;
}

//*****************************************************************
// function evalRequest(request)
//
//	This function evaluates the data returned to the request. 
//	Marshalled jscript objects are re-evaluated on the client.
//*****************************************************************
function evalRequest(request)
{
	request.status = MSRS_COMPLETED;
	var data = request.data;
	var start_index = 0;
	var end_index = 0;
	var start_key = '<' + 'return_value';
	var end_key = '<' + '/return_value>';
	// check if there otherwise switch case 
	if (data.indexOf(start_key) == -1) {
		start_key = start_key.toUpperCase();
		end_key = end_key.toUpperCase();
	}
	if ((start_index = data.indexOf(start_key)) != -1)
	{
		var data_start_index = data.indexOf('>',start_index) + 1;
		end_index = data.indexOf(end_key,data_start_index);
		if (end_index == -1) 
			end_index = data.length;
		var metatag = data.substring(start_index,data_start_index);
		if (metatag.indexOf('SIMPLE') != -1)
		{
			request.return_value = unescape(data.substring(data_start_index,end_index));
		}
		else if (metatag.indexOf('EVAL_OBJECT') != -1)
		{
			request.return_value = data.substring(data_start_index,end_index);
			request.return_value = eval(unescape(request.return_value));
		}
		else if (metatag.indexOf('ERROR') != -1)
		{
			request.status = MSRS_FAIL;
			request.message = unescape(data.substring(data_start_index,end_index));		
		}
	}
	else
	{
		request.status = MSRS_INVALID;
		request.message = 'REMOTE SCRIPTING ERROR: Page invoked does not support remote scripting.';			
		// extra debug for errors
		var win = window.open('','_blank','height=600,width=450,scrollbars=yes');
		win.document.write( request.data );
	}
}

function RSCallObject()
{
	this.status = MSRS_PENDING;
	this.message = '';
	this.data = '';
	this.return_value = '';
	this.context = null;
}