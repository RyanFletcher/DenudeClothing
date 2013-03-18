// JavaScript Document
function validateForm()
{  
var x=document.forms["SignUp"]["FirstName"].value;
if (x==null || x=="" )
  {
  alert("Please enter a valid First Name");
  return false;
  }
var x=document.forms["SignUp"]["Surname"].value;
if (x==null || x=="")
  {
  alert("Please enter a valid Surname");
  return false;
  }
    
var x=document.forms["SignUp"]["EmailAddress"].value;
var atpos=x.indexOf("@");
var dotpos=x.lastIndexOf(".");
if (atpos<1 || dotpos<atpos+2 || dotpos+2>=x.length)
  {
  alert("Please enter a valid Email Address");
  return false;
  }
  
var x=document.forms["SignUp"]["Password"].value;
if (x==null || x=="")
  {
  alert("Please enter a valid Password");
  return false;
  }
  
}

function textonly(e){
var code;
if (!e) var e = window.event;
if (e.keyCode) code = e.keyCode;
else if (e.which) code = e.which;
var character = String.fromCharCode(code);
    var AllowRegex  = /^[\ba-zA-Z\-]$/;
    if (AllowRegex.test(character)) return true;     
    return false; 
}