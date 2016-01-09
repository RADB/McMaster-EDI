// check to make sure the user entres a username and password
function checkForm() 
{ 
	var errors ="";
	var	iLoc = document.login.email.value.indexOf("@");
			  
  // check the username
  if( ( iLoc<1 ) || ( iLoc == (document.login.email.value.length-1) ) )
  {
		errors += "- Email must contain an e-mail address.\n";
  		alert("The following error(s) occurred:\n" + errors);	
		document.login.email.focus()
		return false;
	}
			  
  // check the password
  if( document.login.password.value.length == 0 ) 
  {
		errors += "- Password is required.\n"; 
		alert("The following error(s) occurred:\n" + errors);	
		document.login.password.focus()
		return false;
	}
				
	return true;
}
			
// set the focus to the item that needs it
function checkFocus(intField)
{

	// gets the querystring from the browser
	var strLoc = window.location.search;
	
	if((intField==2) || (strLoc.indexOf("email=") > -1))
	{
		document.login.password.focus();
		document.login.check.value = 1;
	}
	else
		document.login.email.focus();				
}

function checkURL()
{
	// gets the querystring from the browser
	var strLoc = window.location.href.toLowerCase();

	if(strLoc.indexOf("https") == -1)
	{
		window.location = 'https://www.e-edi.ca/';
	}
}
