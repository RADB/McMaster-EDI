/* Added 2012-10-04 - modify later to allow reuse - true part and false part */
function chooseOnlyOne(x, language, province) {
    var strMessage
    if (province == 3) {
        if (language == 'English')
            strMessage = 'You can only respond YES to one of the following questions(7, 8a, 8b). Do you want to confirm your reponse?';
        else
            strMessage = 'Vous ne pouvez répondre OUI à l\'une des questions suivantes (7, 8a, 8b). Voulez-vous confirmer votre reponse?';

        if (document.getElementById(x).selectedIndex == 1) {
            
            var intConfirm = (confirm(strMessage))
            if (intConfirm) {
                switch (x) {
                    case "intSpecial":
                        if (document.getElementById('intLanguageDelay').selectedIndex ==1)
                            document.getElementById('intLanguageDelay').selectedIndex = 0;
                        if (document.getElementById('intDisability').selectedIndex == 1)
                            document.getElementById('intDisability').selectedIndex = 0;
                        return true;
                    case "intLanguageDelay":
                        if (document.getElementById('intDisability').selectedIndex == 1)
                            document.getElementById('intDisability').selectedIndex = 0;
                        if (document.getElementById('intSpecial').selectedIndex == 1)
                            document.getElementById('intSpecial').selectedIndex = 0;
                        return true;
                    case "intDisability":
                        if (document.getElementById('intLanguageDelay').selectedIndex == 1)
                            document.getElementById('intLanguageDelay').selectedIndex = 0;
                        if (document.getElementById('intSpecial').selectedIndex == 1)
                            document.getElementById('intSpecial').selectedIndex = 0;
                        return true;
                }
            }
            else {
                document.getElementById(x).selectedIndex = 0;
            }
        }
    }
}

function confirm_Add(strValue) 
{ 
	document.forms.Screens.Action.value = strValue;
	document.forms.Screens.submit();
}

function goAddStudent(classID)
{
	var errors ="";
	if (eval('document.forms.Class' + classID + '.localid.value.length == 0'))
	{
		errors += "";
		// commented out 2007-01-20
		//errors += ' - You must enter a local id for the student.\n';
	}	
	// if any errors are detected
	if (errors) 
	{	
		alert('The following error(s) occurred:\n\n' + errors);
	}
	else
	{
		// added 2012-01-09
		if (eval('document.forms.Class' + classID + '.localid.value.length == 0'))
		{
			if (confirm('You have not entered a localID for this child.  Do you want to continue without entering a local ID?'))
				eval('document.forms.Class' + classID + '.submit()');
		}
		else
			eval('document.forms.Class' + classID + '.submit()');
	}	
}

function admin_Check(strValue)
{
	var	iLoc = document.forms.Screens.strEmail.value.indexOf("@");
	var errors ="";

	// they must enter an email and it must be valid	
    if (( iLoc<1 ) || ( iLoc == (document.forms.Screens.strEmail.value.length-1) ))
		errors += " - A valid Email address is required.\n";

	if (document.forms.Screens.strName.value.length == 0)
		errors += ' - You must enter a name.\n';
		
	if (document.forms.Screens.strPassword.value.length == 0)
		errors += ' - You must enter a password.\n';
		
	// if any errors are detected
	if (errors) 
	{	
		alert('The following error(s) occurred:\n\n' + errors);
	}
	else
	{
		document.forms.Screens.Action.value = strValue;
		document.forms.Screens.submit();
	}	
}

function changeAdmin(strValue) 
{ 
	document.forms.Screens.Action.value = strValue;
	document.forms.Screens.submit();
}

function confirm_Student_Add(strValue,EDIYear) 
{ 
	document.forms.Children.Action.value = strValue;
	document.forms.Children.frmEDIYear.value = EDIYear;	
	document.forms.Children.submit();
}

// confirm that the user wants to delete
function confirm_Delete(strValue) 
{ 
	var intConfirm = confirm("Are you sure you want to delete this record?")
	if (intConfirm)
	{
		document.forms.Screens.Action.value = strValue;		
		document.forms.Screens.submit();
	}
	else
	{
		return false;
	}
}

// confirm that the user wants to delete
function confirm_Student_Delete(strValue,EDIYear) 
{ 
	var intConfirm = confirm("Are you sure you want to delete this record?")
	if (intConfirm)
	{
		document.forms.Children.Action.value = strValue;
		document.forms.Children.frmEDIYear.value = EDIYear;
		document.forms.Children.submit();
	}
	else
	{
		return false;
	}
}	

// go to summary and confirm on that page
function goConfirm_Lock(EDIYear, strSite, strSchool,strTeacher, strClass, strChild, strAction) 
{ 		
	document.forms.Children.target = '';
	document.forms.Children.action = 'edi_teacher_questionnairelock.asp';
	document.forms.Children.frmEDIYear.value = EDIYear;
	document.forms.Children.frmSite.value = strSite;
	document.forms.Children.frmSchool.value = strSchool;
	document.forms.Children.frmTeacher.value = strTeacher;
	document.forms.Children.frmClass.value = strClass;
	document.forms.Children.frmChild.value = strChild;
	document.forms.Children.frmAction.value = strAction;
	document.forms.Children.submit();
	
}

function goConfirm_Consent(EDIYear, strSite, strSchool,strTeacher, strClass, strChild, intConsent, strAction) 
{ 		
	document.forms.Children.target = '';
	document.forms.Children.action = '';
	document.forms.Children.frmEDIYear.value = EDIYear;
	document.forms.Children.frmSite.value = strSite;
	document.forms.Children.frmSchool.value = strSchool;
	document.forms.Children.frmTeacher.value = strTeacher;
	document.forms.Children.frmClass.value = strClass;
	document.forms.Children.frmChild.value = strChild;
	document.forms.Children.frmConsent.value = intConsent;
	document.forms.Children.frmAction.value = strAction;
	document.forms.Children.submit();
}

// confirm that the user wants to lock the child
function confirm_Lock(EDIYear, strSite, strSchool,strTeacher, strClass, strChild, strAction,language,confirmLanguage) 
{
    var intConfirm = confirm(language);
	if (intConfirm)
	{		
		document.forms.Children.target = '';
		document.forms.Children.frmEDIYear.value = EDIYear;
		document.forms.Children.frmSite.value = strSite;
		document.forms.Children.frmSchool.value = strSchool;
		document.forms.Children.frmTeacher.value = strTeacher;
		document.forms.Children.frmClass.value = strClass;
		document.forms.Children.frmChild.value = strChild;
		document.forms.Children.frmAction.value = strAction;
		document.forms.Children.submit();
	}
}

// confirm that the user wants to lock the child
function confirm_Unlock(EDIYear, strSite, strSchool,strTeacher, strClass, strChild, strAction) 
{ 
	document.forms.Children.frmEDIYear.value = EDIYear;
	document.forms.Children.frmSite.value = strSite;
	document.forms.Children.frmSchool.value = strSchool;
	document.forms.Children.frmTeacher.value = strTeacher;
	document.forms.Children.frmClass.value = strClass;
	document.forms.Children.frmChild.value = strChild;
	document.forms.Children.frmAction.value = strAction;
	document.forms.Children.submit();
}

function email_Password(strEmail)
{
	document.forms.Screens.hiddenAction.value = strEmail;
	document.forms.Screens.submit();
}


function goChild(EDIYear, strSite, strSchool,strTeacher, strClass, strChild)
{	
	// submit the form to the child screen
	document.forms.Children.frmEDIYear.value = EDIYear;
	document.forms.Children.frmSite.value = strSite;
	document.forms.Children.frmSchool.value = strSchool;
	document.forms.Children.frmTeacher.value = strTeacher;
	document.forms.Children.frmClass.value = strClass;
	document.forms.Children.frmChild.value = strChild;
	document.forms.Children.submit();
}

function goChildSection(EDIYear, strSite, strSchool, strTeacher, strClass, strChild, strSection)
{	
	// submit the form to the child screen
	//sends user to the questionnaire page
	document.forms.Children.action = 'edi_teacher_questionnaire.asp';
	document.forms.Children.frmAction.value = '';
	document.forms.Children.frmEDIYear.value = EDIYear;
	document.forms.Children.frmSection.value = strSection;
	document.forms.Children.frmSite.value = strSite;
	document.forms.Children.frmSchool.value = strSchool;
	document.forms.Children.frmTeacher.value = strTeacher;
	document.forms.Children.frmClass.value = strClass;
	document.forms.Children.frmChild.value = strChild;
	document.forms.Children.submit();
}

function goTeacherChild(strSite, strSchool, strTeacher, strClass, strChild, strSection)
{	
	// submit the form to the child screen
	document.forms.Children.frmSection.value = strSection;
	document.forms.Children.frmSite.value = strSite;
	document.forms.Children.frmSchool.value = strSchool;
	document.forms.Children.frmTeacher.value = strTeacher;
	document.forms.Children.frmClass.value = strClass;
	document.forms.Children.frmChild.value = strChild;
	document.forms.Children.submit();
}

function goSaveEDI(EDIYear, strSite, strSchool, strTeacher, strClass, strChild, strSection, strPrevious)
{	
	// submit the form to the child screen
	document.forms.Children.frmAction.value = strPrevious;
	document.forms.Children.frmSection.value = strSection;
	document.forms.Children.frmEDIYear.value = EDIYear;
	document.forms.Children.frmSite.value = strSite;
	document.forms.Children.frmSchool.value = strSchool;
	document.forms.Children.frmTeacher.value = strTeacher;
	document.forms.Children.frmClass.value = strClass;
	document.forms.Children.frmChild.value = strChild;
	document.forms.Children.submit();
}

function goSaveIdentity()
{	
	// submit the identity form
	document.forms.Identity.Action.value = 'Update';
	document.forms.Identity.submit();
}

function goSaveEDIChild(EDIYear, strSite, strSchool, strTeacher, strClass, strChild,strNextChild, strSection, strPrevious)
{	
	// submit the form to the child screen
	document.forms.Children.frmAction.value = strPrevious;
	document.forms.Children.frmSection.value = strSection;
	document.forms.Children.frmEDIYear.value = EDIYear;
	document.forms.Children.frmSite.value = strSite;
	document.forms.Children.frmSchool.value = strSchool;
	document.forms.Children.frmTeacher.value = strTeacher;
	document.forms.Children.frmClass.value = strClass;
	document.forms.Children.frmChild.value = strChild;
	document.forms.Children.frmNextChild.value = strNextChild;
	document.forms.Children.submit();
}

function goEDI(strAction,EDIYear,strSite, strSchool,strTeacher, strClass, strChild)
{		
	// submit the form to the child screen
	document.forms.Children.action = strAction;
	document.forms.Children.target = '';
	document.forms.Children.frmEDIYear.value = EDIYear;
	document.forms.Children.frmSite.value = strSite;
	document.forms.Children.frmSchool.value = strSchool;
	document.forms.Children.frmTeacher.value = strTeacher;
	document.forms.Children.frmClass.value = strClass;
	document.forms.Children.frmChild.value = strChild;
	document.forms.Children.submit();
}

function goReport(strEDIID)
{	
	// submit the form to the child screen
	document.forms.Screens.classes.value = strEDIID;
	document.forms.Screens.rpt.value = 'Generate';
	document.forms.Screens.XML.value = 'class_summary.rpx';
	goWindow('','Reports','520','280','top=0,left=125,resizable=yes,scrollbars=no');
	document.forms.Screens.submit();
}

function goEDIReport(strEDIID)
{	
	// submit the form to the child screen
	document.forms.Screens.Student.value = strEDIID;
	document.forms.Screens.rpt.value = 'Generate';
	document.forms.Screens.XML.value = 'edi_summary.rpx';
	goWindow('','Reports','520','280','top=0,left=125,resizable=yes,scrollbars=no');
	document.forms.Screens.submit();
}

function goAdminEDIReport(strEDIID)
{	
	// submit the form to the child screen
	strTemp = document.forms.Children.action
	document.forms.Children.action = 'edi_admin_reports.asp';
	document.forms.Children.target = 'Reports';
	document.forms.Children.Student.value = strEDIID;
	document.forms.Children.rpt.value = 'Generate';
	document.forms.Children.XML.value = 'edi_summary.rpx';
	goWindow('','Reports','520','280','top=0,left=125,resizable=yes,scrollbars=no');
	document.forms.Children.submit();
	document.forms.Children.action = strTemp;
	document.forms.Children.target = '';
}

function goTeacherEDIReport(strEDIID)
{	
	// submit the form to the child screen
	strTemp = document.forms.Children.action
	document.forms.Children.action = 'edi_teacher_reports.asp';
	document.forms.Children.target = 'Reports';
	document.forms.Children.Student.value = strEDIID;
	document.forms.Children.rpt.value = 'Generate';
	document.forms.Children.XML.value = 'edi_summary.rpx';
	goWindow('','Reports','520','280','top=0,left=125,resizable=yes,scrollbars=no');
	document.forms.Children.submit();
	document.forms.Children.action = strTemp;
	document.forms.Children.target = '';
}

function goTeacherReport(strEDIID,strEmail)
{	
	// submit the form to the child screen
	document.forms.Screens.classes.value = strEDIID;
	document.forms.Screens.email.value = strEmail;
	document.forms.Screens.rpt.value = 'Generate';
	document.forms.Screens.XML.value = 'class_summary.rpx';
	goWindow('','Reports','520','280','top=0,left=125,resizable=yes,scrollbars=no');
	document.forms.Screens.submit();
}

function goTeacherClassEDIReport(strEDIID,strEmail)
{	
	// submit the form to the child screen
	document.forms.Screens.Student.value = strEDIID;
	document.forms.Screens.email.value = strEmail;
	document.forms.Screens.rpt.value = 'Generate';
	document.forms.Screens.XML.value = 'edi_summary.rpx';
	goWindow('','Reports','520','280','top=0,left=125,resizable=yes,scrollbars=no');
	document.forms.Screens.submit();
}



function goTeacherClassReport(strEDIID,strEmail)
{	
	// submit the form to the child screen
	strTemp = document.forms.Children.action
	document.forms.Children.action = 'edi_teacher_reports.asp';
	document.forms.Children.target = 'Reports';
	document.forms.Children.classes.value = strEDIID;
	document.forms.Children.email.value = strEmail;
	document.forms.Children.rpt.value = 'Generate';
	document.forms.Children.XML.value = 'class_summary.rpx';
	goWindow('','Reports','520','280','top=0,left=125,resizable=yes,scrollbars=no');
	document.forms.Children.submit();
	document.forms.Children.action = strTemp;
	document.forms.Children.target = '';
}

function update_Check(strValue)
{
	var	iLoc = document.forms.Screens.email.value.indexOf("@");
	var errors ="";
	
	// they must enter a numeric code
	if ((document.forms.Screens.code.value.length == 0) || (isNaN(document.forms.Screens.code.value)))
		errors += ' - You must enter a numeric code.\n';
	else if (document.forms.Screens.code.value < 1)
		errors += ' - You must enter a numeric code greater than 0.\n';
	else if ((document.forms.Screens.code.value > 0) && (document.forms.Screens.code.value.length < 3))
		errors += ' - Please enter the code with leading zeroes.\n';
		
	// if they enter an email it must be valid	
    if( (document.forms.Screens.email.value.length > 0) && (( iLoc<1 ) || ( iLoc == (document.forms.Screens.email.value.length-1) )) )
		errors += " - A valid Email address is required.\n";

	// if any errors are detected
	if (errors) 
	{	
		alert('The following error(s) occurred:\n\n' + errors);
	}
	else
	{
		document.forms.Screens.Action.value = strValue;
		document.forms.Screens.submit();
	}	
}

function update_Class_Check(strValue,EDIYear)
{
	var errors ="";
	
	// they must enter a numeric code
	// changed 2015-01-02 was if (document.forms.Screens.code.options(document.forms.Screens.code.selectedIndex).value == -1)
	if (document.forms.Screens.code.value == -1)
		errors += ' - You must choose a class time.\n';
	// changed 2015-01-02 was else if (document.forms.Screens.language.options(document.forms.Screens.language.selectedIndex).value == -1)
	else if (document.forms.Screens.language.value == -1)
		errors += ' - You must choose the language the class is taught in.\n';
	
		
	// if any errors are detected
	if (errors) 
	{	
		alert('The following error(s) occurred:\n\n' + errors);
	}
	else
	{
		//document.forms.Screens.strLanguage.value = document.forms.Screens.language.options(document.forms.Screens.language.selectedIndex).text;
		document.forms.Screens.Action.value = strValue;
		document.forms.Screens.frmEDIYear.value = EDIYear;
		document.forms.Screens.submit();
	}
}

function update_Class_Comments(strValue)
{
	document.forms.Screens.Action.value = strValue;
	document.forms.Screens.submit();
}

function checkStatus(intIndex,strWarning)
{
	if (intIndex >1)
	{
		alert(strWarning);		
		document.getElementById("checkLink").click();
	}
}
function update_Student_Check(strValue,EDIYear)
{
	var errors ="";
	var blnMonth = "";	
	
	// they must enter a numeric code
	if ((document.forms.Children.code.value.length == 0) || (isNaN(document.forms.Children.code.value)))
		errors += ' - You must choose a numeric ID number.\n';
	else if (document.forms.Children.code.value < 1)
		errors += ' - You must enter a numeric code greater than 0.\n';
	else if ((document.forms.Children.code.value > 0) && (document.forms.Children.code.value.length < 2))
		errors += ' - Please enter the code with leading zeroes.\n';
		
	// they must enter a numeric local ID - changed 
	if ((document.forms.Children.localID.value.length == 0))
		errors += ' - You must enter a local ID.\n';
	// changed 2015-01-02 - new .value instead of options(....selectedIndex) - deprecated code
	if (document.forms.Children.sex.value == -1)
		errors += ' - You must choose the sex of the child.\n';
		
	if (document.forms.Children.DOBday.value == -1)
		errors += ' - You must choose the day of the date of birth of the child.\n';
	if (document.forms.Children.DOBmonth.value == -1)
		errors += ' - You must choose the month of the date of birth of the child.\n';
	if (document.forms.Children.DOByear.value == -1)
		errors += ' - You must choose the year of the date of birth of the child.\n';
	
	
	blnMonth = check_month(document.forms.Children.DOBmonth.value);
	
	if (blnMonth)
	{}
	else
		errors += ' - Please choose a birthdate that exists in time.\n';;
		
	if (document.forms.Children.postal.value.length == 0)
		errors += ' - You must enter a postal code.\n';
		
	// if any errors are detected
	if (errors) 
	{	
		alert('The following error(s) occurred:\n\n' + errors);
	}
	else
	{
		document.forms.Children.Action.value = strValue;
		document.forms.Children.frmEDIYear.value = EDIYear;
		document.forms.Children.submit();
	}
}


function update_TeacherCheck(strValue)
{
	var	iLoc = document.forms.Screens.email.value.indexOf("@");
	var errors ="";
	
	// they must enter a numeric code
	if ((document.forms.Screens.code.value.length == 0) || (isNaN(document.forms.Screens.code.value)))
		errors += ' - You must enter a numeric code.\n';
	else if (document.forms.Screens.code.value < 1)
		errors += ' - You must enter a numeric code greater than 0.\n';
	else if ((document.forms.Screens.code.value > 0) && (document.forms.Screens.code.value.length < 2))
		errors += ' - Please enter the code with leading zeroes.\n';
		
	// if they enter an email it must be valid	
    if( (document.forms.Screens.email.value.length > 0) && (( iLoc<1 ) || ( iLoc == (document.forms.Screens.email.value.length-1) )) )
		errors += " - A valid Email address is required.\n";
	 else if (document.forms.Screens.email.value.length == 0)
		errors += " - Email is required.\n";
	
	// if any errors are detected
	if (errors) 
	{	
		alert('The following error(s) occurred:\n\n' + errors);
	}
	else
	{
		document.forms.Screens.Action.value = strValue;
		document.forms.Screens.submit();
	}	
}

function update_TeacherParticipationCheck(strValue)
{
	var errors ="";
	
	// if any errors are detected
	if (errors) 
	{	
		alert('The following error(s) occurred:\n\n' + errors);
	}
	else
	{
		document.forms.Screens.Action.value = strValue;
		document.forms.Screens.submit();
	}	
}

function update_TeacherFeedbackCheck(strValue)
{	
	var errors ="";
	
	// if any errors are detected
	if (errors) 
	{	
		alert('The following error(s) occurred:\n\n' + errors);
	}
	else
	{
		document.forms.Screens.Action.value = strValue;
		document.forms.Screens.submit();
	}	
}

function check_month(intMonth)
{	
	//updated 2015-01-02 - deprecated options(...selectedIndex) funtion 
	switch(intMonth)
	{
		case "1": 
			if (document.forms.Children.DOBday.value < 32) 
				return true;
			else
				return false;	
		case "2":
			if (document.forms.Children.DOByear.value % 4 > 0)
			{
				if (document.forms.Children.DOBday.value < 29) 
					return true;
				else
					return false;	
				}
			else
			{
				if (document.forms.Children.DOBday.value < 30) 
					return true;
				else
					return false;	
			}	
		case "3":
			if (document.forms.Children.DOBday.value < 32) 
				return true;
			else
				return false;
		case "4":
			if (document.forms.Children.DOBday.value < 31) 
				return true;
			else
				return false;
		case "5": 
			if (document.forms.Children.DOBday.value < 32) 
				return true;
			else
				return false;
		case "6":
			if (document.forms.Children.DOBday.value < 31)
				return true;
			else
				return false;
		case "7":
			if (document.forms.Children.DOBday.value < 32) 
				return true;
			else
				return false;
		case "8":
			if (document.forms.Children.DOBday.value < 32) 
				return true;
			else
				return false;
		case "9": 
			if (document.forms.Children.DOBday.value < 31) 
				return true;
			else
				return false;
		case "10":
			if (document.forms.Children.DOBday.value < 32) 
				return true;
			else
				return false;
		case "11":
			if (document.forms.Children.DOBday.value < 31) 
				return true;
			else
				return false;
		case "12":
			if (document.forms.Children.DOBday.value < 32) 
				return true;
			else
				return false;
	}
}