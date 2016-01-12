<!-- #include virtual="/shared/dbase.asp" -->
<!-- #include virtual="/shared/browser_check.asp" -->

<%
if NOT session("user") then
%>
<html>
	<head>
		<title>www.e-EDI.ca Security</title>
		<script language="javascript" type="text/javascript" src="js/login.js"></script>
		<link rel="stylesheet" type="text/css" href="Styles/edi.css">
	</head><BODY >
	<%	
	Response.Write "<body "
	if intVersion = 0 then 
		Response.Write "onload=""javascript:checkFocus(login.check.value);"""
	end if
	Response.Write ">"

	strTopBar = "<table width=""760"" border=""1"" cellpadding=""1"" cellspacing=""0"" style=""border-color:#006600;background-color:#dcdcdc"">"
	' bgcolor=""Gainsboro"" = dcdcdc 
	strTopBar = strTopBar & "<tr><td><table width=""750"" border=""0"" cellpadding=""3"" cellspacing=""0""><tr><td align=""left""><font class=""boldTextBlack"">&nbsp;&nbsp;"
	strTopBar = strTopBar & "Session expired/Not logged on"
	strTopBar = strTopBar & "</td><td align=""right""><font class=""boldTextBlack"">" & formatdatetime(now,vblongdate) & " - " & time() & "&nbsp;&nbsp;</font></td></tr></table>"	
	strTopBar = strTopBar &	"</td></tr></table>"
	'strTopBar = strTopBar &	"<br />"
	%>
	
	<!-- #include virtual="/shared/page_header.inc" -->
	<form name="login" method="post" action="default.asp">
	<input type="hidden" name="check" value="0">
	<br />
	<table width="760" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td align="center">				
				<font class="boldTextBlack">Your session has either expired or you do not have permission to be here!</font>
			</td>
		</tr>
	</table>
	<br />
	<br />
	<table width="760" border="1" cellpadding="0" cellspacing="0">
		<tr>
			<td align="middle">			
				<table width="750" border="0" cellpadding="0" cellspacing="0">
				<tr>	
					<td rowspan="5" valign="middle">
						<img src="images/hhsc.jpg" alt="Hamilton Health Sciences" title="Hamilton Health Sciences" name="hhsc" alt"hhsc" />	
					</td>
					<td align="middle" colspan="3" valign="Middle">
						<br />
						<font class="headerBlack">Account Sign On</font>
						<br />						
						<br />
					</td>
					<td rowspan="5" valign="middle">
						<img src="images/fhslogo.jpg" width="150" alt="McMaster University Faculty of Health Sciences" title="McMaster University Faculty of Health Sciences" name="fhslogo" alt="fhs" />
					</td>
				</tr>
				<tr valign="top">
					<td width="100" align="right" nowrap="nowrap"> 
						<font class="boldtextblack">Email :&nbsp;&nbsp;</font>
					</td>
					<td width="175" align="left">
						<input type="text" name="email" value="<%=strEmail%>" size="25">						
					</td>
					<td width="275">
						<!-- default language set to English -->
						<%
						strLanguage = Request.Cookies("e-EDI")("Language")
						' sets default to english
						if srLanguage = "" then 
							strLanguage = "English"
						end if
						if strLanguage = "English" then 
							Response.Write "<INPUT type=""radio"" id=""language"" name=""language"" value=""English"" checked>"
							Response.Write "<font class=""boldtextblack"">English&nbsp;&nbsp;</font>"
							Response.Write "<INPUT type=""radio"" id=""language"" name=""language"" value=""French"">"
							Response.Write "<font class=""boldtextblack"">French&nbsp;&nbsp;</font>"
						elseif strLanguage = "French" then 
							Response.Write "<INPUT type=""radio"" id=""language"" name=""language"" value=""English"">"
							Response.Write "<font class=""boldtextblack"">English&nbsp;&nbsp;</font>"
							Response.Write "<INPUT type=""radio"" id=""language"" name=""language"" value=""French"" checked>"
							Response.Write "<font class=""boldtextblack"">French&nbsp;&nbsp;</font>"
						end if 
						%>
					</td>
				</tr>
				<tr>
				    <td width="100" align="right">
						<font class="boldtextblack" nowrap="nowrap">Password :&nbsp;&nbsp;</font>
					</td>
					<td width="175" align="left">
						<input type="password" name="password" value="" size="25">						
					</td>
					<td width="275">
						&nbsp;
						<input type="submit" name="Login" value="Login">
					</td>
				</tr>
				<tr>
					<td></td>
					<td colspan="2">
						<INPUT type="checkbox" id="savecookie" name="savecookie" checked>
						<font class="regtextblack" nowrap="nowrap">Save my settings</font>
					</td>
				</tr>
				</table>
				<br />
			</td>
		</tr>
		</table>
		<!-- #include virtual="/shared/page_footer.inc" -->
	</form>
	</body>
</html>
<%
	blnSecurity = false
else
	if Request.QueryString("Language") = "English" then
		session("language") = "English"
	elseif 	Request.QueryString("Language") = "French" then
		session("language") = "French"
	end if 
	
	blnSecurity = true
	strTopBar = "<table width=""760"" border=""1"" cellpadding=""1"" cellspacing=""0"" style=""border-color:#006600;background-color:#dcdcdc"">"
	strTopBar = strTopBar & "<tr><td><table width=""750"" border=""0"" cellpadding=""3"" cellspacing=""0""><tr><td align=""left""><font class=""boldTextBlack"">&nbsp;&nbsp;"
		call open_adodb(conn,"MACEDI")
		set rstUser = server.CreateObject("adodb.recordset")
			strQuery = "SELECT strName FROM users WHERE strEmail='" & session("id") & "'"
			rstUser.open strQuery,conn
			if not(rstUser.EOF) then
				strTopBar = strTopBar & trim(rstUser("strname"))
			else 
				strTopBar = strTopBar & "No user found"
			end if 
		call close_adodb(rstUser)
		call close_adodb(conn)
	strTopBar = strTopBar & "</font></td><td align=""right""><font class=""boldTextBlack"">" 
    
	if session("language") = "French" then 
		strTopBar = strTopBar & "le " & french_day(datepart("w",date(),vbSunday)) & " " & day(date()) & " " & French_month(month(date())) & " " & year(date())
		strDemographics = "Démographique"
		strExit = "Quitter"
		strHome = "Accueil"
		strSave = "enregistrer"
		strSaveIdentity = "Sauvegarder/Soumettre"
		lblName = "Nom"
		lblEmail = "Courriel"
		lblSex = "Sexe"
		lblAge = "Âge"
		lblPhone = "téléphone"
		lblLanguage = "Langue"
		lblAdd = "ajoutez"
		lblComments = "Commentaires"
		lblClassInfo = "Questionnaires IMDPE"
		lblClassCrumb = "Class"
		lblUpdate = "Mettre à jour"
		lblStatus = "État du questionnaire"
		lblConsent = "Consentement"
		lblLocal = "Identification locale"
		lblEDI = "IMDPE"
		lblIdentity = "Questionnaire sur le sentiment d’identité"
		lblIdentitySubHeader = "L’identité ayant une relation ou un lien avec:"
		lblPostal = "Code postal"
		lblDOB = "Date de naissance"
		lblEDIID = "ID d'IMDPE"
		lblSummary = "récapitulatif du dossier de l’élève"
		lblClassSummary = "récapitulatif de la classe"
		lblStudent = "Information sur l’élève"
		lblCancel = "Annulez"
		lblMale = "M"
		lblFemale = "F"
		lblTeacher = "Enseignant(e)"
		lblSchool = "École"
		lblSite = "Site"
		lblSaveEDI = "Enregistrer"
		lblFax = "Télécopieur"
		lblPass = "Mot de passe"
		lblPassword = "mot de passe"
		strMsg = "Pour modifier votre nom, courriel ou mot de passe, entrez les nouvelles données sur le texte existant et appuyer sur <b>Enregistrer</b>"
		strSaveMessage = "Voulez vous sauver le changement avant de sortir?"
		strLanguage = "French"
		intLanguage = 2
		strLink = replace(session("strLink"),"\","\\")
		'strLink = "documents\\French%20EDI%20Guide%202003.pdf"
		lblIncomplete = "inachevé et débloqué"
		lblComplete = "Incomplet et non vérouillé"
		strComplete = "complète"
		strIncomplete= "Incomplète"
		lblLock = "Fermez le dossier de l’élève"
		lblCompletion = "Vérifier l’état d’achèvement" '"Vérifiez la complétude"
		lblFinished = "Terminer – Soumettre à McMaster" 
		'strWarning = "Si l’élève est dans la classe moins d’un mois, quitté la classe, quitté l’école, ou autre, enregistrer et fermez le dossier de l’élève."
		strWarning = "Si l’enfant est dans votre classe depuis moins d’un mois, s’il n’est plus dans votre classe ou s’il a quitté l’école, ne complétez pas le reste du formulaire. Vérifiez l’état d’achèvement du questionnaire et faites le parvenir à McMaster."              
		'strWarning = "Étudiant doit être actuellement dans votre classe pour faire le EDI. Si l'enfant est actuellement dans votre classe, mais qui est là depuis moins d'un mois ne remplissez pas le reste du formulaire, enregistrer et verrouiller le questionnaire."
		strAlbertaWarning = "Les étudiants doivent présentement être dans votre classe et doivent avoir obtenu le consentement des parents afin de compléter l’IMDPE. Si ça fait moins d’un mois que l’enfant est dans votre classe et/ou si vous n’avez pas reçu le consentement des parents, ne complétez pas le reste du formulaire. Vérifiez l’état de complétude du questionnaire et faites le parvenir à McMaster."		
		lblNext = "Élève suivant"
		lblPrevious = "Élève précédant"
		lblAddStudent = "Ajouter un élève"
		lblTrainingFeedback = "Évaluation de la session de formation de l’IMDPE"
		lblCode = "Code de l'enseignant(e)"
		lblLogout = "Mettre fin à la session"	    
	
	    strConfirmLanguage = "&Ecirc;tes-vous s&ucirc;r de vouloir verrouiller le questionnaire de cet &eacute;l&egrave;ve?\n\nUne fois v&eacute;rouill&eacute;, vous ne pourrez plus modifier le questionnaire."
	    strConfirmLanguage2 = "Your questionnaire is now being sent to McMaster.  You will be returned to the class information page when complete."
        strQuestions = "Questions et commentaires"
        strABSpecialEduCode = "Please Note: Children in AB can only have ONE Special Education code.  You can only respond YES to one of the following questions(7, 8a, 8b). Refer to Guide for clarification."
        strCommentWarning = "Veuillez éviter d’utiliser les noms des enfants dans les commentaires." 
    '********************************************************
	' participation data 
		lblP1 = "Est-ce que c'est la première fois que vous participez à cette recherche en complétant les questionnaires IMDPE?"
		lblP2 = "Combien de fois dans le passé avez-vous participé?"
		lblP3 = "Est-ce que vous avez déjà participé à une session de formation pour enseignant(e)s?"
		lblP4 = "Si oui, combien de fois?"
		lblP5 = "Est-ce que vous avez reçu une formation pour la présente session de mise en oeuvre de l'IMDPE?"
		lblP6 = "Si oui, est-ce que la session de formation a été utile?"
		lblP7 = "Guide de l'enseignant(e) pour la mise en oeuvre de l'IMDPE (Veuillez cocher toutes les cases qui s'appliquent)"
		lblYesGoto = "Si oui, passez à la question no "					
		lblTitle = "Enseignant(e): formulaire de participation"
		lblTitleSubHeader = "Questions sur la participation de l'enseignant(e)"
		lbl4OrMore = "4 ou plus"
		lblVery = "beaucoup"
		lblSomewhat = "un peu"
		lblNotatall = "pas du tout"
		lblQuestion7YH	= "Je l'ai trouvé utile"
	    lblQuestion7YNH	= "Je ne l'ai pas trouvé très utile"
	    lblQuestion7NNH	= "Il ne me semblait pas très utile"
	    lblQuestion7NNone	= "Je n'en avais pas"
	    lblQuestion7NTime	= "Je n'avais pas suffisamment de temps"
	    lblQuestion7NFamiliar	= "Je le connais déjà"
	    lblQuestion7Other	= "Autre"	   
        lbl5 = "Expérience"
		lbl5a = "à titre d’enseignant(e)"
		lbl5b = "à titre d’enseignant(e) dans cette école"
		'lbl5c = "à titre d’enseignant â cette annèe d’études"
		lbl5c = "à titre d’enseignant(e) à ce niveau"
		lbl5d = "à titre d’enseignant(e) dans cette classe"	
		'lbl6 = "Veuillez spécifier le plus haut niveau de scolairité que vous avez atteints (cochez autant de réponses que nécessaires)"
		lbl6 = "Veuillez spécifier le plus haut niveau de scolarité que vous avez atteint (cochez toutes les réponses nécessaires)"
		lblA = "quelques cours en vue de l’obtention d’un baccalauréat"
		lblB = "un brevet d’enseignement"
		lblC = "un baccalauréat"
		lblD = "un baccalauréat en éducation"
		lblE = "quelques cours après le baccalauréat"
		lblF = "un diplôme ou un certicicat supérieur au baccalauréat"
		lblG = "quelques cours en vue de l’obtention d’un maîtrise"
		lblH = "une maîtrise"
		lblI = "quelques cours en vue de l’obtention d’un doctorat"
		lblJ = "un doctorat"
		lblK = "autres"
		lblYes = "oui"
		lblNo = "non"
		lblYrs = "années"
		lblMths = "mois"
		lblSize = "Nombre"
		lblComplete = "Accompli"
		lblID = "Identification de classe"
		lblProfile = "Profil de l'enseignant"
		lblDemo = "Démographique"
		lblParticipationDemo = "Caractéristiques sociodémographiques"
	    lblStudentsInClass = "Nombre total d'élèves dans cette classe: (Si votre classe est un cours double, prière dínclure tous les élèves)"
	    lblGender = "Sexe de l'enseignant(e)"
	' end participation data 
	'********************************************************
		lblPart1 = ""
		lblPart2 = ""
		lblPart3 = ""
		lblPart4 = ""
		lblPart5 = ""
		lblPart6 = ""
		lblPart7 = ""
		lblPartYes = ""
		lblPartNo = "non"		
		lblPartCode = "Code de l'enseignant"
		lblPartTitle = ""
		lblPartTitleSubHeader = ""
		lblPart4OrMore = ""
		lblPartVery = ""
		lblPartSomewhat = ""
		lblPartNotatall = ""
		lblPartQuestion7YH	= ""
	    lblPartQuestion7YNH	= ""
	    lblPartQuestion7NNH	= ""
	    lblPartQuestion7NNone	= ""
	    lblPartQuestion7NTime	= ""
	    lblPartQuestion7NFamiliar	= ""
	    lblPartQuestion7Other	= ""	   
	'*************************
    'Feedback
		lblFeedbackQ1 = "Est-ce la première fois que vous complétez l’IMDPE?"
		lblFeedbackQ2 = "Avez-vous déjà eu à compléter une version papier de l’IMDPE?"
		lblFeedbackQ3 = "Quelle version préférez-vous?"
	    lblFeedbackYesGoto = "Si oui, passez à la question "				
	    lblFeedbackNoGoto = "Si non, passez à la question "		
	    lblFeedbackYes = "Oui"
	    lblFeedbackNo = "Non"
	    lblFeedbackFeedback ="Prière de bien vouloir compléter le présent formulaire. Les renseignements recueillis nous aiderons à assurer la qualité des formations."
	    lblFeedbackElectronic = "Electronique (e-IMDPE)"
	    lblFeedbackPaper = "Papier"
        '********************
        'Class	
		lblClass = "Identification de la Classe"
		lblTime = "Temps de classe"
		aClassLanguage = array("","Anglais","Francais","Autre")
		strQuestions = "Pour ajouter une nouvelle classe, veuillez envoyer un message à l’administrateur de l’IMDPE"		                
		strNote2009 = "Si un étudiant a été ajouté par erreur, s'il vous plaît envoyez un message avec le numéro de l'EDI devant être supprimé à l'administrateur de l'EDI"		               
	    '***********************
        'Questionnaire        
		lblESpecifyPrint = "veuillez préciser lequel, si vous le connaissez"
        strLock = "Vérouillé"
	else
		strTopBar = strTopBar & formatdatetime(now,vblongdate) 
		strDemographics = "Demographics"
		strExit = "Exit"
		strHome = "Home"
		strSave = "Save"
		strSaveIdentity = "Save/Submit"
		lblName = "Teacher Name"		
		lblEmail = "Email"
		lblSex = "Sex"
		lblAge = "Age"
		lblLanguage = "Language"
		lblPhone = "Phone"
		lblAdd = "Add"
		lblComments = "Comments"
		lblClassInfo = "EDI Questionnaires"
		lblClassCrumb = "Class"
		lblUpdate = "Update"
		lblStatus = "Status"
		lblConsent = "Consent"
		lblLocal = "Local ID"
		lblEDI = "EDI"
		lblIdentity = "Sense of Identity"
		lblIdentitySubHeader = "Identity as relationships with and/or connections to:"
		lblPostal = "Postal Code"
		lblDOB = "Date of Birth"
		lblEDIID = "EDI ID"
		lblSummary = "View Child Summary"
		lblClassSummary = "View Class Summary"
		lblStudent = "Student Information"
		lblCancel = "Cancel"
		lblMale = "Male"
		lblFemale = "Female"
		lblteacher = "Teacher"
		lblSchool = "School"
		lblSite = "Site"
		lblSaveEDI = "Save EDI"
		lblFax = "Fax"
		lblPass = "Username\Password"
		lblPassword = "Password"
		strMsg = "To make a change to your name, e-mail or password, overwrite current entry and press Save button"
		strSaveMessage = "Do you want to save the change before exiting?"
		strLanguage = "English"
		intLanguage = 1
		'strLink = "documents\\EDI%20Guide%202003.pdf"
		strLink = replace(session("strLink"),"\","\\")
		lblIncomplete = "Incomplete and Unlocked"
		lblComplete = "Complete and Locked"
		strComplete = "Complete"
		strIncomplete= "Incomplete"
		lblLock = "Lock Child"
		lblCompletion = "Check for Completeness"
		lblFinished = "Finished/Submit to McMaster"
		'strWarning = "Student must be currently in your class to do the EDI.  If the child is currently in your class but has been there for less than one month do not complete the rest of the form, save and lock the questionnaire."
		'strWarning = "Student must currently be in your class to do the EDI. If the child is in your class but has been there less than a month, do not complete the rest of the form, check for completeness and then finish/submit to McMaster."
		strWarning = "Student must currently be in your class to do the EDI. If the child is in your class but has been there for less than a month, if he/she has changed classes or schools, do not complete the rest of the form. Check for completeness and the finish/sent to McMaster."
		'strAlbertaWarning = "Student must currently be in your class and have parental consent to do the EDI. If the child is in your class but has been there less than a month and/or you do not have parental consent, do not complete the rest of the form, check for completeness and then finish/submit to McMaster."
		strAlbertaWarning = "Student must currently be in your class and have parental consent to do the EDI. If the child is in your class but has been there for less than a month, if he/she has changed classes or schools, or if you do not have parental consent, do not complete the rest of the form. Check for completeness and then finish/send to McMaster."                             
		lblNext = "Next Student"
		lblPrevious = "Previous Student"
		lblAddStudent = "Add Student"
		lblTrainingFeedback = "e-Edi Teacher Training Feedback Form"
		lblCode = "Teacher Code"
        lblLogout = "Logout"
		strConfirmLanguage = "Are you sure you want to lock this student?\n\nOnce locked you will no longer be able to edit their EDI."
	    strConfirmLanguage2 = "Your questionnaire is now being sent to McMaster.  You will be returned to the class information page when complete."
        strQuestions = "Questions or Comments"
        strABSpecialEduCode = "Please Note: Children in AB can only have ONE Special Education code.  You can only respond YES to one of the following questions(7, 8a, 8b). Refer to Guide for clarification."
        strCommentWarning = "Please do not use children’s name in any comments."
        '********************************************************
	    ' participation data 
        lblP1 = "Is this your first time completing the EDI"
		lblP2 = "How many times previously have you completed the EDI?"
		lblP3 = "Did you attend a Teacher Training Session previously?"
		lblP4 = "If yes, how many times?"
		lblP5 = "Did you receive Teacher Training for this implementation?"	
		lblP6 = "If yes, how useful was it?"
		lblP7 = "EDI Teacher Guide Feedback (Please mark all that apply)"
		lblYesGoto = "Yes, go to question "				
		lblTitle = "Teacher Participation Form"
		lblTitleSubHeader = "Teacher Participation Questions"
		lbl4OrMore = "4 or more"		
		lblVery = "Very"
		lblSomewhat = "Somewhat"
		lblNotatall = "Not at all"
		lblQuestion7YH	= "Yes, I used the Guide and found it helpful" 
	    lblQuestion7YNH	= "Yes, I used the Guide but didn't find it helpful"
	    lblQuestion7NNH	= "No, I didn't use the Guide, I didn't find it helpful"
	    lblQuestion7NNone	= "No, I didn't use the Guide, I didn't have one"
	    lblQuestion7NTime	= "No, I didn't use the Guide, I didn't have enough time"
	    lblQuestion7NFamiliar	= "No, I didn't use the Guide, I'm already familiar with it"
	    lblQuestion7Other	= "Other"	    
        lbl5 = "Experience (How long have you been)"
		lbl5a = "a teacher"
		lbl5b = "a teacher at this school"
		lbl5c = "a teacher of this grade"
		lbl5d = "a teacher of this class"	
		lbl6 = "Completed levels of education(Check one or more if applicable)"
		lblA = "some coursework towards a Bachelor's degree"
		lblB = "a teaching certificate, diploma, or license"
		lblC = "a Bachelor's degree"
		lblD = "a Bachelor of Education degree"
		lblE = "some post-baccalaureate coursework"
		lblF = "a post-baccalaureate diploma or certificate"
		lblG = "some coursework towards a Master's degree"
		lblH = "a Master's degree"
		lblI = "some coursework towards a Doctorate"
		lblJ = "a Doctorate"
		lblK = "Other"
		lblYes = "Yes"
		lblNo = "No"
		lblYrs = "Yrs"
		lblMths = "Mths"
		lblSize = "Size"
		lblComplete = "Completed"
		lblID = "Class ID"
		lblProfile = "Teacher Profile"		
		lblDemo = "Demographics"
		lblParticipationDemo = "Demographics"
		lblStudentsInClass = "Total Number of Students"
		lblGender = "Teacher Gender"
        ' end participation data 
        '********************************************************
        lblPart1 = "Is this your first time completing the EDI"
		lblPart2 = "How many times previously have you completed the EDI?"
		lblPart3 = "Did you attend a Teacher Training Session previously?"
		lblPart4 = "If yes, how many times?"
		lblPart5 = "Did you receive Teacher Training for this implementation?"	
		lblPart6 = "If yes, how useful was it?"
		lblPart7 = "EDI Teacher Guide Feedback (Please mark all that apply)"
		lblPartYes = "Yes, go to question "
		lblPartNo = "No"	
		lblPartCode = "Teacher Code"
		lblPartTitle = "Teacher Participation"
		lblPartTitleSubHeader = "Teacher Participation Questions"
		lblPart4OrMore = "4 or more"		
		lblPartVery = "Very"
		lblPartSomewhat = "Somewhat"
		lblPartNotatall = "Not at all"
		lblPartQuestion7YH	= "Yes, I used the Guide and found it helpful" 
	    lblPartQuestion7YNH	= "Yes, I used the Guide but didn't find it helpful"
	    lblPartQuestion7NNH	= "No, I didn't use the Guide, I didn't find it helpful"
	    lblPartQuestion7NNone	= "No, I didn't use the Guide, I didn't have one"
	    lblPartQuestion7NTime	= "No, I didn't use the Guide, I didn't have enough time"
	    lblPartQuestion7NFamiliar	= "No, I didn't use the Guide, I'm already familiar with it"
	    lblPartQuestion7Other	= "Other"	    
	'*************************
    'Feedback
		lblFeedbackQ1 = "Is this your first time completing the EDI?"
		lblFeedbackQ2 = "Have you completed the paper version of the EDI?"
		lblFeedbackQ3 = "Which version did you prefer?"
		lblFeedbackYesGoto = "Yes, go to question "	
		lblFeedbackNo = "No"	
		lblFeedbackNoGoto = "No, go to question "		
	    lblFeedbackYes = "Yes"
		lblFeedbackFeedback ="Please take the time to complete the following Teacher Training Feedback Form.  The information gathered from this form will help us to ensure high quality teacher training practices."
		lblFeedbackElectronic = "Electronic (e-EDI)"
	    lblFeedbackPaper = "Paper"

        '********************
        'Class
	    lblClass = "Class Code"
		lblTime = "Class Time"
		aClassLanguage = array("","English","session","Other")						
		strQuestions = "To add a new class, please send a message to the EDI Administrator"
		strNote2009 = "If a student has been added in error, please send a message with the EDI Number to be deleted to the EDI Administrator"
	    '***********************
        'Questionnaire
        lblESpecifyPrint = "Specify if known:"
        strLock = "Lock"
	end if 
		
	strTopBar = strTopBar & " - " & time() & "&nbsp;&nbsp;</font></td></tr></table>"	
	strTopBar = strTopBar &	"</td></tr></table>"
	'strTopBar = strTopBar &	"<br />"
end if 	


function French_Month(intMonth)
	select case intMonth
		case 1 
			French_Month = "janvier"
		case 2
			French_Month = "février"
		case 3
			French_Month = "mars"
		case 4
			French_Month = "avril"
		case 5 
			French_Month = "mai"
		case 6 
			French_Month = "juin"
		case 7 
			French_Month = "juillet"
		case 8 
			French_Month = "août"
		case 9 
			French_Month = "septembre"
		case 10 
			French_Month = "octobre"
		case 11
			French_Month = "novembre"
		case 12 
			French_Month = "décembre"
	end Select
end function

function French_Day(intDay)
	select case intDay
		case 1 
			French_Day = "dimanche"
		case 2
			French_Day = "lundi"
		case 3 
			French_Day = "mardi"
		case 4 
			French_Day = "mercredi"
		case 5
			French_Day = "jeudi"
		case 6 
			French_Day = "vendredi"
		case 7 
			French_Day = "samedi"
	end select 
end Function

%>