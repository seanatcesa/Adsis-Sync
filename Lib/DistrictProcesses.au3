;--DISTRICT-SPECIFIC FUNCTIONS
;----
;These functions provide rules and logic that may be
;unique to each district
;--



;Function:	UserCreation_RequirementsVerify
;--
;Purpose:	To verify for a certain student record that all requirements are met before creating the user
;--
;District:	SCHOOL DISTRICT OF HORICON
;--
;Date:		July 02, 2015
;
;Returns:	On Success:	1
;			On Failure:	0
Func UserCreation_RequirementsVerify($StudentNumber, $FirstName, $MiddleName, $LastName, $GradYear, $LunchID)
   Local $iErrorCount = 0

   logwrite('UserCreation Requirements: Evaluating new record for Student ' & $FirstName & ' ' & $LastName & '(Number:' & $StudentNumber & ')')
   If StringLen($StudentNumber) < 3 Then
	  ;Error.  Student number appears to be invalid
	  logwrite('UserCreation Requirements: ID Number ' & $StudentNumber & ' is not valid.', 'WARNING')
	  $iErrorCount = $iErrorCount + 1
   EndIf

   If StringLen($LastName) < 2 Then
	  ;Error.  Last Name appears to be invalid
	  logwrite('UserCreation Requirements: Last name ' & $LastName & ' is not valid.', 'WARNING')
	  $iErrorCount = $iErrorCount + 1
   EndIf

   If StringLen($FirstName) < 2 Then
	  ;Error.  First Name appears to be invalid
	  logwrite('UserCreation Requirements: First name ' & $FirstName & ' is not valid.', 'WARNING')
	  $iErrorCount = $iErrorCount + 1
   EndIf

   If StringLen($LunchID) < 3 Then
	  ;Error.  LunchID appears to be invalid
	  logwrite('UserCreation Requirements: Lunch ID ' & $LunchID & ' is not valid.', 'WARNING')
	  $iErrorCount = $iErrorCount + 1
   EndIf

   If $GradYear > (@YEAR + 20) OR $GradYear < @YEAR -1 Then	; Graduation is more than 20 years in the future or exists in the past
	  LogWrite('UserCreation Requirements: Graduation year ' & $GradYear & ' is invalid.', 'ERROR')
	  $iErrorCount = $iErrorCount + 1
   EndIf

   If GradYearToGradeLevel($GradYear) < $Rule_UserCreation_MinimumGradeLevel Then
	  ;Notice.  User's grade level is lower than the minimum grade level for new users
	  logwrite('UserCreation Requirements: Student grade level is ' & GradYearToGradeLevel($GradYear) & '.  The minimum grade level for an account is ' & $Rule_UserCreation_MinimumGradeLevel, 'WARNING')
	  $iErrorCount = $iErrorCount + 1
   EndIf

   If $iErrorCount > 0 Then
	  ;User not created.
	  logwrite('UserCreation Requirements: Account not created because of the ' & $iErrorCount &  ' messages previously reported.')
	  Return 0
   Else
	  logwrite('UserCreation Requirements: User meets requirements for account creation.')
	  Return 1
   EndIf

EndFunc


; Function:		Username_Build
; Purpose:		Given the components of a username, this function is
;				the algorithm that determines a user's username
;
; Notes:		Use of this function will vary from district to district.
;				THIS VERSION WAS BUILT FOR SCHOOL DISTRICT OF HORICON
;				The format they use is: [2 digit grad year][First 2 of First name]{Additional Chars of First, then Middle name on duplication}[Full Last name]
;
; Parameters:	$sFirstName		- User's complete first name
;				$sLastName		- User's complete last name
;				$iGradYear		- 4 digit representation of the student's HS Grad year
;				$iRepeatCount	- The number of times this algorithm has been tried for this user
;								  On the first execution for a user, this should be set to zero.
;								  If a duplicate is found, it should be re-run with 1
;								  If a duplicate is still found, it should be incremented by 1 with each iteration
;								  If a duplicate is still found after we have run out of entropy, then we should throw an error and deal with it manually.
; Return Value(s):  On Success	-	Returns the new username
;
;					On Failure	-	Returns 0.
;								-	@error = 1 Grad year invalid
;								-	@error = 2 First name Invalid or Middle Name invalid for deduplication
;								-	@error = 3 Last name invalid
;									retreived.
Func Username_Build($sFirstName, $sMiddleName, $sLastName, $iGradYear, $iRepeat=0)
   Local $iErrorCount = 0
   logwrite('Building Username for ' & $sFirstName & ' ' & $sMiddleName & ' ' & $sLastName)
   ;------Component 1: Grad Year----------------

   ;--Sanity check
   If $iGradYear > (@YEAR + 20) OR $iGradYear < @YEAR -1 Then
	  LogWrite('Username_Build: Graduation year ' & $iGradYear & ' is invalid.', 'ERROR')
	  $iErrorCount = $iErrorCount + 1
   EndIf

   $Comp1 = StringRight($iGradYear, 2)



   ;------Component 2: First Name and Deduplication----------------
   ;Note: This section is incomplete!!!!  Add collision handling logic.

   ;--Sanity Check
   ;Firstname must be at least 2 characters long
   If StringLen($sFirstname) < 2 Then
	  LogWrite('Username_Build: First Name ' & $sFirstName & ' is invalid.  It should be at least 2 letters long', 'ERROR')
	  $iErrorCount = $iErrorCount + 1
   EndIf
   $Comp2 = StringLeft($sFirstName, 2)


   ;--Handle weirdness
   ;Strip leading, trailing, and double whitespace
   $sFirstname = StringStripWS($sFirstname, 7)




;-----Component 3: Last Name----------------

   ;--Sanity Check
   If StringLen($sLastName) < 2 Then
	  LogWrite('Username_Build: Last Name ' & $sLastName & ' is invalid.  It should be at least 2 letters long', 'ERROR')
	  $iErrorCount = $iErrorCount + 1
   EndIf

   ;Strip leading, trailing, and double whitespace
   $sLastName = StringStripWS($sLastName, 7)

   ;Strip suffixes from the last name
   $sLastName = StringReplace($sLastName, " JR.", "")
   $sLastName = StringReplace($sLastName, " JR", "")
   $sLastName = StringReplace($sLastName, " III", "")
   $sLastName = StringRegExpReplace($sLastName, "[^a-zA-Z]", "") ;Remove any non-alpha characters

   $Comp3 = $sLastName


   ;-----Combined Name----------------------
   $TentativeUsername = $Comp1 & $Comp2 & $Comp3
   $TentativeUsername  = StringLower($TentativeUsername )
   LogWrite('Username_Build: Tentative username is ' & $TentativeUsername )

   ;--Name collision detection
   ;--Horicon District resolves username collisions by extending the first name, then the middle name, then a sequential number until no collision is found
   If AD_UserExists($TentativeUsername , $rAD, CSVHeader($rAD, $AD_UserNameField)) OR PendingNewUser_Exists($TentativeUsername) Then
	  ;Uh-oh!  We have a naming collision

	  $iConflictCount = 0 	; Initialize our loop counter
	  $Comp4 = 0			; Initialize $Comp4 in case we need to add a digit to the username
	  While AD_UserExists($TentativeUsername, $rAD, CSVHeader($rAD, $AD_UserNameField)) OR PendingNewUser_Exists($TentativeUsername)
		 LogWrite('Username_Build: Username collision found at ' & $TentativeUsername, 'WARNING' )
		 If StringLen($sFirstName) > StringLen($Comp1) Then
			;We still have some letters from the first name to add.  Add another letter from the first name.
			$Comp2 = StringLeft($sFirstName, StringLen($Comp2) + 1)
			LogWrite('Username_Build: Adding one more letter to the firstname component.  Component result: ' & $Comp2)

			$TentativeUsername = $Comp1 & $Comp2 & $Comp3
			$TentativeUsername  = StringLower($TentativeUsername )
			LogWrite('Username_Build: Tentative username is ' & $TentativeUsername )
		 Else
			;We are out of letters in the name.  Start adding digits.
			LogWrite('Username_Build: Component letters are exhausted.  Now testing appended digits.')
			$Comp4 = $Comp4 + 1

			$TentativeUsername = $Comp1 & $Comp2 & $Comp3 & $Comp4
			$TentativeUsername  = StringLower($TentativeUsername )
			LogWrite('Username_Build: Tentative username is ' & $TentativeUsername )
		 EndIf

	  WEnd

	  LogWrite('Username_Build: Username collision resolved.  New username is: ' & $TentativeUsername, 'WARNING' )

   EndIf

   ;Now that we have run our username through collsion detection, we can treat it as the completed username.  Move it to the $CompletedUsername variable.
   $CompletedUsername = $TentativeUsername

   If $iErrorCount >0 Then
	  LogWrite('Username_Build: Username could not be built because of ' & $iErrorCount & ' errors.', 'ERROR')
	  return 0
   Else
	  LogWrite('Username_Build: Generated username: ' & $CompletedUsername)
	  return $CompletedUsername
   EndIf

EndFunc ;==>Username_Build


;Function:	UserCreate
;---
;
Func UserCreate($StudentNumber, $FirstName, $MiddleName, $LastName, $GradYear, $Password, $username)
   $c 	= 'DSADD USER '
   $c	&= ' "CN='&$FirstName & ' ' & $LastName &',OU='&$GradYear&','&$BASEDN_Users&'"  '
   $c	&= ' -UPN "' & $username & '@' & $WindowsDomain & '" '
   $c	&= ' -samid "'& $username &'" '
   $c	&= ' -fn "'& $FirstName &'" '
   $c	&= ' -ln "'& $LastName &'" '
   $c	&= ' -display "'& $Firstname & ' ' & $LastName &'" '
   $c	&= ' -empid "' & $StudentNumber & '" '
   $c	&= ' -pwd "'&$Password&'" '
   $c	&= ' -desc "CLASS OF '&$GradYear&'" '
   $c	&= ' -email "'&$username & '@' & $EmailDomain&'" '
   $c	&= ' -memberof "CN=Students-'&$GradYear&','&$BaseDN_Groups&'" '
   $c	&= ' -hmdir "\\'&$WindowsDomain&'\Student\Home\'&$GradYear&'\'&$username&'" '
   $c	&= ' -hmdrv H: '
   $c	&= ' -loscr Student.bat '
   $c	&= ' -pwdneverexpires yes '
   $c	&= ' -canchpwd no '
   $c	&= ' -disabled no '
   $c	&= ' -mustchpwd no'

   LogWrite('UserCreate: ' & $c)

   If $Simulation = False Then
	  ShellExecuteWait($c)
   Else
	  LogWrite('(Simulation only.  No changes written.)')
   EndIf
   ;Add the new user to our list of Pending New Users
   PendingNewUser_Add($username)
   Return $c

EndFunc;==>UserCreation_CommandBuild

;Function:	HomeDirCreate
;---
;
Func HomeDirCreate($Username, $GradYear)
   $homedir = '\\' & $WindowsDomain & '\Student\Home\' & $GradYear & '\' & $Username
   LogWrite('HomeDirCreate: Creating home directory for user ' & $username & ' : ' & $homedir)
   If $Simulation = False Then
	  DirCreate($homedir)
	  Sleep(100)	; Changing the ACL on the directory too quickly has caused the cacls command to fail on occasion.
   Else
	  LogWrite('(Simulation only.  No changes written.)')
   EndIf

   $c	= 'cacls "' & $homedir & '" /E /C /G ' & $username & '@' & $WindowsDomain & ':F'
   LogWrite('HomeDirCreate: Granted permissions to home directory with this command: ' & $c)
   If $Simulation = False Then
	  ShellExecuteWait($c)
   Else
	  LogWrite('(Simulation only.  No changes written.)')
   EndIf


EndFunc;==>HomeDirCreation_CommandBuild

;Function:	UserDecommission
;---
;Purpose:	Delete, Disable, or otherwise process a user account that is no longer needed
;			This process will vary depending on District policy
;---
;District:	SCHOOL DISTRICT OF HORICON
;---
;Input:		$DN:		String. The full DistinguishedName of the user to be decommissioned.
;						No quotes should be placed around the string.
;			$username:	String.  The username of the account to be decommissioned.
;						This makes for easier reporting on what users were decommissioned on the last execution.
Func UserDecommission($DN, $username)
   If StringLen($DN) < 5 OR StringLen($username) < 2 Then
	  LogWrite('UserDecommission: Blank or Invalid DN or Username was passed to UserDecommission().  Ignored.', 'ERROR')
	  Return 0
   EndIf

   LogWrite('UserDecommission:  Decommissioning user account ' & $username)
   $c = 'dsmod user "' & $DN & '" -disabled yes'
   LogWrite('UserDecommission:  Command: ' & $c)
   If $Simulation = False Then
	  $return = ShellExecuteWait($c)
   Else
	  LogWrite('(Simulation only.  No changes written.)')
   EndIf

   $note = 'DISABLED AUTOMATICALLY BY ' & $sProductName & ' on ' & @MON & '-' & @MDAY & '-' & @YEAR
   $n = 'dsmod user "' & $DN & '" -desc "' & $note & '"'
   LogWrite('UserDecommission:  Adding note to account description: ' & $n)
   LogWrite('UserDecommission:  Command: ' & $n)
   If $Simulation = False Then
	  $return = ShellExecuteWait($n)
   Else
	  LogWrite('(Simulation only.  No changes written.)')
   EndIf

EndFunc














