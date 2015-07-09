;Current state: So far new user creation works in a neutered state.  It should properly generate all commands.
;To do:
;			- Keep/generate a running list of decomissioned users for reporting.
;			- Handle changed users
;			- Handle deleted users
;			- Prevent running twice on the same CSV
;			- Perhaps handle or at least report errors that occur during account creation
;			- Email reports when finished

#include <file.au3>
#include <Array.au3>
#include ".\Lib\DistrictProcesses.au3"	; Functions that should be customized for each district to reflect their policies and practices
#include ".\Lib\IntegratedSupport.au3"	; Custom Functions that are integral to the operation of this program
#include ".\Lib\MD5.au3"				; Third-party MD5 support

;Project:		Adsis Sync
;Version:
;
;Purpose:		This program should facilitate the integration of a Student Information System with Active Directory
;
;
;
;Author:		Sean Nelson (snelson@cesa6.org)
;				Systems Administrator, CESA 6
;
;Support:		Support is available from Cooperative Educational Services Agency #6
;
;---
;Version History:



;--------------------------------------------
;--OPTIONS-----------------------------
;--
Global $sConfigFileName		= @ScriptDir & '\' & 'Config.ini'

;====Configuration that could be moved to an external file in the future
Global $Simulation				= True ; If true, this program will simulate action and write logs, but not write changes to Active Directory.
Global $CurrentGradYear			= 2016
Global $sADExportFileName 		= "C:\Users\Default.Default-PC\Desktop\Horicon SIS to AD via CSV\From_AD_7-6-2015_1.csv"
Global $sSISExportFilename		= "C:\Users\Default.Default-PC\Desktop\Horicon SIS to AD via CSV\From_SIS_7-6-2015_1.csv"

Global $EmailDomain				= 'horicon.k12.wi.us'
Global $WindowsDomain			= 'horicon.k12.wi.us'
Global $DomainControllerAddress	= 'HS-DC-1.horicon.k12.wi.us'
Global $BaseDN_Users			= 'OU=Students,OU=Users,OU=SDH,DC=horicon,DC=k12,DC=wi,DC=us';Base DN where users affected by this program reside
Global $BaseDN_Groups			= 'OU=SDH,DC=horicon,DC=k12,DC=wi,DC=us' ;Base DN where groups used by this program reside
Global $sDefaultFolder 			= ".\" ;The default folder to search for AD and SIS data files

;=Rules
Global $Rule_UserCreation_MinimumGradeLevel = 2

;===Options that can be built-in to the compiled script
Global $sStartTimeStamp		= @YEAR & '-' & @MON & '-' & @MDAY & '_' & @HOUR & '-' & @MIN & '-' & @SEC		; Time stamp to mark the start of program execution.  This can be used to name output files.
Global $sProductName 		= "Adsis Sync"
Global $sLogDirectory		= ".\Logs"	; Directory where execution logs are stored.  NO TRAILING SLASH.
Global $sLogfilePath		= $sLogDirectory & '\' & $sProductName & '_' & $sStartTimestamp & '.log'

;=Field Definitions
Global $SIS_CSVHeader				= 'Student_Number,Last_Name,First_Name,Grade_Level,SchoolID,Enroll_Status,ClassOf,Lunch_ID,Home_Room';To be used when the CSV is generated with no field headers.  Enter a replacement CSV header to be used instead.  (And make sure it is valid!)
Global $SIS_IDNumberField			= 'Student_Number'
Global $SIS_FirstNameField			= 'First_Name'
Global $SIS_LastNameField			= 'Last_Name'
Global $SIS_UserNameField			= ''
Global $SIS_GradYearField			= 'ClassOf'
Global $SIS_LunchIDField			= 'Lunch_ID'
Global $SIS_EmailField				= ''
Global $SIS_DistinguishedNameField	= ''
Global $SIS_HomeRoomField			= 'Home_Room'


Global $AD_CSVHeader				= '';To be used when the CSV is generated with no field headers.  Enter a replacement CSV header to be used instead.  (And make sure it is valid!)
Global $AD_IDNumberField			= 'EmployeeID'
Global $AD_FirstNameField			= 'GivenName'
Global $AD_LastNameField			= 'Surname'
Global $AD_UserNameField			= 'SamAccountName'
Global $AD_GradYearField			= ''
Global $AD_LunchIDField				= ''
Global $AD_EmailField				= 'mail'
Global $AD_DistinguishedNameField	= 'DistinguishedName'
Global $AD_HomeRoomField			= ''

;=Update Field Definitions with Prefixes
Global $SIS_CSVFieldPrefix	= 'SIS_'	; For internal use.  This prefix prevents name collisions among CSV headers
$SIS_IDNumberField			= $SIS_CSVFieldPrefix & $SIS_IDNumberField
$SIS_FirstNameField			= $SIS_CSVFieldPrefix & $SIS_FirstNameField
$SIS_LastNameField			= $SIS_CSVFieldPrefix & $SIS_LastNameField
$SIS_UserNameField			= $SIS_CSVFieldPrefix & $SIS_UserNameField
$SIS_GradYearField			= $SIS_CSVFieldPrefix & $SIS_GradYearField
$SIS_LunchIDField			= $SIS_CSVFieldPrefix & $SIS_LunchIDField
$SIS_EmailField				= $SIS_CSVFieldPrefix & $SIS_EmailField

Global $AD_CSVFieldPrefix	= 'AD_'		; For internal use.  This prefix prevents name collisions among CSV headers
$AD_IDNumberField			= $AD_CSVFieldPrefix & $AD_IDNumberField
$AD_FirstNameField			= $AD_CSVFieldPrefix & $AD_FirstNameField
$AD_LastNameField			= $AD_CSVFieldPrefix & $AD_LastNameField
$AD_UserNameField			= $AD_CSVFieldPrefix & $AD_UserNameField
$AD_GradYearField			= $AD_CSVFieldPrefix & $AD_GradYearField
$AD_LunchIDField			= $AD_CSVFieldPrefix & $AD_LunchIDField
$AD_EmailField				= $AD_CSVFieldPrefix & $AD_EmailField
$AD_DistinguishedNameField 	= $AD_CSVFieldPrefix & $AD_DistinguishedNameField


;--------------------------------------------
;---INITIALIZE-------------------------------
;--

;==Initialize list of to-be-created usernames
;	This is necessary because we check for username conflicts against the latest Active Directory export, not against live data.
;	So there is the possibility that we create a new user, then later in the list attempt to create a conflicting username
;	without seeing that a conflict exists.  Thus, conflict checks should be made both against the AD CSV export as well as the
;	below Array
Global $PendingNewUsers[0]
Global $PendingDecommissionedUsers[0]




;--------------------------------------------
;---PROCESSING-------------------------------
;--

;Fixup the headers.  If $SIS_CSVHeader or $AD_CSVHeader are defined, we need to add those headers to the CSV file.
If StringLen($SIS_CSVHeader) > 3 Then
   LogWrite('$SIS_CSVHeader is defined.  Checking to see if the header needs to be replaced on ' & $sSISExportFilename)
   ;$SIS_CSVHeader has been set.  Prepend it to the file before we process it.

   ;Read the first line of the file
   $sSISExportFileLineOne = FileReadLine($sSISExportFilename)

   ;Compare to our header to be sure we haven't already written it previously.
   If Not (StringStripWS($sSISExportFileLineOne,8) = StringStripWS($SIS_CSVHeader,8)) Then
	  ;Line 1 is not the same as our header.  Replace it.
	  LogWrite('Adding this header to ' & $sSISExportFilename & ' : ' & $SIS_CSVHeader)
	  FilePrepend($sSISExportFilename, $SIS_CSVHeader & @CRLF)
   Else
	  ;Line 1 is the same as our header.  Continue.
	  LogWrite('The header in ' & $sSISExportFilename & ' is already identical to $SIS_CSVHeader. Ignoring the header this time.')
   EndIf


EndIf

;Import SIS and AD records to their respective arrays
$rSIS	= CSVFileToArray($sSISExportFilename, $SIS_CSVFieldPrefix)
$rAD 	= CSVFileToArray($sADExportFileName, $AD_CSVFieldPrefix)


;===PROCESS NEW USERS===
LogWrite('')
LogWrite('-------------------------------------------------------------------')
LogWrite('--------NOW DISCOVERING NEW USERS----------------------------------')
LogWrite('-------------------------------------------------------------------')
LogWrite('')


;Compare SIS with AD.
;Compare AD with our latest SIS data
$rSISAD = ArrayJoin($rSIS, $rAD, $SIS_IDNumberField, $AD_IDNumberField, 'rSISADMatches', 'rSISADMisses')
;_ArrayDisplay($rSISADMatches, 'MATCHES')
;_ArrayDisplay($rSISADMisses, 'MISSES')



;Process each record
;	IF the user is present in the MISS table of an ArrayJoin of the SIS and AD users
;	AND the user passes the NewUserRequirements
;	Then create the user in AD.
For $i = 1 to UBound($rSISADMisses) - 1	;Start count at 1 because of header in row 0.
   LogWrite('-----------------DISCOVERED USERS: Processing new record---------------')
   $IDNumber			= $rSISADMisses[$i][CSVHeader($rSISADMisses, $SIS_IDNumberField)]
   $FirstName			= $rSISADMisses[$i][CSVHeader($rSISADMisses, $SIS_FirstNameField)]
   $MiddleName			= "";Not currently supported!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
   $LastName			= $rSISADMisses[$i][CSVHeader($rSISADMisses, $SIS_LastNameField)]
   $GradYear			= $rSISADMisses[$i][CSVHeader($rSISADMisses, $SIS_GradYearField)]
   $LunchID				= $rSISADMisses[$i][CSVHeader($rSISADMisses, $SIS_LunchIDField)]

   If UserCreation_RequirementsVerify($IDNumber, $FirstName, $MiddleName, $LastName, $GradYear, $LunchID) = 1 Then
	  ;The user meets requirements to have an account created.  Let's make it.
	  logwrite('UserCreation_Requirements Verify as successful.  Continuing with user account creation.')
	  $sUsername = Username_Build($Firstname, $MiddleName, $LastName, $GradYear)
	  UserCreate($IDNumber, $FirstName, $MiddleName, $LastName, $GradYear, $LunchID, $sUsername)
	  HomeDirCreate($sUsername, $GradYear)
   Else
	  ;The user did not meet the requirements or somethign went wrong.  This user will not get created.  More details in the log.
	  logwrite('Skipping user creation for ' & $IDNumber & ' because of errors reported by UserCreation_RequirementsVerify()')

   EndIf


Next


;===PROCESS MODIFIED USERS===
;... This may have to wait...







;===PROCESS DELETED USERS===
LogWrite('-------------------------------------------------------------------')
LogWrite('-------------------------------------------------------------------')
LogWrite('-------------------------------------------------------------------')
LogWrite('--------NOW DISCOVERING DELETED USERS------------------------------')
LogWrite('-------------------------------------------------------------------')
LogWrite('-------------------------------------------------------------------')
LogWrite('-------------------------------------------------------------------')
$rSISAD = ArrayJoin($rAD, $rSIS, $AD_IDNumberField,$SIS_IDNumberField, 'rADSISMatches', 'rADSISMisses')

;_ArrayDisplay($rADSISMatches, 'MATCHES')	; Users who exist in AD and SIS
;_ArrayDisplay($rADSISMisses, 'MISSES')		; Users who exist in AD but NOT SIS.  Since SIS is authoritative, the AD record needs to be handled.

;--Process each record
For $i = 1 to UBound($rADSISMisses) - 1		;Start count at 1 because of header in row 0.
   LogWrite('-----------------DELETED USERS: Processing new record---------------')
   $DN = $rADSISMisses[$i][CSVHeader($rADSISMisses, $AD_DistinguishedNameField)]
   $username = $rADSISMisses[$i][CSVHeader($rADSISMisses, $AD_UserNameField)]
   UserDecommission($DN, $username)
Next





;===Output summary of actions===
;_ArrayDisplay($PendingNewUsers)
LogWrite('')
LogWrite('')
LogWrite('')
LogWrite('-----------------------------------------------------------------')
LogWrite('SUMMARY OF ACTIONS:')
LogWrite('====New Users Created: ' & UBound($PendingNewUsers))
For $i = 0 to UBound($PendingNewUsers) - 1
   LogWrite('     ' & $PendingNewUsers[$i])
Next



;===Move our input CSV files to a time-stamped name so we don't run them again by accident
LogWrite('')
LogWrite('Cleaning up...')

$sADExportNewFileName = StringStripFileExt($sADExportFileName) & $sStartTimeStamp & '.csv'
LogWrite('Moving input file ' & $sADExportFileName & ' to ' & $sADExportNewFileName)
If $Simulation = False Then
   FileMove($sADExportFileName, $sADExportNewFileName)
Else
   LogWrite('(Simulation only.  No changes written.)')
EndIf

$sSISExportNewFileName = StringStripFileExt($sSISExportFileName) & $sStartTimeStamp & '.csv'
LogWrite('Moving input file ' & $sSISExportFileName & ' to ' & $sSISExportNewFileName)
If $Simulation = False Then
   FileMove($sSISExportFileName, $sSISExportNewFileName)
Else
   LogWrite('(Simulation only.  No changes written.)')
EndIf



