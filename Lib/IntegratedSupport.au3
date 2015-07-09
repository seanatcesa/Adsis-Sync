
;--------------------------------------------
;--FUNCTIONS---------------------------------
;--


;Function:		ArrayJoin
;---
;Purpose:		Joins two 2D arrays together according to a match on a defined column
;				The effect is similar to a SQL Join
;---
;Syntax:		ArrayJoin($Array1, $Array2, $Array1JoinField, $Array2JoinField, $Array1HeaderPrefix, $Array2HeaderPrefix, $MatchArrayName, $MissArrayName)
;
;--
;Example:		ArrayJoin($rSIS, $rAD, $SIS_IDNumberField, $AD_IDNumberField, 'SIS_', 'AD_', 'rMatches', 'rMisses')
;
;---
;Input:			$Array1:			Array. The array we will iterate through to develop a list of Matches/Misses
;
;				$Array2:			Array. The array we will search $Array1 against
;
;				$Array1JoinField:	String. The header name of the column to be used as a search criterea from Array1
;
;				$Array2JoinField:	String. The header name of the column of Array 2 to be searched using the value from the JoinField of Array1.
;
;				$MatchArrayName:	String.  The name of the variable where ArrayJoin should place its matches
;
;				$MissArrayName:		String.  The name of the variable where ArrayJoin should save its misses
;									(Misses are rows from Array1 that had no match in Array2)
;---
;Return:		On success: 		Integer.  The number of rows processed from Array1
;
;				On failure:			Integer. -1
;---
Func ArrayJoin($Array1, $Array2, $Array1JoinField, $Array2JoinField, $MatchArrayName, $MissArrayName)

   ;Convert the Header Names to Column Numbers
   $Array1JoinField = CSVHeader($Array1, $Array1JoinField)
   $Array2JoinField = CSVHeader($Array2, $Array2JoinField)

   ;Determine the number of columns the new Array should have
   $iArray1Columns = UBound($Array1, 2)
   $iArray2Columns = UBound($Array2, 2)
   $iMatchArrayColumns = $iArray1Columns + $iArray2Columns

   ;Define our Array to contain our matches
   Local $rMatch[1][$iMatchArrayColumns]

   ;Define our Array to contain our misses
   Local $rMiss[1][$iArray1Columns]

   ;Add our combined column headers to the new rMatch array
   For $i=0 to $iMatchArrayColumns - 1
	  If $i < $iArray1Columns Then
		 ;Populate the column headers of our Match array with headers from Array1
		 $rMatch[0][$i] = $Array1[0][$i]
	  Else
		 ;Populate the remaining column headers of our Match array with the headers from Array2
		 $rMatch[0][$i] = $Array2[0][$i - $iArray1Columns]
	  EndIf

   Next

   ;Copy our column headers from Array1 to our Miss array
   For $iColumn = 0 to $iArray1Columns - 1
	  $rMiss[0][$iColumn] = $Array1[0][$iColumn]
   Next

   ;Initialize our counters for the next loop operation
   Local $iMatchCount = 0	; Counter to keep track of matches
   Local $iMissCount = 0		; Counter to keep track of the misses (Searches that yielded no match)

   ;Loop through the Data in Array 1 JoinField, searching for matches in Array2 JoinField
   ;We expect a data header in the Array, so we'll start at row 1.
   For $iSearchRow = 1 to (UBound($Array1) - 1)

	  ;Search Array2 for the contents of our JoinField in Array1
	  $resultRow = _ArraySearch($Array2, $Array1[$iSearchRow][$Array1JoinField],0, 0, 0, 0, 1, $Array2JoinField, 0)
	  ;_ArraySearch ( Const ByRef $aArray, $vValue [, $iStart = 0 [, $iEnd = 0 [, $iCase = 0 [, $iCompare = 0 [, $iForward = 1 [, $iSubItem = -1 [, $bRow = False]]]]]]] )
	  If $resultRow > -1 Then
		 ;We found a match!
		 ;Increment the match counter
		 $iMatchCount = $iMatchCount + 1
		 ;MsgBox(0, 0, 'Array1 row ' & $iSearchRow & '(' & $Array1[$iSearchRow][$Array1JoinField] & ') has a match in Array 2, row ' & $resultRow)

		 ;Make room in the array for our new match
		 ReDim $rMatch[$iMatchCount + 1][$iMatchArrayColumns]
		 ;_ArrayDisplay($rMatch, 'Resized the $rMatch array to ' & $iMatchCount + 1 & ' rows.')

		 ;Add the new match to our Match array
		 For $iColumn = 0 to $iMatchArrayColumns - 1
			If $iColumn < $iArray1Columns Then
			   ;Populate the column headers of our Match array with headers from Array1
			   $rMatch[$iMatchCount][$iColumn] = $Array1[$iSearchRow][$iColumn]
			Else
			   ;Populate the remaining column headers of our Match array with the headers from Array2
			   $rMatch[$iMatchCount][$iColumn] = $Array2[$resultrow][$iColumn - $iArray1Columns]
			EndIf
		 Next


	  Else
		 ;We found NO match
		 ;Increment the miss counter
		 $iMissCount = $iMissCount + 1

		 ;Make room in the array for our new match
		 ReDim $rMiss[$iMissCount + 1][$iArray1Columns]

		 ;Add the miss to the rMiss array
		 For $iColumn = 0 to $iArray1Columns - 1
			$rMiss[$iMissCount][$iColumn] = $Array1[$iSearchRow][$iColumn]
		 Next

	  EndIf


   Next

   ;Global $$MatchArrayName = $rMatch
   Assign($MatchArrayName, $rMatch, 2)
   Assign($MissArrayName, $rMiss, 2)

   Return $iSearchRow

EndFunc;==>ArrayJoin

;Function:	AD_UserExists
;--
;Purpose:	Search the table of Active Directory Users and determine whether a username exists
;--
;Input:		$username:			String.  The username to check against the array of AD Users previously imported.
;
;			$AD_array:			Array.	The array of Active Directory users to check against.
;
;			$UsernameColumn:	Integer.  The column number where usernames are stored.  This will probably come from CSVHeader()
;---
;Output:	On match-found:		1
;
;			On match-not-found:	0
Func AD_UserExists($username, $AD_array, $UsernameColumn)

   ;Search the $AD_array starting at row 1 until the end, not case sensitive, with no change in data type, for the name $username within the $UsernameColumn
   $r = _ArraySearch($AD_Array, $username,1,0,0,0,1,$UsernameColumn)
   If $r > 0 Then
	  return 1
   EndIf

   Return 0  ;$r was less than 1.  Return 0.

EndFunc


; Func:		CSVFileToArray
; Purpose:	Given A CSV, generate a 2-dimensional array
; ---
; Input:	$sInputCSV: 	String.  A full path and filename to the CSV File
;
;			$sHeaderPrefix:	String. A prefix that will be applied to the header names to keep them unique
;							to avoid name collisions when comparing arrays.

Func CSVFileToArray($sInputCSVFile, $sHeaderPrefix = '')
   $rInputCSVFile = FileReadToArray($sInputCSVFile)
   $iInputCSVFileUBound = UBound($rInputCSVFile)	; Find the number of rows for the 2D array
   $rCSVLineUBound = UBound(CSVLineToArray($rInputCSVFile[$iInputCSVFileUBound - 1])) 	;Find the number of columns in the last row of the CSV file to determine our number of columns.
																					   ;(Using the last row prevents problems with headers and pre-file comments)

   Local $CSV2D[$iInputCSVFileUBound][$rCSVLineUBound] ;Declare our 2D array given the number of rows and columns we counted above
   Local $iOmittedLinesCount = 0

   ;_ArrayDisplay($rInputCSVFile)

   ;Step through each line of the input CSV File
   For $iLine = 0 to (Ubound($rInputCSVFile) - 1)

	  ;Read the next line of the input file
	  $rCSVLine = CSVLineToArray($rInputCSVFile[$iLine])

	  If IsArray($rCSVLine) Then
		 ;CSVLineToArray gave us a valid array.  Process it.
		 ;Step through each element of the line and assign to columns in the current row
		 For $iElement = 0 to (UBound($rCSVLine) - 1)
			If ($iLine - $iOmittedLinesCount = 0) Then
			   ;We are handling row zero.  This is the header row.  Apply the Column Name Prefixes to the values
			   $CSV2D[$iLine - $iOmittedLinesCount][$iElement] = $sHeaderPrefix & $rCSVLine[$iElement]
			Else
			   ;We are handling any other value.  Simply copy the value.
			   $CSV2D[$iLine - $iOmittedLinesCount][$iElement] = $rCSVLine[$iElement]
			EndIf

		 Next
	  Else
		 ;CSVLineToArray returned something that is not an array.
		 ;Omit this line.
		 ;Increment the OmittedLines counter so we can fill the array sequentially
		 $iOmittedLinesCount = $iOmittedLinesCount + 1

	  EndIf

   Next

   return $CSV2D
EndFunc


; Function: 	CSVLineToArray
;---
; Syntax: 		CSVLineToArray($sCSVInputLine)
;---
; Input:		$sCSVInputLine	- 	A single line from a CSV file, formatted as CSV.
;									Parsing is done with the RegEx search string (".*?",|.*?,|,)
;---
; Requirements:	Include Array.au3, String.au3
;---
;Returns		On Success:		Array of matches starting from zero.
;
;				On failure:		Integer.
;								0 = The line was commented out with a "#" character
Func CSVLineToArray($CSVInputLine)
   $string = $CSVInputLine
   ;=ERROR HANDLING===
   If StringLeft(StringStripWS($string, 3), 1) = '#' Then
	  ;This line has been commented out with a # character.  Return 0.
	  return 0
   EndIf

   ;=PROCESSING===
   ;Make sure the line ends in a comma.  This will allow our RegEx to get the last data field properly.
   If Not (StringRight(StringStripWS($string, 3), 1) = ',') Then
	 $string = $string & ","
	 ;MsgBox(0, 'Modified String:', $string)
   EndIf

   ;The magic: RegEx search to return the columns.
   $rCSVLine = StringRegExp($string, '(".*?",|.*?,|,)', 3)

   ;=CLEAN UP THE DATA====
   ;Remove the trailing comma left in the value of each of our CSV values as a result of our Regex.
   For $i = 0 To UBound($rCSVLine) - 1
	  If (StringRight($rCSVLine[$i], 1) = ',') Then
		 $rCSVLine[$i] = StringTrimRight($rCSVLine[$i], 1)
		 $rCSVLine[$i] = $rCSVLine[$i]
	  EndIf

	  ;If the field is enclosed in quotes, remove the quotes
	  If StringRight($rCSVLine[$i], 1) = '"' AND StringLeft($rCSVLine[$i], 1) = '"' Then
		 $rCSVLine[$i] = StringTrimRight($rCSVLine[$i], 1)
		 $rCSVLine[$i] = StringTrimLeft($rCSVLine[$i], 1)
	  EndIf

	  ;Strip leading and trailing whitespace from the line.
	  $rCSVLine[$i] = StringStripWS($rCSVLine[$i], 3)

   Next


   return $rCSVLine

EndFunc;==>CSVReadLine()


;Function:	CSVHeader
;---
;Purpose: 	When a CSV file is processed in this program we store the whole CSV file as an array
;			Row 0 of that array should always be the header names.
;			This function can be used to refer to the header columns by name
;---
;Input:		$rCSV			= Array.  The array containing the CSV data you will be referring to
;			$HeaderName		= String. The name of the header you're trying to refer to
;Returns:	On Success:	Column number of a given header
;			On Failure:	-1 (Column doesn't exist)
;---
;Requires:	Include Array.au3
Func CSVHeader($rCSVContents, $Headername)

   ;Search the array for the column number that corresponds to the headername
   $r = _ArraySearch($rCSVContents,$Headername,0,0,0,0,0,0,True)

   If $r < 0 Then
	  LogWrite('CSVHeader() was unable to find the header named "' & $Headername & '" .  Array search returned ' & $r & '.  This will likely cause unexepcted results.', 'FATAL ERROR')
   EndIf

   Return $r

EndFunc;==>CSVHeader


;Function:	GradYearToGradeLevel
;---
;Purpose:	Converts a student's Grad year to their current grade level
;---
;Input:		$sStudentGradYear: The student's current graduation year
;---
;Returns:	Integer.  	The current grade level.
;						0 indicates kindergarten
;						-1 indicates 4k
;						Other negative values may have custom meaning.
;						Check with the SIS secretaries for more details.
;--
;Requires:	The global $CurrentGradYear must be set
Func GradYearToGradeLevel($sStudentGradYear)
   $iGradeLevel = $CurrentGradYear + 12 - $sStudentGradYear
   Return $iGradeLevel
EndFunc;==>GradYearToGradeLevel


;------------------------------------------------------
; Function:	logwrite
; Purpose:	Write a log of program activity to disk
Func logwrite($log_message, $msgtype = 'INFO', $file = $sLogfilePath)
   $DateStamp = @Hour & ":" & @Min & ":" & @SEC & " " & @MON & "/" & @MDAY & "/" & @YEAR
   $line = $DateStamp & @TAB & $msgtype & @TAB & $log_message & @CRLF
   FileWriteLine($file, $DateStamp & @TAB & $msgtype & ':' & @TAB & $log_message)
   ConsoleWrite($line)
   ;MsgBox(0, '', $log_message)	;Uncomment to troubleshoot.
EndFunc



;Function:			FilePrepend
;--
;Purpose:			Prepend data to the beginning of a File
;--
;Requirements:		File.au3
;--
;Input:				$file	: String.  The path and name of the file to be modified
;					$data	: String.	The data to be prepended to the file
Func FilePrepend($file, $data)
   $originalcontents = FileRead($file)
   FileDelete($file)
   FileWrite($file, $data & $originalcontents)
   return 0
EndFunc


;Function:		PendingNewUser_Add
;---
;Purpose:		Keep track of what new usernames we have defined so we can check against this list for conflicts as we process
;				the full list of new users.
;---
;Input:			$username:			String.  The final new username that will be created.  It should already have
;									undergone sanity checks and conflict checks.
;---
;Returns:		Integer.  The number of usernames rows in the NewUsernames array after the latest addition
;---
;Requires:		1D array $PendingNewUsers must already have been defined by the main program which calls this function.
Func PendingNewUser_Add($username)
   LogWrite('Adding username to PendingNewUsers tracking array: ' & $username & '.')
   $iCount = UBound($PendingNewUsers)
   ReDim $PendingNewUsers[$iCount + 1]
   $PendingNewUsers[$iCount] = $username
   LogWrite('There are now ' & $iCount + 1 & ' New usernames being tracked pending creation.')
   return $iCount + 1
EndFunc

;Function:		PendingNewUser_Exists
;---
;Purpose:		Search whether a certain username exists in the global $PendingNewUsers tracking array
;---
;Input:			$Username: 		String. The username we will check against the $PendingNewUsers tracking array
;---
;Requires:		1D array $PendingNewUsers must already have been defined by the main program which calls this function.
;---
;Returns:		0 if $username does NOT exist in the $PendingNewUsers array
;				1 if $username DOES exist in the $PendingNewUsers array
Func PendingNewuser_Exists($username)
   LogWrite('Checking to see if ' & $username & ' exists in the PendingNewUsers tracking array.')
   $r = _ArraySearch($PendingNewUsers, $username)
   if $r > 0 Then
	  ;The username was found in the array
	  LogWrite($username & ' already exists in the PendingNewUsers tracking array.')
	  return 1
   Else
	  ;The username was not found in the array.
	  LogWrite($username & ' does not yet exist in the PendingNewUsers tracking array.')
	  return 0
   EndIf

EndFunc

;Function:		StringStripFileExt
;---
;Purpose:		Strips the file extension from a given filename
;---
;Input:			$filename:	String.  The name of a file that might have a file extension on the end.
;---
;Returns:		$filename without a file extension.
;				If $filename didn't have an "." in the first place, $filename is returned unmodified.
Func StringStripFileExt($filename)

   If Not IsString($filename) Then
	  Return 0
   EndIf

   $filesplit = StringSplit($filename, ".")
   If Not IsArray($filesplit) OR ($filesplit[0] = 1) Then
	  ;The string apparently didn't have a dot in it.  Just return the original.
	  return $filename
   EndIf

   local $sOut = ''	;initialize the variable where we'll build our output

   ;Split the filename by "."
   For $i = 1 to ($filesplit[0] - 1)

	  ;Reassemble the filename string
	  $sOut &= $fileSplit[$i]
	  if $i < $filesplit[0] - 1 Then
		 ;Add all but the last "." (so we have Ex.am.ple from Ex.am.ple.csv with no trailing dot.)
		 $sOut &= '.'
	  EndIf

   Next

   return $sOut

EndFunc









