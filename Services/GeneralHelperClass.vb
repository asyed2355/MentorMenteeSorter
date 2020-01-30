Option Explicit

'This class's function is to perform general functions that are called upon throughout the program.

'--------------------------------------------------
'FUNCTION DESCRIPTIONS:
	'--General Functions--'
		'getLength: return the length of an array.
		
		'getLengthLong: the same as GetLength, with a return type of Long rather than Integer.
		
		'sheetExists: Checks whether or not a given worksheet exists.
		
		'convertToLetter: Converts column number to column letter (e.g. 1 = "A", 2 = "B", 27 = "AA", etc.).
		
		'returnErrorMessage: Returns an error message as a string (this was implemented before the CustomError class was built).
		
		'createDuplicateStudentCollection: Creates a duplicate student collection. This is helpful when sorting printing students' information to worksheets, as a duplicate collection allows students to be deleted from the duplicate collection when they've been sorted without risking the deletion of any actual data - e.g. used when sorting mentees by stream.

	'--Range Functions--'			
		'findLastRow: Finds last row with data in a given worksheet. Used extensively whenever it is necessary to loop through a worksheet.
		
		'findFirstRowWithStudent: Searches through a given worksheet and returns the first row where a student is found.
		
	'--Data Validation Functions--'	
		'lettersFound: Determines if letters are found in an input string (useful for differenciating SID's and UniKeys, for instance).
		
		'SIDfound: Determines if the input is a valid SID.
		
		'unikeyFound: Determines if the input is a valid UniKey.
		
		'dataExistsInRange: Determines if any non-blank cells are found in a given range.
		
	'--Text Formatting Functions--'
		'removeSpacesFromSID: Function meant to specifically remove all characters from a string so that only numbers remain.
		
		'removeSpacesFromStartAndEnd: Remove all spaces from the start and the end of a string while leaving all other spaces intact (e.g. "  hello world " becomes "hello world"). Used when extracting data from mentor/mentee lists and converting to Mentor/Mentee objects to help mitigate typos.
		
	'--Misc--
		'assignGroupNumbersToMentors: Assigns a group identifier to all mentors (e.g. "Gr 1-11", "Gr 1-12", "Gr 3-38" etc.)
	
	'--Matching Function--
		'MatchMenteesWithMentors: Performs the actual matching of mentors with mentees. Four itterations are performed, with the criteria becoming more relaxed in each itteration. Any students who haven't been allocated by the end of the four itterations are added to the 'Unmatched Students' tab.
		
	'--Testing--
		'generateRandomNumber: Generates a random number within a given minimum and maximum range.
		
		'createDummyStudentCollection: Generates a collection of dummy students.
		
		'generateRandomStream: Generates a random stream. Used when assigning stream values to dummy students.
		
		'generateRandomMajor: Generates a random major. Used when assigning major values to dummy students.
		
		'generateRandomDegree: Generates a random degree. Used when assigning degree values to dummy students.	

'--------------------------------------------------

'-----GENERAL FUNCTIONS-----'

Public Function getLength(a As Variant) As Integer
   If IsEmpty(a) Then
      Let getLength = 0
   Else
      Let getLength = UBound(a) - LBound(a) + 1
   End If
End Function

Public Function getLengthLong(a As Variant) As Long
   If IsEmpty(a) Then
      getLengthLong = 0
   Else
      getLengthLong = UBound(a) - LBound(a) + 1
   End If
End Function

Public Function sheetExists(wb As Workbook, SheetName As String) As Boolean
    Dim sheet As Worksheet
    sheetExists = False
        For Each sheet In wb.Worksheets
            If SheetName = sheet.Name Then
                sheetExists = True
                Exit Function
            End If
        Next sheet
End Function

Function convertToLetter(iCol As Long) As String
   'Taken from this source: https://docs.microsoft.com/en-gb/office/troubleshoot/excel/convert-excel-column-numbers
   Dim iAlpha As Integer
   Dim iRemainder As Integer
   iAlpha = Int(iCol / 27)
   iRemainder = iCol - (iAlpha * 26)
   If iAlpha > 0 Then
      convertToLetter = Chr(iAlpha + 64)
   End If
   If iRemainder > 0 Then
      convertToLetter = convertToLetter & Chr(iRemainder + 64)
   End If
End Function

Public Function returnErrorMessage(Optional ErrorType As Byte) As String
	'Error codes:
	'1 = SID not found
	'2 = SID and UniKey not found
	'3 = Course not found (couldn't match it to a group)

	If IsMissing(ErrorType) Then
		'Return generic error message
		returnErrorMessage = "An error occured."
	Else
		Select Case ErrorType
			Case 1
			returnErrorMessage = "Valid Student ID not found."
			Case 2
			returnErrorMessage = "Valid Student ID or UniKey not found."
			Case 3
			returnErrorMessage = "Course not recognised (group not found)."
			Case Else
			returnErrorMessage = "Error occured."
		End Select
	End If
End Function

Public Function createDuplicateStudentCollection( _
ByRef StudentCollection As Collection, _
Optional studentType As Byte) _
As Collection

	'studentType (1) = Mentee
	'studentType (2) = Mentor

	Dim newCollection As Collection: Set newCollection = New Collection

	If StudentCollection.Count > 0 Then
	   Dim newStudent As Object
	   
	   Select Case studentType
			Case 1
				Set newStudent = New Mentee
			Case 2
				Set newStudent = New Mentor
			Case Else
				Set newStudent = New Mentee
		End Select

		Dim i As Long
		For i = 1 To StudentCollection.Count
			Set newStudent = StudentCollection(i)
			newCollection.Add StudentCollection(i)
			Set newStudent = Nothing
		Next i
	End If

	Set createDuplicateStudentCollection = newCollection
End Function

'-----RANGE FUNCTIONS-----'

Public Function findLastRow(wb As Workbook, ByRef WorksheetName As String) As Long
    Dim helper As New GeneralHelperClass
    If IsMissing(WorksheetName) Or Not (helper.sheetExists(wb, WorksheetName)) Then
        findLastRow = 1
        Exit Function
    Else
        Dim sheet As Worksheet: Set sheet = wb.Sheets(WorksheetName)
        findLastRow = sheet.UsedRange.Rows(sheet.UsedRange.Rows.Count).row
    End If
End Function

Public Function findFirstRowWithStudent( _
ByRef wb As Workbook, _
ByRef WorksheetName As String, _
SID_ColumnNumber As Integer, _
Optional ByVal lastRow As Integer, _
Optional StartingRow As Long) As Long

    Dim helper As New GeneralHelperClass
    Dim studentRowNumber As Long: studentRowNumber = 1
    
    If IsMissing(lastRow) Then
        lastRow = findLastRow(wb, WorksheetName)
    End If
    
    'Determine the first row to start counting from
    Dim firstRow As Long: firstRow = 1
    If Not (IsMissing(StartingRow)) Then
        If StartingRow < 1 Then
            firstRow = 2
        Else
            firstRow = StartingRow
        End If
    End If

    wb.Sheets(WorksheetName).Activate
    Dim row As Range
    Dim i As Long
    For i = firstRow To lastRow
      Set row = Range("'" & WorksheetName & "'!A" & i & ":" & "P" & i)
        If (helper.SIDfound(row.Cells(1, SID_ColumnNumber)) Or helper.unikeyFound(row.Cells(1, SID_ColumnNumber))) Then
            studentRowNumber = i
            Exit For
        End If
    Next i  
    findFirstRowWithStudent = studentRowNumber
End Function

'-----DATA VALIDATION FUNCTIONS-----'

Public Function lettersFound(Value As String) As Boolean
    'This function checks the ASCII values of each char in Value and returns TRUE if a letter is found (a-z, A-Z)
    'Note: Used primarily to check if streams are valid.
    Dim found As Boolean: Let found = False
    
    If Value <> "" And Value <> " " And Value <> "-" And Value <> " - " Then
        Dim i As Long
        Dim ASCII_Value As String
        
        For i = 1 To Len(Value)
            Let ASCII_Value = Asc(Mid(Value, i, 1))
            If (ASCII_Value >= 65 And ASCII_Value <= 90) Or (ASCII_Value >= 97 And ASCII_Value <= 122) Then
                found = True
                Exit For
            End If
        Next i
    End If
    
    Let lettersFound = found
End Function

Public Function SIDfound(TextToCheck As String) As Boolean
'SID in this context = 123456789 (9 numbers and 0 letters)
	
	'A preliminary check to see if a positive result can be ruled out from the get go
	If IsMissing(TextToCheck) Or Not(IsNumeric(TextToCheck)) Or TextToCheck = "" Or Len(TextToCheck) <> 9 Then
		SIDfound = False
	Else
		Dim sidCharsCorrect As Boolean: sidCharsCorrect = True
		Dim i As Integer
		Dim ASCII_Value As Integer
		'Check the ASCII characters of the TextToCheck to check that they're all numbers
		For i = 1 To 9
			ASCII_Value = Asc(Mid(TextToCheck, i, 1))
			If Not (ASCII_Value >= 48 And ASCII_Value <= 57) Then
				sidCharsCorrect = False
				Exit For
			End If
		Next i
		SIDfound = sidCharsCorrect
	End If
End Function

Public Function unikeyFound(TextToCheck As String) As Boolean
'Unikey in this context = abcd1234 (4 letters and 4 numbers, in that order)

'A preliminary check to see if a positive result can be ruled out from the get go
If IsMissing(TextToCheck) Or TextToCheck = "" Or Len(TextToCheck) < 8 Then
    unikeyFound = False
Else
    Dim unikeyCharsCorrect As Boolean: unikeyCharsCorrect = True
    Dim i As Integer
    
    'Check the first 4 characters of TextToCheck to see if they're all letters
    'This is being done by checking each characters ASCII value.
    For i = 1 To 4
        Dim ASCII_Value As String: ASCII_Value = Asc(Mid(TextToCheck, i, 1))
        If Not ((ASCII_Value >= 65 And ASCII_Value <= 90) Or (ASCII_Value >= 97 And ASCII_Value <= 122)) Then
            unikeyCharsCorrect = False
            Exit For
        End If
    Next i
End If
    'If the last check failed then there is no need to continue.
    If Not (unikeyCharsCorrect) Then
        unikeyFound = unikeyCharsCorrect
    Else
        'Check the next 4 characters of TextToCheck to see if they're all numbers
        'If a non-number is found, stop the loop and set the return value of the function to False
        For i = 5 To 8
            ASCII_Value = Asc(Mid(TextToCheck, i, 1))
            If Not (ASCII_Value >= 48 And ASCII_Value <= 57) Then
                unikeyCharsCorrect = False
                Exit For
            End If
        Next i
        unikeyFound = unikeyCharsCorrect
    End If
End Function

Public Function dataExistsInRange(R As Range) As Boolean
    Dim lastColumnNumber As Integer
    Let lastColumnNumber = R.Cells(1, R.Columns.Count).End(xlToLeft).Column
    Dim lastRowNumber As Long
    Let lastRowNumber = R.Cells(R.Rows.Count, 1).End(xlUp).row
    
    Dim dataFound As Boolean: Let dataFound = False
    
    Dim i As Long
    Dim j As Integer
    
    For i = 1 To lastRowNumber
        For j = 1 To lastColumnNumber
            If Len(R.Cells(i, j)) <> 0 Then
                If R.Cells(i, j).Value <> "" And R.Cells(i, j).Value <> 0 And R.Cells(i, j).Value <> " " Then
                    dataFound = True
                End If
                Exit For
            End If
        Next j
    Next i
    dataExistsInRange = dataFound
End Function

'-----TEXT FORMATTING FUNCTIONS-----'

Public Function removeSpacesFromSID(Text As String) As String
	If IsMissing(Text) Or Text = "" Or Len(Text) < 9 Then
		removeSpacesFromSID = Text
	Else
		Dim result As String
		Let result = ""
		Dim i As Integer
		Dim ASCII_Value As Integer
		'Check the ASCII characters of the TextToCheck to check that they're all numbers
		For i = 1 To Len(Text)
			ASCII_Value = Asc(Mid(Text, i, 1))
			If ASCII_Value >= 48 And ASCII_Value <= 57 Then
				result = result & Mid(Text, i, 1)
			End If
		Next i
		removeSpacesFromSID = result
	End If
End Function

Public Function removeSpacesFromStartAndEnd(Text As String) As String
'The ASCII code for a space is '32'.
	If Text = "" Or Len(Text) = 0 Then
		removeSpacesFromStartAndEnd = ""
	Else
		Dim i As Integer
		Dim startChar As Integer
		Dim endChar As Integer
		
		'This Loop deals with the start of Text
		For i = 1 To Len(Text)
			If Asc(Mid(Text, i, 1)) <> 32 Then
				startChar = i
				Exit For
			End If
		Next i
		
		For i = Len(Text) To startChar Step -1
			If Asc(Mid(Text, i, 1)) <> 32 Then
				   endChar = i
				   Exit For
			End If
		Next i
		
		If startChar < endChar Then
			removeSpacesFromStartAndEnd = Mid(Text, startChar, Len(Text) - (Len(Text) - endChar))
		Else
			removeSpacesFromStartAndEnd = ""
		End If
	End If
End Function

'-----MISC-----'

Public Sub assignGroupNumbersToMentors( _
ByRef mentorCollection As Collection, _
ByRef GroupNames As Variant)

Dim groupCount As Long
Dim i As Integer
Dim j As Long

For i = getLength(GroupNames) - 1 To 0 Step -1
    Let groupCount = 1
    For j = mentorCollection.Count To 1 Step -1
        If mentorCollection(j).course.Group = GroupNames(i) Then
            mentorCollection(j).groupNumber = groupCount
            Let groupCount = groupCount + 1
        End If
    Next j
Next i

End Sub


'-----MATCHING FUNCTION-----'

Public Sub MatchMenteesWithMentors( _
ByRef menteeCollection, _
ByRef mentorCollection, _
MaxAllocations As Byte)

	Dim tempStudentCollection As Collection
	Dim matchFound As Boolean
	Dim majorFound As Boolean
	Dim currentMajor As String

	Dim mentorMajors(2) As String
	Dim menteeMajors(2) As String

	Dim i As Byte
	Dim j As Long
	Dim k As Long
	Dim l As Byte
	Dim m As Byte


	'--- Itterations (i) ---'

	'First Itteration (i = 1)'
		'First major, AND
		'Second major, AND
		'Stream, AND
		'Course
		
	'Second Itteration (i = 2)'
		'First major OR Second major, AND
		'Stream, AND
		'Course
		
		
	'Third Itteration (i = 3)'
		'Stream, AND
		'Course
		
	'Fourth Itteration (i = 4)'
		'Stream, AND
		'Group
		
	'Note: I've stacked IF statements on purpose;
	'the most unlikely statements to be true are the outtermost ones -
	'i.e. Major2 is very unlikely to be a match. Major1 is relatively more likely.
	'Course is more likely than majors and stream is much more likely than all of them.
	'I've opted for this method rather than one long line of AND statements so that non-matches can
	'be ruled out quickly with minimal calculations needing to take place.
	
	'Also of note - spaces are removed from major names and all characters are converted to upper case
	'to help deal with potential typos.

	For i = 1 To 4
		For j = menteeCollection.Count To 1 Step -1

			'Add current mentee's majors to menteeMajor array
			If menteeCollection(j).Major1 = Null Or menteeCollection(j).Major1 = "" Then
				Let menteeMajors(0) = ""
			Else
				Let menteeMajors(0) = Replace(UCase(menteeCollection(j).Major1), " ", "")
			End If
			
			If menteeCollection(j).Major2 = Null Or menteeCollection(j).Major2 = "" Then
				Let menteeMajors(1) = ""
			Else
				Let menteeMajors(1) = Replace(UCase(menteeCollection(j).Major2), " ", "")
			End If
			
			For k = 1 To mentorCollection.Count
				'If mentor is at capacity, move on to the next one
				If Not (mentorCollection(k).AtCapacity) Then
				
					'Add current mentee's majors to menteeMajor array
					If mentorCollection(k).Major1 = Null Or mentorCollection(k).Major1 = "" Then
						Let mentorMajors(0) = ""
					Else
						Let mentorMajors(0) = Replace(UCase(mentorCollection(k).Major1), " ", "")
					End If
					
					If mentorCollection(k).Major2 = Null Or mentorCollection(k).Major2 = "" Then
						Let mentorMajors(1) = ""
					Else
						Let mentorMajors(1) = Replace(UCase(mentorCollection(k).Major2), " ", "")
					End If
					
					'--- Switch statement to determine match based on itteration number (start)---'
					Let matchFound = False
					Select Case i
						Case 1
						'Itteration 1
						If menteeMajors(1) = mentorMajors(1) Then
							If menteeMajors(0) = mentorMajors(0) Then
								If UCase(menteeCollection(j).course.CourseName) = UCase(mentorCollection(k).course.CourseName) Then
									If menteeCollection(j).stream.FullName = mentorCollection(k).stream.FullName Then
										matchFound = True
									End If
								End If
							End If
						End If
						Case 2
						'Itteration 2
						'Nested loop to check whether or not major1 or major2 are a match
						Let majorFound = False
						For l = 0 To 1
							If mentorMajors(l) <> "" Then
								For m = 0 To 1
									If menteeMajors(m) <> "" Then
										If mentorMajors(l) = menteeMajors(m) Then
											Let majorFound = True
											Exit For
										End If
									End If
								Next m
							End If
						Next l
						
						If majorFound Then
							If UCase(menteeCollection(j).course.CourseName) = UCase(mentorCollection(k).course.CourseName) Then
								If menteeCollection(j).stream.FullName = mentorCollection(k).stream.FullName Then
									matchFound = True
								End If
							End If
						End If
						Case 3
						'Itteration 3
						If UCase(menteeCollection(j).course.CourseName) = UCase(mentorCollection(k).course.CourseName) Then
							If menteeCollection(j).stream.FullName = mentorCollection(k).stream.FullName Then
								matchFound = True
							End If
						End If
						Case 4
						'Itteration 4
						If UCase(menteeCollection(j).Group) = UCase(mentorCollection(k).Group) Then
							If menteeCollection(j).stream.FullName = mentorCollection(k).stream.FullName Then
								matchFound = True
							End If
						End If
					End Select
					'--- Switch statement to determine match based on itteration number (end)---'
					
					If matchFound Then
						'Extract mentor's student collection (to tempCollection),
						'add mentee to temp collection,
						'Override mentor's student collection with tempCollection,
						'Empty tempCollection.
						
						Set tempStudentCollection = mentorCollection(k).Students
						Let menteeCollection(j).MatchType = i
						tempStudentCollection.Add menteeCollection(j)
						mentorCollection(k).Students = tempStudentCollection
						Let mentorCollection(k).studentCount = mentorCollection(k).studentCount + 1
						Set tempStudentCollection = Nothing
						
						'If the mentor's collection max has been reached
						If mentorCollection(k).Students.Count >= MaxAllocations Then
							mentorCollection(k).AtCapacity = True
						End If
						
						'Remove matched mentee from mentee collection
						menteeCollection.Remove (j)
						Exit For
					End If
				End If
			Next k
		Next j
	Next i

	'By the end of this sub, the only students left in the mentee collection should be the ones unmatched in all itterations.
	'These can then be matched manually.
End Sub

'-----TESTING FUNCTIONS-----'

Public Function generateRandomNumber(min As Integer, max As Integer) As Long
    If min < max Then
        generateRandomNumber = Int((max - min + 1) * Rnd() + min)
    Else
        generateRandomNumber = 1
    End If
End Function

Public Function createDummyStudentCollection(collectionSize As Long, studentType As Byte) As Collection
    
    Dim m As Object
    Dim c As Collection: Set c = New Collection

    Dim i As Long
        
    For i = collectionSize To 1 Step -1
        Select Case studentType
            '1: Mentee
            Case 1
                Set m = New Mentee
            '2: Mentor
            Case 2
                Set m = New Mentor
            'Option default: Mentee
            Case Else
                Set m = New Mentee
        End Select
    
        With m
            .studentID = 440440440
            .FirstName = "Dummy"
            .LastName = "Student"
            .Email = "dummy.student@sydney.edu.au"
            '.Mobile = "0411111111"
            .stream = generateRandomStream()
            .course = generateRandomDegree()
            .Major1 = generateRandomMajor()
            .Major2 = generateRandomMajor()
        End With
        c.Add m
    Next i
    
    Set createDummyStudentCollection = c
End Function

Public Function generateRandomStream() As stream
    Dim stream As stream: Set stream = New stream
    Dim n As Integer: Let n = generateRandomNumber(1, 4)
        
        Select Case n
            '1: DL
            Case 1
                stream.FullName = "Dalyell"
                stream.ShortName = "DL"
            '2: Int
            Case 2
                stream.FullName = "International"
                stream.ShortName = "Int"
            '3: 25+
            Case 3
                stream.FullName = "25+"
            stream.ShortName = "25+"
            '4: Dom
            Case 4
                stream.FullName = "Domestic"
            stream.ShortName = "Dom"
            'Option default: Dom
            Case Else
                stream.FullName = "Domestic"
                stream.ShortName = "Dom"
        End Select
        
    Set generateRandomStream = stream
End Function

Public Function generateRandomMajor() As String
    Dim result As String: Let result = ""
    Dim n As Integer: Let n = generateRandomNumber(1, 5)
        
        Select Case n
            Case 1
                result = "Economics"
            Case 2
                result = "Sociology"
            Case 3
                result = "Political Economy"
            Case 4
                result = "Gender Studies"
            Case 5
                result = "Modern Greek Studies"
            Case Else
                result = "Art History"
        End Select
    Let generateRandomMajor = result
End Function

Public Function generateRandomDegree() As course
    Dim c As course: Set c = New course
    Dim n As Integer: Let n = generateRandomNumber(1, 5)
        
        Select Case n
            Case 1
                c.CourseName = "Bachelor of Arts"
                c.Group = "1"
            Case 2
                c.CourseName = "Bachelor of Arts/Bachelor of Advanced Studies"
                c.Group = "1"
            Case 3
                c.CourseName = "Bachelor of Arts and Bachelor of Laws"
                c.Group = "3"
            Case 4
                c.CourseName = "Bachelor of Economics"
                c.Group = "3"
            Case 5
                c.CourseName = "Bachelor of Education"
                c.Group = "2"
            Case Else
                c.CourseName = "Bachelor of Arts and Doctor of Medicine"
                c.Group = "1"
        End Select
    Set generateRandomDegree = c
End Function
