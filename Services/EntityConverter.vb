Option Explicit

'This class reads data from the relevant data sources and converts rows of data to instances of classes.

Public Sub createStudentCollections( _
ByRef Workbook As Workbook, _
ByRef MenteeWorksheet As Worksheet, _
ByRef MentorWorksheet As Worksheet, _
ByRef CourseCollection As Collection, _
ByRef menteeCollection As Collection, _
ByRef mentorCollection As Collection, _
ByRef StudentsWithErrorsCollection As Collection)
    
    ' **Probably the most important sub in the entire project** '
    
    'This sub loops through the mentee and mentor lists and save them as mentee/mentor objects.
    'In other words, this sub extracts data from the mentor/mentee list and converts it to objects/classes.
    'Converting students into classes/objects allows all other steps in the project to take place.
    
    'Note: As VBA doesn't allow for class inheritance, mentors and mentees don't share a parent class.
    'To work with this, I have defined the variable 'm' as a generic object that will either be
    'cast into a mentee/mentor, depending on which is needed at any given time.
  
    
    '---Global Objects---'
    Dim m As Object
    Dim Worksheet As Worksheet
    Dim course As course
    Dim stream As stream
    Dim err As CustomError: Set err = New CustomError
    Dim helper As GeneralHelperClass: Set helper = New GeneralHelperClass
    
    '---Range-related variables---'
    Dim lastRow As Long
    Dim lastColumn As String
    Dim lastColumnNumber As Long
    Dim SIDcolumn As String
    Dim major1Column As Long
    Dim major2Column As Long
    Dim tabRow As Range

    '---Error checking variables---'
    Dim SID_Found As Boolean
    Dim UniKey_Found As Boolean
    Dim moveStatus As String
    Dim errorFound As Boolean
    Dim cellWithDataFound As Boolean
    Dim errorNumber As Long
    
    '---Data validation variables---'
    Dim studentID As Long
    Dim nextStudentID As String
    Dim uniKey As String
    Dim addStudentToCollection As Boolean
    
    '---Mentor-specific variables---'
    'Group count ("Gr 1 - 1", "Gr 3 - 10" etc.)
    Dim group1MentorCount As Long
    Dim group2MentorCount As Long
    Dim group3MentorCount As Long
    
    '---Loop variables---'
    Dim h As Byte
    Dim i As Long
    'Dim j As Byte
    
    'Loop h*:
    'h = 1 (mentees spreadsheet)
    'h = 2 (mentors spreadsheet)
        'Loop i:
        'Looping through each row in the spreadsheet and converting to a mentee/mentor object
    '*h was a later addition, which explains the unorthadox combo of 'h' and 'i' (instead of 'i' and 'j').
        
    For h = 1 To 2
        Let group1MentorCount = 0
        Let group2MentorCount = 0
        Let group3MentorCount = 0
        
        'I've set this up in such a way in case anything ever differs between mentors and mentees.
        Select Case h
            'h 1: Mentee
            Case 1
                Set m = New Mentee
                Set Worksheet = MenteeWorksheet
                Let SIDcolumn = "A"
                Let lastColumn = "L"
                Let major1Column = 8
                Let major2Column = 9
            'h 2: Mentor
            Case 2
                Set m = New Mentor
                Set Worksheet = MentorWorksheet
                Let SIDcolumn = "A"
                Let lastColumn = "L"
                Let major1Column = 6
                Let major2Column = 7
            'Option default: Mentee
            Case Else
                Set m = New Mentee
                Set Worksheet = MenteeWorksheet
                Let SIDcolumn = "A"
                Let lastColumn = "L"
                Let major1Column = 8
                Let major2Column = 9
        End Select
        
        Let lastRow = helper.findLastRow(Workbook, Worksheet.Name)
        Let lastColumnNumber = Range(lastColumn & 1).Column
        
        'The below function will first sort the relevant worksheet by SID.
        'As the function cycles through students, it will look
        'at the next student's SID and check whether or not there's a match
        'with the current student's SID. As the worksheet will be sorted,
        'duplicates should bundle together.
        'This is the quickest method I can think of in terms of performance to ensure duplicates are found.
        
        '--- Sort worksheet by SID ---'
        Workbook.Sheets(Worksheet.Name).Range("A1:" & lastColumn & lastRow).Sort _
        Key1:=Workbook.Sheets(Worksheet.Name).Range("A1:A" & lastRow), _
        Header:=xlYes
               
        For i = 2 To lastRow
            Set tabRow = Worksheet.Range("A" & i & ":" & lastColumn & i)
            Set course = New course
            Set stream = New stream
            Let moveStatus = "Y"
            Let addStudentToCollection = True
            Let errorNumber = -1
            
            Select Case h
                Case 1
                    Set m = New Mentee
                    course.CourseName = helper.removeSpacesFromStartAndEnd(tabRow.Cells(1, 7).Value)
                Case 2
                    Set m = New Mentor
                    course.CourseName = helper.removeSpacesFromStartAndEnd(tabRow.Cells(1, 5).Value)
                Case Else
                    Set m = New Mentee
                    course.CourseName = helper.removeSpacesFromStartAndEnd(tabRow.Cells(1, 7).Value)
            End Select
           
            '---DATA VALIDATION---'
            
            '---Valid SID/UniKey check---'
            If addStudentToCollection Then
                Let SID_Found = helper.SIDfound(Worksheet.Range(SIDcolumn & i).Value)
                If SID_Found Then
                    Let m.studentID = CLng(tabRow.Cells(1, 1).Value)
                    Let m.uniKey = ""
                ElseIf helper.SIDfound(helper.removeSpacesFromSID(CStr(Worksheet.Range(SIDcolumn & i).Value))) Then
                    Let m.studentID = CLng(helper.removeSpacesFromSID(Worksheet.Range(SIDcolumn & i).Value))
                    Let m.uniKey = ""
                ElseIf helper.unikeyFound(Worksheet.Range(SIDcolumn & i).Value) Then
                    'If a valid SID wasn't found, look for a UniKey.
                    Let m.studentID = 1
                    Let m.uniKey = tabRow.Cells(1, 1).Value
                    Let moveStatus = "Y - " & helper.returnErrorMessage(1)
                    Let errorNumber = 0
                Else
                    'If a valid UniKey wasn't found, check to see if the row has any data in it.
                    If helper.dataExistsInRange(tabRow) Then
                        Let moveStatus = "Y - " & helper.returnErrorMessage(2)
                        Let errorNumber = 1
                    Else
                        Let addStudentToCollection = False
                        Let moveStatus = "N"
                    End If
                End If
             End If
             
            '---Valid Major/s found---'
            If helper.lettersFound(tabRow.Cells(1, major1Column).Value) Then
                m.Major1 = helper.removeSpacesFromStartAndEnd(tabRow.Cells(1, major1Column).Value)
            End If
            
            If helper.lettersFound(tabRow.Cells(1, major2Column).Value) Then
                m.Major2 = helper.removeSpacesFromStartAndEnd(tabRow.Cells(1, major2Column).Value)
            End If
            
            '---CREATE MENTEE/MENTOR OBJECT---'
            
            '---If addStudentToCollection is still true, continue. Else i++
            If Not (addStudentToCollection) Then
                Set m = Nothing
                moveStatus = ""
            Else
                
                Select Case h
                    Case 1
                        With m
                            .FirstName = tabRow.Cells(1, 2).Value
                            .LastName = tabRow.Cells(1, 3).Value
                            .International_YorN = tabRow.Cells(1, 4).Value
                            .Dalyell_YorN = tabRow.Cells(1, 5).Value
                            .MatureAge_YorN = tabRow.Cells(1, 6).Value
                            .Email = tabRow.Cells(1, 10).Value
                            .Dietary = tabRow.Cells(1, 11).Value
                            .Mobility = tabRow.Cells(1, 12).Value
                        End With
                    Case 2
                        With m
                            .FirstName = tabRow.Cells(1, 3).Value
                            .LastName = tabRow.Cells(1, 4).Value
                            .International_YorN = tabRow.Cells(1, 10).Value
                            .Dalyell_YorN = tabRow.Cells(1, 11).Value
                            .MatureAge_YorN = tabRow.Cells(1, 12).Value
                            .Email = tabRow.Cells(1, 8).Value
                            .Students = New Collection
                        End With
                    Case Else
                        With m
                            .FirstName = tabRow.Cells(1, 2).Value
                            .LastName = tabRow.Cells(1, 3).Value
                            .International_YorN = tabRow.Cells(1, 4).Value
                            .Dalyell_YorN = tabRow.Cells(1, 5).Value
                            .MatureAge_YorN = tabRow.Cells(1, 6).Value
                            .Email = tabRow.Cells(1, 10).Value
                            .Dietary = tabRow.Cells(1, 11).Value
                            .Mobility = tabRow.Cells(1, 12).Value
                        End With
                End Select
                
                '---ERROR CHECKING---'
                'Errors are checked in order of least problematic to most.
                'If an error is found, but then down the line a more severe error is found,
                'the initial finding is overridden with the subsequent one.
                
                '---Checking valid stream---'
                
                If helper.lettersFound(m.International_YorN) Or _
                   helper.lettersFound(m.Dalyell_YorN) Or _
                   helper.lettersFound(m.MatureAge_YorN) Then
                
                    Set stream = stream.CreateStream(m.International_YorN, m.Dalyell_YorN, m.MatureAge_YorN)
                
                Else
                    stream.FullName = ""
                    stream.ShortName = ""
                    Let errorNumber = 3
                    m.Moved = False
                End If
                
                '---Checking that course is recognised---'
                
                course.Group = course.DetermineCourseGroup(course.CourseName, CourseCollection)
                If course.Group = "Course Not Found" Then
                    Let moveStatus = "N - " & helper.returnErrorMessage(3)
                    Let errorNumber = 4
                    m.Moved = False
                End If
                
                Let nextStudentID = Worksheet.Range(SIDcolumn & (i + 1)).Value
                
                '---Check for duplicate---'
                If CStr(m.studentID) = nextStudentID Or m.uniKey = nextStudentID Then
                    Let moveStatus = "N (Duplicate)"
                    Let errorNumber = 2
                    m.Moved = False
                End If
                
                m.course = course
                m.stream = stream
                
                '---ADDING ERROR OBJECT TO STUDENT---
                
                If errorNumber >= 0 Then
                    Set err = err.createCustomError(errorNumber, Worksheet.Name, i)
                    m.Error = err
                End If
                
                '---ADDING STUDENT TO APPROPRIATE COLLECTION---'
                
                If m.Moved Then
                    Select Case h
                        Case 1
                            menteeCollection.Add Item:=m
                        Case 2
                            Select Case m.course.Group
                                Case "1"
                                    Let group1MentorCount = group1MentorCount + 1
                                    m.groupNumber = group1MentorCount
                                Case "2"
                                    Let group2MentorCount = group2MentorCount + 1
                                    m.groupNumber = group2MentorCount
                                Case "3"
                                    Let group3MentorCount = group3MentorCount + 1
                                    m.groupNumber = group3MentorCount
                                Case Else
                                    m.groupNumber = 0
                            End Select
                            mentorCollection.Add Item:=m
                        Case Else
                            menteeCollection.Add Item:=m
                    End Select
                Else
                    StudentsWithErrorsCollection.Add Item:=m
                End If
            End If
            
            'Indicate on spreadsheet that they've been moved.
            tabRow.Cells(1, (lastColumnNumber + 1)).Value = moveStatus
            If addStudentToCollection Then
                tabRow.Cells(1, (lastColumnNumber + 2)).Value = stream.FullName
                tabRow.Cells(1, (lastColumnNumber + 3)).Value = course.Group
            End If
        Next i
    Next h
End Sub

Public Function createCourseCollection(wb As Workbook, TabName As String) As Collection
    Dim CourseCollection As Collection: Set CourseCollection = New Collection
    Dim DegreesByCategoryTab As Worksheet: Set DegreesByCategoryTab = wb.Worksheets(TabName)
    Dim lastRow As Long: Let lastRow = Application.WorksheetFunction.CountIf( _
    wb.Worksheets(DegreesByCategoryTab.Name).Range("A:A"), "<>" & "")
    Dim course As course
    
    Dim i As Long
    For i = 2 To lastRow
        If Not (IsMissing(DegreesByCategoryTab.Range("A" & i))) Then
            Set course = New course
            With course
                .CourseName = DegreesByCategoryTab.Range("A" & i).Value
                .Group = DegreesByCategoryTab.Range("B" & i).Value
            End With
            CourseCollection.Add Item:=course
            Set course = New course
        End If
    Next i
    
    Set createCourseCollection = CourseCollection
End Function
