'Named "TabClass" in final project.

Option Explicit

Public Sub createTab( _
Workbook As Workbook, _
Name As String, _
R As Byte, _
G As Byte, _
B As Byte, _
Optional TabAfter As String, _
Optional AutoFit As Boolean)
    
    Dim helper As GeneralHelperClass: Set helper = New GeneralHelperClass
    If Not (helper.sheetExists(Workbook, Name)) Then
        Dim newTab As Worksheet

        If IsMissing(TabAfter) Or Not (helper.sheetExists(Workbook, TabAfter)) Or TabAfter = "" Then
            Set newTab = Workbook.Sheets.Add(Type:=xlWorksheet, after:=Workbook.Sheets(Sheets.Count))
        Else
            Set newTab = Workbook.Sheets.Add(Type:=xlWorksheet, after:=Workbook.Worksheets(TabAfter))
        End If
        With newTab
            .Name = Name
            .Tab.Color = RGB(R, G, B)
        End With
        
        If AutoFit Or IsMissing(AutoFit) Then
            newTab.Cells.EntireColumn.AutoFit
        End If
    End If
    
End Sub

Public Sub CreateStreamTab( _
TabName As String, _
Workbook As Workbook, _
Optional TabAfter As String)

    If IsMissing(TabAfter) Then
        Call createTab(Workbook, TabName, 128, 0, 128, , True)
    Else
        Call createTab(Workbook, TabName, 128, 0, 128, TabAfter, True)
    End If
        
    Call AddHeaders(TabName, Workbook, 1)

End Sub

Public Sub CreateErrorsTab( _
TabName As String, _
Workbook As Workbook, _
Optional TabAfter As String)

    If IsMissing(TabAfter) Then
        Call createTab(Workbook, TabName, 255, 0, 0, , True)
    Else
        Call createTab(Workbook, TabName, 255, 0, 0, TabAfter, True)
    End If
        
    Call AddHeaders(TabName, Workbook, 1)

End Sub

'Public Sub CreateReportingTab( _
'TabName As String, _
'Workbook As Workbook)
'
'    If Workbook.Sheets.Count < 1 Then
'        Call createTab(Workbook, TabName, 255, 255, 255, , True)
'    Else
'        Call createTab(Workbook, TabName, 255, 255, 255, TabAfter, True)
'    End If
'
'    Call AddHeaders(TabName, Workbook, 1)
'
'End Sub

Public Sub AddUnmatchedMenteesTab( _
ByRef Workbook As Workbook, _
ByRef menteeCollection As Collection, _
Optional TabAfter As String)

    If IsMissing(TabAfter) Then
        Call createTab(Workbook, "Unmatched Mentees", 255, 0, 0, , True)
    Else
        Call createTab(Workbook, "Unmatched Mentees", 255, 0, 0, TabAfter, True)
    End If
    
    Call AddHeaders("Unmatched Mentees", Workbook, 3)
    
    Dim i As Long
    Dim RowNumber As Long: Let RowNumber = 2
    For i = 1 To menteeCollection.Count
        With Workbook.Sheets("Unmatched Mentees")
            'Add SID/Unikey
            If menteeCollection(i).studentID <= 1 Then
               .Range("C" & RowNumber).Value = menteeCollection(i).uniKey
            Else
                .Range("C" & RowNumber).Value = menteeCollection(i).studentID
            End If
            
            .Range("B" & RowNumber).Value = menteeCollection(i).stream.FullName
            .Range("D" & RowNumber).Value = menteeCollection(i).FirstName
            .Range("E" & RowNumber).Value = menteeCollection(i).LastName
            .Range("F" & RowNumber).Value = menteeCollection(i).course.CourseName
            .Range("G" & RowNumber).Value = menteeCollection(i).Major1
            .Range("H" & RowNumber).Value = menteeCollection(i).Major2
            .Range("I" & RowNumber).Value = menteeCollection(i).Email
            .Range("J" & RowNumber).Value = menteeCollection(i).Mobility
        End With
        Let RowNumber = RowNumber + 1
    Next i
    Workbook.Sheets("Unmatched Mentees").Cells.EntireColumn.AutoFit

End Sub

Public Sub CreateDuplicatesTab( _
Workbook As Workbook, _
Optional TabAfter As String)

    If IsMissing(TabAfter) Then
        Call createTab(Workbook, "Duplicates", 255, 201, 14, , True)
    Else
        Call createTab(Workbook, "Duplicates", 255, 201, 14, TabAfter, True)
    End If
    
    Call AddHeaders("Duplicates", Workbook, 1)

End Sub

Public Sub CreateGroupTab( _
TabName As String, _
GroupName As String, _
Workbook As Workbook, _
Optional TabAfter As String)

    If IsMissing(TabAfter) Then
        Let TabAfter = ""
    End If
        
    Select Case GroupName
        Case "1"
            Call createTab(Workbook, TabName, 142, 169, 219, TabAfter, True)
        Case "2"
            Call createTab(Workbook, TabName, 169, 208, 142, TabAfter, True)
        Case "3"
            Call createTab(Workbook, TabName, 255, 0, 0, TabAfter, True)
        Case Else
            Call createTab(Workbook, TabName, 255, 255, 255, TabAfter, True)
    End Select
    
    'Add Headers
    Call AddHeaders(TabName, Workbook, 3)
    
End Sub

Public Sub CreateCombinedMentorTab( _
TabName As String, _
GroupName As String, _
Workbook As Workbook, _
Optional TabAfter As String)

    If IsMissing(TabAfter) Then
        Let TabAfter = ""
    End If
        
    Select Case GroupName
        Case "1"
            Call createTab(Workbook, TabName, 142, 169, 219, TabAfter, True)
        Case "2"
            Call createTab(Workbook, TabName, 169, 208, 142, TabAfter, True)
        Case "3"
            Call createTab(Workbook, TabName, 255, 0, 0, TabAfter, True)
        Case Else
            Call createTab(Workbook, TabName, 255, 255, 255, TabAfter, True)
    End Select
    
    'Add Headers
    Call AddHeaders(TabName, Workbook, 3)
    
End Sub

Public Sub AddErrorsTab( _
ByRef wb As Workbook, _
ByVal errorCollection As Collection, _
ByRef TabAfter As String)

    Dim helper As GeneralHelperClass: Set helper = New GeneralHelperClass
    
    If IsMissing(TabAfter) Or Not (helper.sheetExists(wb, TabAfter)) Then
        Let TabAfter = wb.Sheets(Sheets.Count).Name
    End If
    
    Call createTab(wb, "Errors", 255, 0, 0, TabAfter, True)
    Call AddHeaders("Errors", wb, 4)
    
    If errorCollection.Count > 0 Then

        Dim RowNumber As Long
        Dim i As Long
        
        For i = 1 To errorCollection.Count
            Let RowNumber = i + 1
            With wb.Sheets("Errors")
                If errorCollection(i).studentID <= 1 Then
                    .Range("A" & RowNumber).Value = errorCollection(i).uniKey
                Else
                    .Range("A" & RowNumber).Value = errorCollection(i).studentID
                End If
                .Range("B" & RowNumber).Value = errorCollection(i).FirstName
                .Range("C" & RowNumber).Value = errorCollection(i).LastName
                .Range("D" & RowNumber).Value = errorCollection(i).International_YorN
                .Range("E" & RowNumber).Value = errorCollection(i).Dalyell_YorN
                .Range("F" & RowNumber).Value = errorCollection(i).MatureAge_YorN
                .Range("G" & RowNumber).Value = errorCollection(i).course.CourseName
                .Range("H" & RowNumber).Value = errorCollection(i).Major1
                .Range("I" & RowNumber).Value = errorCollection(i).Major2
                .Range("J" & RowNumber).Value = errorCollection(i).Error.returnErrorLocation()
                .Range("K" & RowNumber).Value = errorCollection(i).Error.ErrorName
                .Range("L" & RowNumber).Value = errorCollection(i).Error.ErrorDescription
            End With
        Next i
    End If
    
    wb.Sheets("Errors").UsedRange.Borders.LineStyle = xlContinuous
    wb.Sheets("Errors").Cells.EntireColumn.AutoFit
    
    Let TabAfter = "Errors"

End Sub

Public Sub AddStreamTabs( _
ByRef wb As Workbook, _
ByVal menteeCollection As Collection, _
ByRef Stream_Names As Variant, _
ByRef Group_Names As Variant, _
Optional ByRef TabAfter As String)
    
        'TESTING
        Dim t As TestingClass: Set t = New TestingClass
        Dim StartTime As Double
        Dim SecondsElapsed As Double
        Dim StartTimeAll As Double
        Dim SecondsElapsedAll As Double
        
        Call t.restartTestingTimer(StartTimeAll)
    
    'Instantiating global instances of classes whose functions need to be accessed:
    Dim tabGenerator As TabClass: Set tabGenerator = New TabClass
    Dim helper As GeneralHelperClass: Set helper = New GeneralHelperClass

    If IsMissing(TabAfter) Or Not (helper.sheetExists(wb, TabAfter)) Then
        Let TabAfter = wb.Sheets(Sheets.Count).Name
    End If

    Dim i As Long
    Dim j As Long
    
    Dim m As Mentee: Set m = New Mentee
    
    'As mentees are removed from their collection once they're sorted,
    'it's necessary to create a duplicate mentee collection so as not to lose data in the original.

            Debug.Print "Students in collection before sorting occurs: " & menteeCollection.Count
            Debug.Print ""

    'CREATE STREAMS TABS
    For i = 0 To helper.getLength(Stream_Names) - 1
            
            Debug.Print "---Itteration #" & i & " (" & Stream_Names(i) & ")---"
            
            Call t.restartTestingTimer(StartTime)
            
            
        Dim streamTabName As String: streamTabName = GenerateStreamTabName(CStr(Stream_Names(i)))
        Call tabGenerator.CreateStreamTab(streamTabName, wb, TabAfter)
        
        'Add menteeCollection to tab based on stream name
        Dim RowNumber As Long: RowNumber = 2
        For j = menteeCollection.Count To 1 Step -1

        If menteeCollection(j).stream.ShortName = CStr(Stream_Names(i)) Then
            With wb.Sheets(streamTabName)
                If menteeCollection(j).studentID <= 1 Then
                    .Range("A" & RowNumber).Value = menteeCollection(j).uniKey
                Else
                    .Range("A" & RowNumber).Value = menteeCollection(j).studentID
                End If
                .Range("B" & RowNumber).Value = menteeCollection(j).FirstName
                .Range("C" & RowNumber).Value = menteeCollection(j).LastName
                .Range("D" & RowNumber).Value = menteeCollection(j).International_YorN
                .Range("E" & RowNumber).Value = menteeCollection(j).Dalyell_YorN
                .Range("F" & RowNumber).Value = menteeCollection(j).MatureAge_YorN
                .Range("G" & RowNumber).Value = menteeCollection(j).course.CourseName
                .Range("H" & RowNumber).Value = menteeCollection(j).Major1
                .Range("I" & RowNumber).Value = menteeCollection(j).Major2
                .Range("J" & RowNumber).Value = menteeCollection(j).Email
                .Range("K" & RowNumber).Value = menteeCollection(j).Dietary
                .Range("L" & RowNumber).Value = menteeCollection(j).Mobility
                '.Range("M" & RowNumber).Value = menteeCollection(j).Mobile
                .Range("M" & RowNumber).Value = menteeCollection(j).stream.FullName
                .Range("N" & RowNumber).Value = menteeCollection(j).course.Group
            End With
        RowNumber = RowNumber + 1
        'menteeCollection.Remove (j)
        End If
        
            
        
        Next j
        wb.Sheets(streamTabName).UsedRange.Borders.LineStyle = xlContinuous
        wb.Sheets(streamTabName).Cells.EntireColumn.AutoFit
        TabAfter = streamTabName
        
            Call t.endTestingTimer(StartTime, SecondsElapsed, "   Processing time: ")
            Debug.Print "   Students remaining in collection: " & menteeCollection.Count
            Debug.Print ""
    Next i
    
        Call t.endTestingTimer(StartTimeAll, SecondsElapsedAll, "Function Processing Time: ")
End Sub

Public Sub AddGroupTabs( _
ByRef wb As Workbook, _
ByRef mentorCollection As Collection, _
ByRef Stream_Names As Variant, _
ByRef Group_Names As Variant, _
ByRef menteesPerMentor As Byte, _
ByRef TabAfter As String)
    
    'A nested loop will loop through group names(i) and stream names(j).
    'Tab order:
    'Mentors - Combined 1 (i = 0)
    '    1 Groups Dom!!
    '    1 Groups Int!!
    '    1 Groups Dalyell!!
    '    1 Groups 25+!!
    'Mentors - Combined 2 (i = 1)
    '    2 Groups Dom!!
    '    2 Groups Int!!
    '    2 Groups Dalyell!!
    '    2 Groups 25+!!
    'Mentors - Combined 3 (i = 2)
    '    3 Groups Dom!!
    '    3 Groups Int!!
    '    3 Groups Dalyell!!
    '    3 Groups 25+!!
    
    Dim helper As GeneralHelperClass: Set helper = New GeneralHelperClass
    
    If IsMissing(TabAfter) Or Not (helper.sheetExists(wb, TabAfter)) Then
        Let TabAfter = wb.Sheets(Sheets.Count).Name
    End If
    
    Dim RowNumber As Long
    Dim mentorCount As Long
    Dim mentorLabel As String
    Dim studentCount As Byte
    
    'Global variables for loops:
    Dim i As Long
    Dim j As Long
    Dim k As Byte
    Dim l As Byte
    
    '---LOOP i---
    'Looping through group names ("1", "2", "3")
    
    For i = 0 To helper.getLength(Group_Names) - 1
        '1. Create combined mentee tab
        Dim combinedMentorTabName As String:
        Let combinedMentorTabName = GenerateCombinedMenteeTabName(CStr(Group_Names(i)))
        
        Call CreateCombinedMentorTab( _
        combinedMentorTabName, _
        CStr(Group_Names(i)), _
        wb, _
        TabAfter)

            'Add mentors to tab based on group
            Let RowNumber = 2
            Let mentorCount = 1
            
            '---LOOP j---
            'Looping through mentors
            
            For j = 1 To mentorCollection.Count
            If CStr(mentorCollection(j).course.Group) = CStr(Group_Names(i)) Then
                With wb.Sheets(combinedMentorTabName)
                    'Add group number
                    If Not (IsMissing(mentorCollection(j).groupNumber)) Then
                        .Range("A" & RowNumber).Value = mentorCollection(j).returnGroupLabel()
                    End If
                    
                    'Add SID/Unikey
                    If mentorCollection(j).studentID <= 1 Then
                       .Range("D" & RowNumber).Value = mentorCollection(j).uniKey
                    Else
                        .Range("D" & RowNumber).Value = mentorCollection(j).studentID
                    End If
                    .Range("B" & RowNumber).Value = mentorCollection(j).Students.Count
                    .Range("C" & RowNumber).Value = mentorCollection(j).stream.FullName
                    .Range("E" & RowNumber).Value = mentorCollection(j).FirstName
                    .Range("F" & RowNumber).Value = mentorCollection(j).LastName
                    .Range("G" & RowNumber).Value = mentorCollection(j).course.CourseName
                    .Range("H" & RowNumber).Value = mentorCollection(j).Major1
                    .Range("I" & RowNumber).Value = mentorCollection(j).Major2
                    .Range("J" & RowNumber).Value = mentorCollection(j).Email
                    .Range("K" & RowNumber).Value = "-"
                End With
            Let RowNumber = RowNumber + 1
            Let studentCount = mentorCollection(j).Students.Count
                
                '--- LOOP k ---
                'Adding mentors' mentees:
                For k = 1 To menteesPerMentor
                    If k <= studentCount Then
                        With wb.Sheets(combinedMentorTabName)
                            'Add SID/Unikey
                            If mentorCollection(j).Students(k).studentID <= 1 Then
                               .Range("D" & RowNumber).Value = mentorCollection(j).Students(k).uniKey
                            Else
                                .Range("D" & RowNumber).Value = mentorCollection(j).Students(k).studentID
                            End If
                            
                            .Range("C" & RowNumber).Value = mentorCollection(j).Students(k).stream.FullName
                            .Range("E" & RowNumber).Value = mentorCollection(j).Students(k).FirstName
                            .Range("F" & RowNumber).Value = mentorCollection(j).Students(k).LastName
                            .Range("G" & RowNumber).Value = mentorCollection(j).Students(k).course.CourseName
                            .Range("H" & RowNumber).Value = mentorCollection(j).Students(k).Major1
                            .Range("I" & RowNumber).Value = mentorCollection(j).Students(k).Major2
                            .Range("J" & RowNumber).Value = mentorCollection(j).Students(k).Email
                            .Range("K" & RowNumber).Value = mentorCollection(j).Students(k).Mobility
                            .Range("L" & RowNumber).Value = mentorCollection(j).Students(k).MatchType
                        End With
                    End If
                    RowNumber = RowNumber + 1
                Next k
            End If
        Next j
        TabAfter = combinedMentorTabName
    
            '---LOOP l---
            'Looping through streams ("Int", "Dom" etc.)
            
            For l = 0 To helper.getLength(Stream_Names) - 1
                Dim groupTabName As String: groupTabName = GenerateGroupPlusStreamTabName(CStr(Group_Names(i)), CStr(Stream_Names(l)))
                Call CreateGroupTab(groupTabName, CStr(Group_Names(i)), wb, TabAfter)
                'Add mentors to each tab
                Let RowNumber = 2
                Let mentorCount = 1
                
                '---LOOP j---
                'Looping through mentors
                
                For j = 1 To mentorCollection.Count
                If CStr(mentorCollection(j).course.Group) = CStr(Group_Names(i)) And _
                mentorCollection(j).stream.ShortName = Stream_Names(l) _
                Then
                    With wb.Sheets(groupTabName)
                        'Add group number
                        If Not (IsMissing(mentorCollection(j).groupNumber)) Then
                            .Range("A" & RowNumber).Value = mentorCollection(j).returnGroupLabel()
                        End If
                        
                        'Add SID/Unikey
                        If mentorCollection(j).studentID <= 1 Then
                           .Range("D" & RowNumber).Value = mentorCollection(j).uniKey
                        Else
                            .Range("D" & RowNumber).Value = mentorCollection(j).studentID
                        End If
                        .Range("B" & RowNumber).Value = mentorCollection(j).Students.Count
                        .Range("C" & RowNumber).Value = mentorCollection(j).stream.FullName
                        .Range("E" & RowNumber).Value = mentorCollection(j).FirstName
                        .Range("F" & RowNumber).Value = mentorCollection(j).LastName
                        .Range("G" & RowNumber).Value = mentorCollection(j).course.CourseName
                        .Range("H" & RowNumber).Value = mentorCollection(j).Major1
                        .Range("I" & RowNumber).Value = mentorCollection(j).Major2
                        .Range("J" & RowNumber).Value = mentorCollection(j).Email
                        .Range("K" & RowNumber).Value = "-"
                    End With
                Let RowNumber = RowNumber + 1
                Let studentCount = mentorCollection(j).Students.Count
                
                '---LOOP k---
                'Adding mentors' mentees:
                
                For k = 1 To menteesPerMentor
                    If k <= studentCount Then
                        With wb.Sheets(groupTabName)
                            'Add SID/Unikey
                            If mentorCollection(j).Students(k).studentID <= 1 Then
                               .Range("D" & RowNumber).Value = mentorCollection(j).Students(k).uniKey
                            Else
                                .Range("D" & RowNumber).Value = mentorCollection(j).Students(k).studentID
                            End If
                            
                            .Range("C" & RowNumber).Value = mentorCollection(j).Students(k).stream.FullName
                            .Range("E" & RowNumber).Value = mentorCollection(j).Students(k).FirstName
                            .Range("F" & RowNumber).Value = mentorCollection(j).Students(k).LastName
                            .Range("G" & RowNumber).Value = mentorCollection(j).Students(k).course.CourseName
                            .Range("H" & RowNumber).Value = mentorCollection(j).Students(k).Major1
                            .Range("I" & RowNumber).Value = mentorCollection(j).Students(k).Major2
                            .Range("J" & RowNumber).Value = mentorCollection(j).Students(k).Email
                            .Range("K" & RowNumber).Value = mentorCollection(j).Students(k).Mobility
                            .Range("L" & RowNumber).Value = mentorCollection(j).Students(k).MatchType
                        End With
                    End If
                    RowNumber = RowNumber + 1
                Next k
            End If
            Next j
            wb.Sheets(groupTabName).UsedRange.Borders.LineStyle = xlContinuous
            wb.Sheets(groupTabName).Cells.EntireColumn.AutoFit
            TabAfter = groupTabName
        Next l
        wb.Sheets(combinedMentorTabName).UsedRange.Borders.LineStyle = xlContinuous
        wb.Sheets(combinedMentorTabName).Cells.EntireColumn.AutoFit
    Next i
End Sub

Sub AddHeaders(Tab_Name As String, Workbook As Workbook, Optional headerType As Long)
    Dim helper As New GeneralHelperClass
    Dim headerNames As Variant
    
    'headerType indicates which type of header is required
    '   OPTION 1: Mentee
    '   OPTION 2: Mentor
    '   OPTION 3: Combined
    '   OPTION 4: Errors

    If IsMissing(headerType) Then
        headerType = 1
    End If
        
    Select Case headerType
        'Mentee headers:
        Case 1
            headerNames = Array("SID", "FirstName", "LastName", "Int", "Dayell", "25+", "Course", "Major 1", "Major 2", "Email", "Dietary", "Mobility", "Stream", "Group")
        'Mentor headers:
        Case 2
            headerNames = Array("SID", "FirstName", "LastName", "Int", "Dayell", "25+", "Course", "Major 1", "Major 2", "Email", "Dietary", "Mobility", "Stream", "Group")
        'Combined headers:
        Case 3
            headerNames = Array("Group No", "Mentee Count", "Stream", "SID", "Mentor/Mentee First", "Mentor/Mentee Last", "Course", "Major1", "Major2", "Email", "Mobility", "Mentee Match Type", "Mentor Start Time", "Mentee Start Time", "Small Group", "Degree Talk", "Degree Talk Time", "Degree Talk Venue", "eLearning")
        'Else, default to mentee headers
        Case 4
            headerNames = Array("SID", "FirstName", "LastName", "Int", "Dayell", "25+", "Course", "Major 1", "Major 2", "Error Location", "Error Name", "Error Details")
        Case Else
            headerNames = Array("SID", "FirstName", "LastName", "Int", "Dayell", "25+", "Course", "Major 1", "Major 2", "Email", "Dietary", "Mobility", "Stream", "Group", "Moved?")
    End Select
    
    Dim headerLength As Long: Let headerLength = CLng(helper.getLength(headerNames))
        
    'Dim headers As Range: Set headers = Workbook.Worksheets(Tab_Name).Range(Cells(1, 1), Cells(1, headerLength))
    Dim i As Long
    For i = 1 To headerLength
        With Workbook.Worksheets(Tab_Name).Cells(1, i)
            .Value = headerNames(i - 1)
            .Font.Bold = True
            .Interior.Color = RGB(180, 198, 231)
        End With
    Next i
End Sub

Public Sub DeleteAllTabs( _
ByRef Workbook As Workbook, _
ByRef TabsToKeep As Variant, _
ByRef GlobalTabAfter As String)

    Dim helper As GeneralHelperClass: Set helper = New GeneralHelperClass
    
    'Turn off notifications
    Application.DisplayAlerts = False
    
    Dim i As Integer
    Dim j As Byte
    Dim sheetOnTabsToKeepList As Boolean
    
    'Creating a dummy sheet so that at least one tab remains once all tabs are deleted.
    If Not (helper.sheetExists(Workbook, "DummyWorksheet")) Then
        Dim dummyWorksheet As Worksheet
        Set dummyWorksheet = Workbook.Worksheets.Add(Type:=xlWorksheet, after:=Sheets(Sheets.Count))
        With dummyWorksheet
            .Name = "DummyWorksheet"
        End With
    End If
    
    'Set global 'TabAfter' variable to dummy tab's name
    Let GlobalTabAfter = "DummyWorksheet"
    
    'Loop through
    For i = (Workbook.Worksheets.Count - 1) To 1 Step -1
    sheetOnTabsToKeepList = False
        For j = 0 To helper.getLength(TabsToKeep) - 1
            If Workbook.Worksheets(i).Name = TabsToKeep(j) Then
                sheetOnTabsToKeepList = True
                Exit For
            End If
        Next j
        If Not (sheetOnTabsToKeepList) Then
            Workbook.Worksheets(i).Delete
        End If
    Next i

    'Turn notifications back on
    Application.DisplayAlerts = True
    
End Sub


Public Function GenerateTabName( _
Prefix As String, MainName As String, _
Suffix As String, AddSpaces As Boolean) As String

    If AddSpaces Then
        If Prefix = "" Then
        GenerateTabName = MainName & " " & Suffix
        ElseIf Suffix = "" Then
        GenerateTabName = Prefix & " " & MainName
        Else
        GenerateTabName = Prefix & " " & MainName & Suffix
        End If
    Else
        GenerateTabName = Prefix & MainName & Suffix
    End If
End Function

Public Function GenerateStreamTabName(StreamName As String) As String
    If IsMissing(StreamName) Or StreamName = "" Then
        GenerateStreamTabName = ""
    Else
        GenerateStreamTabName = StreamName & " Stream"
    End If
End Function

Public Function GenerateGroupPlusStreamTabName(GroupName As String, StreamName As String) As String
    If IsMissing(StreamName) Or IsMissing(GroupName) Then
        GenerateGroupPlusStreamTabName = ""
    Else
        GenerateGroupPlusStreamTabName = GroupName & " Groups " & StreamName & "!!"
    End If
End Function

Public Function GenerateCombinedMenteeTabName(GroupName As String) As String
        GenerateCombinedMenteeTabName = "Mentees - Combined " & GroupName
End Function