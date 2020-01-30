Option Explicit

'This class is designed to test the application in various ways.
'Primarily, this is used to test performance/calculation times.

Public Sub restartTestingTimer(ByRef StartTimeVariable As Double)
    'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault
    Let StartTimeVariable = Timer
End Sub

Public Sub endTestingTimer( _
ByRef StartTimeVariable As Double, _
ByRef SecondsElapsedVariable As Double, _
msg As String)
    'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault
    Let SecondsElapsedVariable = Round(Timer - StartTimeVariable, 2)
    Debug.Print msg; ": " & SecondsElapsedVariable & " seconds."
End Sub

Public Sub clearStudentsFromMemory(ByRef mentees As Collection, ByRef mentors As Collection)
    Set mentees = Nothing
    Set mentors = Nothing
End Sub

Public Sub testStreamTabs()
    Call testDeleteAllTabs
    
    '--- Array names ---
    Dim streamNames() As Variant: Let streamNames = Array("Dom", "Int", "DL", "25+")
    Dim groups() As Variant: Let groups = Array("1", "2", "3")
    Dim h As GeneralHelperClass: Set h = New GeneralHelperClass
    Dim t As TabClass: Set t = New TabClass
    Dim mentees As Collection: Set mentees = h.createDummyStudentCollection(500, 1)
    Dim mentors As Collection: Set mentors = h.createDummyStudentCollection(50, 2)
    
    
    
    Dim StartTime As Double
    Dim SecondsElapsed As Double
    
    Call restartTestingTimer(StartTime)
    'Run Function
    Call t.AddStreamTabs(ActiveWorkbook, mentees, streamNames, groups, ActiveWorkbook.Sheets(Sheets.Count).Name)
    Call endTestingTimer(StartTime, SecondsElapsed, "Stream Tab function")

End Sub

Public Sub testDeleteAllTabs()
    'Turn off notifications
    Application.DisplayAlerts = False
    
    Dim helper As GeneralHelperClass: Set helper = New GeneralHelperClass
    
    'Creating a dummy sheet so that at least one tab remains once all tabs are deleted.
    If Not (helper.sheetExists(ActiveWorkbook, "DummyWorksheet")) Then
        Dim dummyWorksheet As Worksheet
        Set dummyWorksheet = ActiveWorkbook.Worksheets.Add(Type:=xlWorksheet, after:=Sheets(Sheets.Count))
        With dummyWorksheet
            .Name = "DummyWorksheet"
        End With
    End If
    
    Dim i As Long

    'Loop through and delete all tabs except for the dummy worksheet
    For i = (ActiveWorkbook.Worksheets.Count - 1) To 1 Step -1
        If ActiveWorkbook.Worksheets(i).Name <> "DummyWorksheet" Then
            ActiveWorkbook.Worksheets(i).Delete
        End If
    Next i
    
    'Turn notifications back on
    Application.DisplayAlerts = True
End Sub