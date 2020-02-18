'MAIN SUB-ROUTINE
Option Explicit

'---File locations---
Public strFilename As String
Public degreesByGroupFilename As String

'---Tab names---
Public MenteeTabName As String
Public MentorTabName As String
Public DegreesByCategoryTabName As String

'---Mentees per mentor---
Public menteesPerMentor As Byte

'--- Helper class ---
Public helper As GeneralHelperClass
Public entityCreator As EntityConverter
Public DbContext As DbContext

'--- Check to ensure that form wasn't closed---
Public runProgram As Boolean

Public Sub FormatMentees()
'
' FormatMentees Macro
'

'
    Let runProgram = False
    
    'Open Userform
    OpeningForm.Show vbModal
    
    If Not (runProgram) Then
        Exit Sub
    End If
    
    'Instantiate global helper and dbContext objects
    Set helper = New GeneralHelperClass
    Set entityCreator = New EntityConverter
    Set DbContext = New DbContext
    
    'Open workbooks
    Dim CurrentWorkbook As Workbook: Set CurrentWorkbook = DbContext.openWorkbook(strFilename)
    Dim degreesByGroupWorkbook As Workbook: Set degreesByGroupWorkbook = DbContext.openWorkbook(degreesByGroupFilename)
    
    'Check that worksheets exists. If not, exit sub
    If Not (helper.sheetExists(CurrentWorkbook, MenteeTabName)) Then
        MsgBox "Mentee tab not found."
        Exit Sub
    ElseIf Not (helper.sheetExists(CurrentWorkbook, MentorTabName)) Then
        MsgBox "Mentor tab not found."
        Exit Sub
    ElseIf Not (helper.sheetExists(degreesByGroupWorkbook, DegreesByCategoryTabName)) Then
        MsgBox "Degree List tab not found."
        Exit Sub
    End If
        
'<----- GLOBAL VARIABLES start ----->

    '--- Worksheet objects ---
    Dim MenteeTab As Worksheet: Set MenteeTab = CurrentWorkbook.Worksheets(MenteeTabName)
    Dim MentorTab As Worksheet: Set MentorTab = CurrentWorkbook.Worksheets(MentorTabName)
    
    '--- Stream Names ---
    Dim domesticStream_ABR As String: domesticStream_ABR = "Dom"
    Dim internationalStream_ABR As String: internationalStream_ABR = "Int"
    Dim dalyellStream_ABR As String: dalyellStream_ABR = "DL"
    Dim matureAgeStream_ABR As String: matureAgeStream_ABR = "25+"
    Dim suffixStreamTab As String: suffixStreamTab = "Stream"
    
    '--- Arrays (names) ---
    Dim mainTabNameArray() As Variant
    Dim streamNames() As Variant
    Dim groups() As Variant
    Let mainTabNameArray = Array(MenteeTabName, MentorTabName, DegreesByCategoryTabName)
    Let streamNames = Array(domesticStream_ABR, internationalStream_ABR, dalyellStream_ABR, matureAgeStream_ABR)
    Let groups = Array("1", "2", "3")
    
    '--- Misc ---
    Dim GlobalTabAfter As String: Let GlobalTabAfter = "START_HERE"
    Dim newTab As TabClass: Set newTab = New TabClass
 
    '--- Mentor, mentee and course collections ---
    Dim courses As Collection: Set courses = entityCreator.createCourseCollection(degreesByGroupWorkbook, DegreesByCategoryTabName)
        
    Dim mentees As Collection: Set mentees = New Collection
    Dim mentors As Collection: Set mentors = New Collection
    Dim studentsWithErrors As Collection: Set studentsWithErrors = New Collection
    Dim unmatchedMentees As Collection: Set unmatchedMentees = New Collection
    
'<----- GLOBAL VARIABLES end ----->

'<----- FUNCTION CALLS start ----->

    'Close 'Degrees By Group' spreadsheet
    Call DbContext.closeWorkbook(degreesByGroupWorkbook, False)
 
    'Populate mentee, mentor and error collections
    Call entityCreator.createStudentCollections(CurrentWorkbook, _
    MenteeTab, MentorTab, courses, mentees, mentors, studentsWithErrors)
    
    'Close mentee/mentor spreadsheet
    Call DbContext.closeWorkbook(CurrentWorkbook, False)
    
    'Set CurrentWorkbook to this workbook
    Set CurrentWorkbook = ThisWorkbook
    Call newTab.DeleteAllTabs(CurrentWorkbook, GlobalTabAfter)
    Call newTab.AddStreamTabs(CurrentWorkbook, mentees, streamNames, groups, GlobalTabAfter)
    Call newTab.AddErrorsTab(CurrentWorkbook, studentsWithErrors, GlobalTabAfter)
    Call helper.MatchMenteesWithMentors(mentees, mentors, menteesPerMentor)
    Call newTab.AddGroupTabs(CurrentWorkbook, mentors, streamNames, groups, menteesPerMentor, GlobalTabAfter)
    Call newTab.AddUnmatchedMenteesTab(CurrentWorkbook, mentees, GlobalTabAfter)

'<----- FUNCTION CALLS end ----->

    If helper.sheetExists(CurrentWorkbook, "DummyWorksheet") Then
        Application.DisplayAlerts = False
        CurrentWorkbook.Sheets("DummyWorksheet").Delete
        Application.DisplayAlerts = True
    End If
End Sub

Public Sub ClearAllTabs()
    Dim a As Integer
    a = MsgBox("Are you sure you want to clear all the tabs in this spreadsheet?", vbQuestion + vbYesNo + vbDefaultButton2, "Clear all tabs?")
    
    If a = vbYes Then
        Dim t As TabClass: Set t = New TabClass
        Call t.DeleteAllTabs(ActiveWorkbook, "START_HERE")
    End If
    
End Sub
