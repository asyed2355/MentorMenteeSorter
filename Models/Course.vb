Option Explicit

'---PROPERTIES---

Private pCourseName As String
Private pGroup As String

Public Property Get CourseName() As String
    CourseName = pCourseName
End Property
Public Property Let CourseName(Value As String)
    pCourseName = Value
End Property

Public Property Get Group() As String
    Group = pGroup
End Property
Public Property Let Group(Value As String)
    pGroup = Value
End Property

'---METHODS---

Public Function DetermineCourseGroup( _
CourseName As String, _
DegreesByCategoryCollection As Collection) As String
    Dim helper As GeneralHelperClass: Set helper = New GeneralHelperClass
    Dim course As course: Set course = New course
    
    Dim result As String: result = "Course Not Found"
    
    For Each course In DegreesByCategoryCollection
        If UCase(course.CourseName) = UCase(CourseName) Then
            result = course.Group
            Exit For
        End If
    Next course
    DetermineCourseGroup = result
End Function
