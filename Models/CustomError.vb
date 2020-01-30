Option Explicit

'---CONSTRUCTOR---

Public Function createCustomError( _
errorNumber As Long, _
TabName As String, _
RowNumber As Long) _
As CustomError
    
    '---ERROR TYPES---
    'Ordered ascendingly by importance
    '   0 - Invalid SID
    '   1 - Invalid SID And uniKey
    '   2 - Duplicate Entry
    '   3 - Stream Unknown
    '   4 - Course Unknown
    '   Else - Error (general error)
    
    Dim err As CustomError: Set err = New CustomError
    err.errorNumber = errorNumber
    err.TabFound = TabName
    err.RowNumber = RowNumber
    
    Select Case errorNumber
        Case 0
            err.ErrorName = "Invalid SID"
            err.ErrorDescription = "A valid Student ID (SID) wasn't found. A valid SID consists of 9 numbers and no letters."
        Case 1
            err.ErrorName = "Invalid SID and UniKey"
            err.ErrorDescription = "Neither a valid Student ID (SID) nor a valid UniKey were found."
        Case 2
            err.ErrorName = "Duplicate Entry"
            err.ErrorDescription = "Another student was found with the same SID/UniKey."
        Case 3
            err.ErrorName = "Stream Unknown"
            err.ErrorDescription = "This student hasn't properly indicated which stream they belong to."
        Case 4
            err.ErrorName = "Course Unknown"
            err.ErrorDescription = "This student's course wasn't found. Please check the course/group list and ensure that it is updated."
        Case Else
            err.ErrorName = "Error Occured"
            err.ErrorDescription = "An unknown error occured with this student. Please review."
    End Select

    Set createCustomError = err
End Function

'---PROPERTIES---

Private pErrorNumber As Long
Private pErrorName As String
Private pErrorDescription As String
Private pTabFound As String
Private pRowNumber As Long

Public Property Get errorNumber() As Long
    errorNumber = pErrorNumber
End Property
Public Property Let errorNumber(Value As Long)
    pErrorNumber = Value
End Property

Public Property Get ErrorName() As String
    ErrorName = pErrorName
End Property
Public Property Let ErrorName(Value As String)
    pErrorName = Value
End Property

Public Property Get ErrorDescription() As String
    ErrorDescription = pErrorDescription
End Property
Public Property Let ErrorDescription(Value As String)
    pErrorDescription = Value
End Property

Public Property Get TabFound() As String
    TabFound = pTabFound
End Property
Public Property Let TabFound(Value As String)
    pTabFound = Value
End Property

Public Property Get RowNumber() As Long
    RowNumber = pRowNumber
End Property
Public Property Let RowNumber(Value As Long)
    pRowNumber = Value
End Property

'---METHODS---

Public Function returnErrorDetails()
    returnErrorDetails = _
    "Error " & pErrorNumber & ": " & pErrorName & _
    " (" & pTabFound & ", Row " & pRowNumber & ") - " & pErrorDescription
End Function

Public Function returnErrorLocation()
    returnErrorLocation = pTabFound & ", Row " & pRowNumber
End Function