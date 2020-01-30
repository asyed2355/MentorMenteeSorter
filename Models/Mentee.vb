Option Explicit

'---CONSTRUCTOR---

Private Sub Class_Initialize()
    Moved = True
End Sub

'---PROPERTIES---

Private pStudentID As Long
Private pUniKey As String
Private pFirstName As String
Private pLastName As String
Private pCourse As course
Private pMajor1 As String
Private pMajor2 As String
Private pEmail As String
Private pDietary As String
Private pMobility As String
Private pInternational_YorN As String
Private pDalyell_YorN As String
Private pMatureAge_YorN As String
Private pStream As stream
Private pGroup As String
Private pError As CustomError
Private pMoved As Boolean
Private pMatchType As Byte

Public Property Get studentID() As Long
    studentID = pStudentID
End Property
Public Property Let studentID(Value As Long)
    If IsNumeric(Value) Then
        If Len(Value) > 9 Then
            pStudentID = CLng(Left(Value, 9))
        Else
            pStudentID = Value
        End If
    Else
        pStudentID = 0
    End If
End Property

Public Property Get uniKey() As String
    uniKey = pUniKey
End Property
Public Property Let uniKey(Value As String)
    pUniKey = Value
End Property

Public Property Get FirstName() As String
    FirstName = pFirstName
End Property
Public Property Let FirstName(Value As String)
    pFirstName = Value
End Property

Public Property Get LastName() As String
    LastName = pLastName
End Property
Public Property Let LastName(Value As String)
    pLastName = Value
End Property

Public Property Get course() As course
    Set course = pCourse
End Property
Public Property Let course(Value As course)
    Set pCourse = Value
End Property

Public Property Get Major1() As String
    Major1 = pMajor1
End Property
Public Property Let Major1(Value As String)
    If Value = "-" Or Value = " - " Then
        pMajor1 = ""
    Else
        pMajor1 = Value
    End If
End Property

Public Property Get Major2() As String
    Major2 = pMajor2
End Property
Public Property Let Major2(Value As String)
    pMajor2 = Value
End Property

Public Property Get Email() As String
    Email = pEmail
End Property
Public Property Let Email(Value As String)
    pEmail = Value
End Property

Public Property Get Dietary() As String
    Dietary = pDietary
End Property
Public Property Let Dietary(Value As String)
    pDietary = Value
End Property

Public Property Get Mobility() As String
    Mobility = pMobility
End Property
Public Property Let Mobility(Value As String)
    pMobility = Value
End Property

Public Property Get International_YorN() As String
    International_YorN = pInternational_YorN
End Property
Public Property Let International_YorN(Value As String)
    pInternational_YorN = Value
End Property

Public Property Get Dalyell_YorN() As String
    Dalyell_YorN = pDalyell_YorN
End Property
Public Property Let Dalyell_YorN(Value As String)
    pDalyell_YorN = Value
End Property

Public Property Get MatureAge_YorN() As String
    MatureAge_YorN = pMatureAge_YorN
End Property
Public Property Let MatureAge_YorN(Value As String)
    pMatureAge_YorN = Value
End Property

Public Property Get stream() As stream
    Set stream = pStream
End Property
Public Property Let stream(Value As stream)
    Set pStream = Value
End Property

Public Property Get Group() As String
    Group = pGroup
End Property
Public Property Let Group(Value As String)
    pGroup = Value
End Property

Public Property Get Error() As CustomError
    Set Error = pError
End Property
Public Property Let Error(Value As CustomError)
    Set pError = Value
End Property

Public Property Get Moved() As Boolean
    Moved = pMoved
End Property
Public Property Let Moved(Value As Boolean)
    pMoved = Value
End Property

Public Property Get MatchType() As Byte
    MatchType = pMatchType
End Property
Public Property Let MatchType(Value As Byte)
    pMatchType = Value
End Property