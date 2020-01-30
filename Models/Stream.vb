Option Explicit

'---CONSTRUCTOR---

Public Function CreateStream( _
ByVal Int_Column_Y_or_N As String, _
ByVal Dalyell_Column_Y_or_N As String, _
ByVal MatureAge_Column_Y_or_N As String) As stream

    Dim stream As stream: Set stream = New stream
    
    If UCase(Dalyell_Column_Y_or_N) = "YES" Or UCase(Dalyell_Column_Y_or_N) = "Y" Then
        stream.FullName = "Dalyell"
        stream.ShortName = "DL"
    ElseIf UCase(Int_Column_Y_or_N) = "YES" Or UCase(Int_Column_Y_or_N) = "Y" Then
        stream.FullName = "International"
        stream.ShortName = "Int"
    ElseIf UCase(MatureAge_Column_Y_or_N) = "YES" Or UCase(MatureAge_Column_Y_or_N) = "Y" Then
        stream.FullName = "25+"
        stream.ShortName = "25+"
    Else
        stream.FullName = "Domestic"
        stream.ShortName = "Dom"
    End If
    
    Set CreateStream = stream
End Function

'---PROPERTIES---

Private pFullName As String
Private pShortName As String

Public Property Get FullName() As String
    FullName = pFullName
End Property
Public Property Let FullName(Value As String)
    pFullName = Value
End Property

Public Property Get ShortName() As String
    ShortName = pShortName
End Property
Public Property Let ShortName(Value As String)
    pShortName = Value
End Property

'---METHODS---

Public Function ValidStreamFound( _
ByVal Int_Column_Y_or_N As String, _
ByVal Dalyell_Column_Y_or_N As String, _
ByVal MatureAge_Column_Y_or_N As String, _
Optional ByRef helper As GeneralHelperClass) As Boolean

	Dim lettersFound As Boolean: Let lettersFound = False

	If IsMissing(helper) Then
		Set helper = New GeneralHelperClass
	End If
	
	If helper.lettersFound(Int_Column_Y_or_N) Or _
	   helper.lettersFound(Dalyell_Column_Y_or_N) Or _
	   helper.lettersFound(MatureAge_Column_Y_or_N) Then
	   
	   Let lettersFound = True	   
	End If
	
	Let ValidStreamFound = lettersFound
End Function
