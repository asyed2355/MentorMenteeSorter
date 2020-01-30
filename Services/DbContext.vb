Option Explicit

'This class is responsible for opening/closing external Excel sheets and databases.

'---METHODS---

Public Function openWorkbook(directory As String) As Workbook
    'Check that the spreadsheets can be found
    If Dir(directory) <> "" Then
        Set openWorkbook = Workbooks.Open(Filename:=directory)
    Else
        Set openWorkbook = ActiveWorkbook
    End If
End Function

Public Sub closeWorkbook(wb As Workbook, Optional saveChangesBeforeClosing As Boolean)
    If IsMissing(saveChangesBeforeClosing) Then
        Let saveChangesBeforeClosing = False
    End If

    wb.Close saveChanges:=saveChangesBeforeClosing
End Sub

Public Function workbookExists(fileDirectory As String) As Boolean
    If Dir("fileDirectory") <> "" Then
        Let workbookExists = True
    Else
        Let workbookExists = False
    End If
End Function
