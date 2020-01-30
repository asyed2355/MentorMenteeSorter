VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OpeningForm 
   Caption         =   "Mentee/Mentor Sorter"
   ClientHeight    =   7440
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   11400
   OleObjectBlob   =   "OpeningForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "OpeningForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---Userform Initialize/Terminate handlers---

Private Sub UserForm_Initialize()
    txtMenteeMentorLocation.Text = "C:\Users\a\Mentor Mentee List.xlsx"
    txtDegreeByGroupLocation.Text = "C:\Users\a\Degrees By Group.xlsx"
    txtMenteesPerMentor.Text = "10"
    txtMenteeTab.Text = "Mentee Dump"
    txtMentorTab.Text = "Mentors"
    txtDegreeTab.Text = "Degree_List"
End Sub

Private Sub cmdUnload_Click()
    Unload Me
    End
End Sub


'---Button event handlers---

Private Sub degreeByCategoryFileSelect_Click()
    Dim txt As String: Let txt = txtDegreeByGroupLocation.Text
    txtDegreeByGroupLocation.Text = openDirectory(txt, "Open Degrees-by-Group reference file")
End Sub

Private Sub menteeMentorFileSelect_Click()
    Dim txt As String: Let txt = txtMenteeMentorLocation.Text
    txtMenteeMentorLocation.Text = openDirectory(txt, "Open Mentor/Mentee student list")
End Sub

Private Sub btnRun_Click()
    'Check that the spreadsheets can be found
    If Dir(txtMenteeMentorLocation.Text) = "" Then
        MsgBox "Mentee/Mentor spreadsheet NOT FOUND (" & txtMenteeMentorLocation.Text & ")."
    ElseIf Dir(txtDegreeByGroupLocation.Text) = "" Then
        MsgBox "Degree-By-Group spreadsheet NOT FOUND (" & txtDegreeByGroupLocation.Text & ")."
    'If both spreadsheets are found, continue
    Else
        Let strFilename = txtMenteeMentorLocation.Text
        Let degreesByGroupFilename = txtDegreeByGroupLocation.Text
        
        'Check that the number entered for #mentees/mentors is valid. If not, change to default value (10).
        If IsNumeric(txtMenteesPerMentor.Text) And _
           CLng(txtMenteesPerMentor.Text) > 0 And _
           CLng(txtMenteesPerMentor.Text) <= 255 Then
                Let menteesPerMentor = CByte(txtMenteesPerMentor.Text)
        Else
                Let menteesPerMentor = 10
        End If
        
        Let MenteeTabName = txtMenteeTab.Text
        Let MentorTabName = txtMentorTab.Text
        Let DegreesByCategoryTabName = txtDegreeTab.Text

        Unload OpeningForm
    End If
End Sub


'---Function to allow users to open file using FileDiablog box---

Private Function openDirectory( _
Optional currentTextBoxValue As String, _
Optional dialogBoxTitle As String _
) As String

    'Function that allows the user to select files using a FileDialog box.
    'SOURCE*: https://stackoverflow.com/questions/10304989/open-windows-explorer-and-select-a-file
    '*Modified slightly to suit this project.
    
    'Handle optional arguments if they aren't provided.
    If IsMissing(currentTextBoxValue) Then
        Let currentTextBoxValue = ""
    End If
    If IsMissing(dialogBoxTitle) Then
        Let dialogBoxTitle = "Please select file"
    End If
       
    Dim f As Office.FileDialog
    Set f = Application.FileDialog(msoFileDialogFilePicker)
    
    With f
        .AllowMultiSelect = False
        .Title = dialogBoxTitle
        .Filters.Clear
        .Filters.Add "Excel File", "*.xlsx"
        .Filters.Add "Excel Macro-Enabled File", "*.xlsm"
        .Filters.Add "Excel (Legacy Format)", "*.xls"
        .Filters.Add "All Files", "*.*"
        
        ' Show the dialog box. If the .Show method returns True, the
        ' user picked at least one file. If the .Show method returns
        ' False, the user clicked Cancel.
        If .Show = True Then
            openDirectory = .SelectedItems(1)
        Else
            openDirectory = currentTextBoxValue
        End If
    End With
End Function
