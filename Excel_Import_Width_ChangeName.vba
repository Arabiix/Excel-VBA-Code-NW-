Sub ImportExcelDirectly()
    Dim ws As Worksheet
    Dim filePath As String
    Dim fileDialog As Object
    Dim downloadsFolder As String
    Dim importWorkbook As Workbook
    Dim newSheet As Worksheet
    Dim sheetName As String
    Dim i As Integer
    
    On Error GoTo ErrorHandler ' Activate the general error handler for the entire subroutine
    
    ' Get the Downloads folder path directly using Windows Shell
    downloadsFolder = Environ("USERPROFILE") & "\Downloads\"
    Debug.Print "Downloads folder path: " & downloadsFolder
    
    ' Create a File Dialog object to select a file
    Set fileDialog = Application.FileDialog(msoFileDialogFilePicker)
    
    ' Set the initial folder to the user's Downloads folder
    fileDialog.InitialFileName = downloadsFolder
    
    ' Set the file dialog filters (only show Excel files)
    fileDialog.Filters.Clear
    fileDialog.Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm"
    
                                                                    '-Add loop for file selection - if file is not selected after second attempt stop script!!!!
    ' Show the dialog box and get the selected file
    If fileDialog.Show = -1 Then
        ' If the user selects a file
        filePath = fileDialog.SelectedItems(1)
        Debug.Print "Selected file: " & filePath
    Else
        ' If the user cancels the selection
        MsgBox "No file selected. The import process has been canceled.", vbExclamation
        Exit Sub
    End If
    
                                                                    '-Redo function!!!!
    ' Check if the file exists
    If Dir(filePath) = "" Then
        MsgBox "The specified file does not exist. Please check the file path.", vbExclamation
        Exit Sub
    End If
    
    ' Show a message that the document is being imported
    MsgBox "Document is being imported...", vbInformation, "Importing"
    
    ' Check for existing ImportedData sheets and generate a new name
    i = 1 ' Start at 1 for ImportedData1
    Do
        sheetName = "ImportedData" & i
        On Error Resume Next ' Disable error handling temporarily
        Set ws = ThisWorkbook.Sheets(sheetName)
        On Error GoTo ErrorHandler ' Re-enable error handling with our custom handler
        If ws Is Nothing Then Exit Do ' If the sheet does not exist, exit loop
        Set ws = Nothing ' Clear the ws object for the next iteration
        i = i + 1
    Loop
    
    ' Create a new worksheet with the unique name
    Set newSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.ActiveSheet)
    newSheet.Name = sheetName
    Debug.Print "Created new sheet: " & newSheet.Name
    
    ' Open the selected Excel file (no validation, just open)
    Set importWorkbook = Workbooks.Open(filePath)
    
                                                                    '-Add loop for file selection - if file is not selected after second attempt stop script!!!!
    ' Check if the workbook opened successfully
    If importWorkbook Is Nothing Then
        MsgBox "Failed to open the file. Please ensure the file is an Excel file and not corrupted.", vbCritical
        Debug.Print "Error: Failed to open file"
        Exit Sub
    End If
    Debug.Print "Imported workbook opened: " & importWorkbook.Name
    
                                                                    '-Redo function!!!!
    ' Copy data from the first sheet of the imported workbook
    importWorkbook.Sheets(1).UsedRange.Copy Destination:=newSheet.Range("A1")
    
                                                                    '-Redo function!!!!
    ' Close the imported workbook (optional, if you don't need it open anymore)
    importWorkbook.Close SaveChanges:=False
    
    ' Display a confirmation message after the import
    MsgBox "Document has been imported!", vbInformation, "Import Complete"
    Exit Sub ' Exit here so that the error handler doesn't trigger when there's no error

ErrorHandler:
    ' Error handling block
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error: " & Err.Number
    Debug.Print "Error number: " & Err.Number & " - " & Err.Description
    Resume Next ' Continue after displaying the error
End Sub
