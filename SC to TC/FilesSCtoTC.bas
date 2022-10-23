Attribute VB_Name = "FilesSCtoTC"
Sub FilesSCtoTC()
''Simplified Chinese to Traditional Chinese for all files in a folder
'subfolder will not be read
'output to specified folder

    Dim NewName As String
    Dim DialogInFolder, DialogOutFolder As FileDialog
    Dim xFolder As Variant
    Dim FileCurrent, InputPath, OutputPath As String
    Application.ScreenUpdating = False
    'select folder containing input files
    Set DialogInFolder = Application.FileDialog(msoFileDialogFolderPicker)
    If DialogInFolder.Show <> -1 Then Exit Sub
    InputPath = DialogInFolder.SelectedItems(1) + "\"
    FileCurrent = Dir(InputPath & "*.txt", vbNormal)
    
    'select folder for output files
    Set DialogOutFolder = Application.FileDialog(msoFileDialogFolderPicker)
    If DialogOutFolder.Show <> -1 Then Exit Sub
    OutputPath = DialogOutFolder.SelectedItems(1) + "\"
    While FileCurrent <> ""
        Documents.Open FileName:=InputPath & FileCurrent, _
            ReadOnly:=False, Format:=wdOpenFormatAuto, Encoding:=65001
        NewName = FileCurrent
        'getting name pattern of files
        'the number "6" here is determined by the extension of file
        'e.g. file2.docx counts 4 (from "d" to "x")
        'e.g. files(12).doc counts 3 (from "d" to "c")
        NewName = Left(NewName, Len(NewName) - 3)
        'translate FileCurrent
        Selection.WholeStory
        Selection.Range.TCSCConverter WdTCSCConverterDirection:= _
            wdTCSCConverterDirectionSCTC
        'save as new file with NewName in .txt extention
        ActiveDocument.SaveAs2 FileName:=OutputPath + NewName + "txt", _
            FileFormat:=wdFormatText, Encoding:=65001, _
            LineEnding:=wdCRLF, CompatibilityMode:=0
        ActiveDocument.Close
        'loop to next file in the list of files
        FileCurrent = Dir()
    Wend
    Application.ScreenUpdating = True
End Sub


