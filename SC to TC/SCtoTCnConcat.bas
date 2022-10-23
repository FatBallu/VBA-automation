Attribute VB_Name = "SCtoTCnConcat"
Sub SCtoTCnConcat()
''Simplified Chinese to Traditional Chinese with concatenation into a single file
'file names must be have certain pattern with only difference in numbers
'such as file1.doc, file2.doc; or name _ (4).txt, name _ (5).txt
'content will be opened with unicode UTF-8 (Encoding:=65001)
'concatenated file will be remained opened for modification and unsaved

    Dim NamePattern As String
    Dim DialogFolder, Dialog1File, DialogConcat As FileDialog
    Dim FileCurrent, PatternGet, FileConcat As String
    Dim FileNum, FileMax As Integer
    Application.ScreenUpdating = False

    'select one of the SC files
    Set Dialog1File = Application.FileDialog(msoFileDialogFilePicker)
    If Dialog1File.Show <> -1 Then Exit Sub
    PatternGet = Dialog1File.SelectedItems(1)
    Documents.Open FileName:=PatternGet
    ActiveDocument.Close
    NamePattern = PatternGet
    'getting name pattern of files
    'the number "6" here is determined by the extension of file
    'e.g. file2.doc counts 5 (from "2" to "c")
    'e.g. files(12).txt counts 7 (from "1" to "t")
    NamePattern = Left(NamePattern, Len(NamePattern) - 5)
    
    'select output file
    'concatenating in order
    Set DialogConcat = Application.FileDialog(msoFileDialogFilePicker)
    If DialogConcat.Show <> -1 Then Exit Sub
    FileConcat = DialogConcat.SelectedItems(1)
    
    'number of files: counts from FileMin to FileMax
    FileMin = 8
    Counter = FileMin
    FileMax = 16
    
    'in the concatenaing file
    'move to the end of the file
    Selection.EndKey Unit:=wdStory
    Selection.TypeParagraph

    While Counter <= FileMax
    'the text in ").txt" must match the pattern after the counter number
    'such as ".txt" in some cases
        FileCurrent = NamePattern & Counter & ").txt"
        Documents.Open FileName:=FileCurrent, _
            ReadOnly:=False, Format:=wdOpenFormatAuto, Encoding:=65001

        Selection.WholeStory
        'translate FileCurrent
        Selection.Range.TCSCConverter WdTCSCConverterDirection:= _
        wdTCSCConverterDirectionSCTC
        Selection.WholeStory
        Selection.Copy
        ActiveWindow.Close Savechanges:=False
        'back to the concatenating file
        Documents.Open FileName:=FileConcat
        'go to the end of the concatenating file
        Set Range0 = ActiveDocument.Content
        Range0.Collapse Direction:=wdCollapseEnd
        Range0.Paste
        'increment of counter
        Counter = Counter + 1
    Wend
    Application.ScreenUpdating = True
End Sub

