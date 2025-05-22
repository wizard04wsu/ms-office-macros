Attribute VB_Name = "Files"
'requires reference to Microsoft Office 14.0 Object Library

'for changing the current directory to a network drive
'see https://stackoverflow.com/questions/42366047/excel-macro-open-file-on-network-directory#42368253
'see https://www.pcreview.co.uk/threads/path-names-in-getopenfilename.965570/#post-2820465
Public Declare PtrSafe Function SetCurrentDirectoryA Lib "kernel32" (ByVal lpPathName As String) As Long

'File/folder picker dialog box
'see https://msdn.microsoft.com/en-us/VBA/Office-Shared-VBA/articles/filedialog-object-office
'see https://msdn.microsoft.com/en-us/VBA/Office-Shared-VBA/articles/filedialog-members-office
'see https://msdn.microsoft.com/en-us/vba/office-shared-vba/articles/msofiledialogtype-enumeration-office

'@param pathAndOrFileFilter   specifies what directory to open the dialog box in, and/or what files to display (e.g., "C:\users\" or "*.xls*" or "C:\users\me\desktop\*.txt")
'@return   if user clicked Cancel, returns an empty string
Public Function GetFolderName(Optional pathAndOrFileFilter As String) As String
    Dim buttonClicked As Long
    With Application.FileDialog(msoFileDialogFolderPicker)
        .initialFileName = pathAndOrFileFilter
        buttonClicked = .Show
        If .SelectedItems.Count = 0 Then Exit Function 'user clicked Cancel
        GetFolderName = .SelectedItems(1)
    End With
End Function

'@param pathAndOrFileFilter   specifies what directory to open the dialog box in, and/or what type of files to display (e.g., "C:\" or "*.xls*" or "C:\*.txt")
'@return   if user clicked Cancel, returns an empty string
Public Function GetFileName(Optional pathAndOrFileFilter As String) As String
    Dim buttonClicked As Long
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .initialFileName = pathAndOrFileFilter
        .Show
        If .SelectedItems.Count = 0 Then Exit Function 'user clicked Cancel
        GetFileName = .SelectedItems(1)
    End With
End Function

'@return   True if successful
Public Function ChangeDirectory(path As String) As Boolean
    On Error Resume Next
    ChDir path
    On Error GoTo 0
    If Err <> 0 Then    'might be a network drive
        If SetCurrentDirectoryA(path) = 0 Then    'not found
            ChangeDirectory = False
            Exit Function
        End If
    End If
    ChangeDirectory = True
End Function

'@param fileFilter   specifies what files to count (e.g., "*.xls*")
Public Function GetFileCount(thePath As String, Optional fileFilter As String = "*.*") As Long
    
    Dim theFile As String
    Dim theDir As String
    Dim sDirList As String: sDirList = ""
    Dim arDirList() As String
    Dim i As Long: i = 1
    
    GetFileCount = 0
    
    If ChangeDirectory(thePath) Then
        
        theFile = Dir(fileFilter)
        Do While theFile <> ""  'for each file in this folder
            'full path is: thePath & "\" & theFile
            GetFileCount = GetFileCount + 1
            theFile = Dir
        Loop
        
        theDir = Dir("*.", vbDirectory)
        Do While theDir <> ""   'for each subdirectory
            If theDir <> "." And theDir <> ".." Then sDirList = sDirList & ";" & theDir 'add it to the list
            theDir = Dir
        Loop
        arDirList = Split(sDirList, ";")    'convert the subdirectory list to an array
        Do While i <= UBound(arDirList) 'for each subdirectory
            GetFileCount = GetFileCount + GetFileCount(thePath & "\" & arDirList(i))   'recurse
            i = i + 1
        Loop
        
    End If
    
End Function



