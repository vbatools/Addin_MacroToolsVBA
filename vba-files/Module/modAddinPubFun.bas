Attribute VB_Name = "modAddinPubFun"
Option Explicit
Option Private Module

#If VBA7 Then
    ' For 64-bit systems
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Enum        : enumParametrVersion
'* Created    : 12-01-2026 13:46
'* Author     : VBATools
'* Copyright  : Apache License
'* Description: Enumeration for selecting application version parameters.
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Enum enumParametrVersion
    enName = 1
    enAuthor
    enVersion
    enLicense
    enDateOfCreation
    enDateOfUpdate
    enDescription
    enAll
    [_First] = enName
    [_Last] = enAll
End Enum

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : RegexEscape - escaping special characters for regular expressions
'* Created    : 12-01-2026 13:46
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                                     Description
'*
'* ByVal sInput As String                         : Source string
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function RegexEscape(ByVal sInput As String) As String
    Dim sSpecial As String, i As Long, sChar As String
    sSpecial = "\^$.|?*+()[]{}"
    For i = 1 To Len(sSpecial)
        sChar = Mid(sSpecial, i, 1)
        sInput = Replace(sInput, sChar, "\" & sChar)
    Next i
    RegexEscape = sInput
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : Version - retrieving information about the application version and parameters
'* Created    : 12-01-2026 13:46
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                                     Description
'*
'* ByVal Parametr As enumParametrVersion          : Requested parameter
'* Optional bOnlyValue As Boolean = False         : Return only the value (without the name)
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function Version(ByVal Parametr As enumParametrVersion, Optional bOnlyValue As Boolean = False) As String
    Dim sRes        As String
    Dim arr         As Variant
    arr = shSettings.ListObjects(TB_ABOUT).DataBodyRange.Value2
    Select Case Parametr
        Case enumParametrVersion.enAll:
            Dim i   As Byte
            For i = enumParametrVersion.[_First] To enumParametrVersion.[_Last] - 1
                If sRes <> vbNullString Then sRes = sRes & vbNewLine
                If arr(i, 3) = 1 Then arr(i, 2) = VBA.Format$(arr(i, 2), FORMAT_DATE)
                sRes = sRes & arr(i, 1) & ": " & arr(i, 2)
            Next i
        Case Else:
            If arr(Parametr, 3) = 1 Then arr(i, 2) = VBA.Format$(arr(Parametr, 2), FORMAT_DATE)
            If bOnlyValue Then
                sRes = arr(Parametr, 2)
            Else
                sRes = arr(Parametr, 1) & ": " & arr(Parametr, 2)
            End If
    End Select
    Version = sRes
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : modPublicFunctions - global public functions
'* Created    : 12-01-2026 13:46
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : getCommentVBATools - returns ASCII art signature for VBATools
'* Created    : 12-01-2026 13:46
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function getCommentVBATools() As String
    Const STR_1     As String = "' __      ______       _______         ,_ "
    Const STR_2     As String = "' \ \    / /  _ \   /\|__   __|        | |"
    Const STR_3     As String = "'  \ \  / /| |_) | /  \  | | ___   ___ | |___"
    Const STR_4     As String = "'   \ \/ / |  _ < / /\ \ | |/ _ \ / _ \| / __|"
    Const STR_5     As String = "'    \  /  | |_) / ____ \| | (_) | (_) | \__ \"
    Const STR_6     As String = "'     \/   |____/_/    \_\_|\___/ \___/|_|___/"
    getCommentVBATools = STR_1 & vbNewLine & STR_2 & vbNewLine & STR_3 & vbNewLine & STR_4 & vbNewLine & STR_5 & vbNewLine & STR_6 & vbNewLine
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Procedure  : URLLinks - opening a link in the default browser
'* Created    : 12-01-2026 13:46
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                                     Description
'*
'* ByVal url_str As String                        : URL address to open
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub URLLinks(ByVal url_str As String)
    If VBA.Len(url_str) = 0 Then Exit Sub
    On Error GoTo ErrorHandler

    Dim appEX       As Object
    Set appEX = CreateObject("Wscript.Shell")
    appEX.Run url_str
    Set appEX = Nothing
    Exit Sub
ErrorHandler:
    Select Case Err
        Case Else:
            Call MsgBox("An error occurred in URLLinks" & vbNewLine & Err.Number & vbNewLine & Err.Description, vbOKOnly + vbCritical, "Error in URLLinks")
    End Select
    Set appEX = Nothing
    Err.Clear
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : FileHave - checks if a file exists
'* Created    : 08-10-2020 13:53
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                                     Description
'*
'* sPath As String                                : - string, path to file or folder
'* Optional Atributes As FileAttribute = vbNormal : - check type, file or folder
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function FileHave(ByVal filePath As String, Optional ByVal fileAttribute As VbFileAttribute = vbNormal) As Boolean
    On Error GoTo ErorHandler
    If filePath = vbNullString Then FileHave = False

    If Right$(filePath, 1) = Application.PathSeparator Then
        filePath = Left$(filePath, Len(filePath) - 1)
    End If

    Select Case fileAttribute
        Case VbFileAttribute.vbDirectory
            FileHave = (Dir$(filePath, vbDirectory) <> vbNullString)
        Case VbFileAttribute.vbNormal
            FileHave = (Dir$(filePath, vbNormal) <> vbNullString)
        Case Else
            Err.Raise vbObjectError + 1, "FileHave", "Unsupported file attribute: " & fileAttribute
    End Select

    GoTo CleanExit

ErorHandler:
    Err.Raise Err.Number, "FileHave", Err.Description

CleanExit:
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : sGetBaseName - extracting the file name without extension
'* Created    : 12-01-2026 13:46
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                                     Description
'*
'* ByVal filePath As String                       : Full path to the file
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function sGetBaseName(ByVal filePath As String) As String
    On Error GoTo ErrorHandler
    
    If filePath = vbNullString Then Exit Function

    Dim separator As String
    separator = Application.PathSeparator
    
    Dim fileName As String
    If InStrRev(filePath, separator) > 0 Then
        fileName = Mid$(filePath, InStrRev(filePath, separator) + 1)
    Else
        fileName = filePath
    End If
    
    Dim dotPosition As Long
    dotPosition = InStrRev(fileName, ".")
    
    ' If dot is found and it is not the first character (hidden files in Linux, for example)
    If dotPosition > 1 Then
        sGetBaseName = Left$(fileName, dotPosition - 1)
    Else
        ' If there is no dot or it is at the beginning, return the name as is
        sGetBaseName = fileName
    End If
    
    Exit Function

ErrorHandler:
    ' Re-throw error to the calling code
    Err.Raise Err.Number, "sGetBaseName", Err.Description
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : sGetExtensionName - retrieving the file extension
'* Created    : 12-01-2026 13:46
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                                     Description
'*
'* ByVal filePath As String                       : Full path to the file
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function sGetExtensionName(ByVal filePath As String) As String
    On Error GoTo ErrorHandler
    
    If filePath = vbNullString Then Exit Function
    
    Dim separator As String
    separator = Application.PathSeparator
    
    Dim fileName As String
    If InStrRev(filePath, separator) > 0 Then
        fileName = Mid$(filePath, InStrRev(filePath, separator) + 1)
    Else
        fileName = filePath
    End If
    
    Dim dotPosition As Long
    dotPosition = InStrRev(fileName, ".")
    
    ' If dot is found and it is not the first character (e.g., file ".gitignore")
    If dotPosition > 1 Then
        ' Return text after the dot in lower case (standard for extensions)
        sGetExtensionName = LCase$(Mid$(fileName, dotPosition + 1))
    Else
        ' If there is no dot or it is at the beginning, return an empty string
        sGetExtensionName = vbNullString
    End If
    
    Exit Function

ErrorHandler:
    ' Re-throw error to the calling code
    Err.Raise Err.Number, "sGetExtensionName", Err.Description
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : sGetFileName - retrieving the file name with extension
'* Created    : 12-01-2026 13:46
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                                     Description
'*
'* ByVal filePath As String                       : Full path to the file or folder
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function sGetFileName(ByVal filePath As String) As String
    On Error GoTo ErrorHandler
    
    ' 1. Input validation
    If filePath = vbNullString Then Exit Function

    Dim separator As String
    separator = Application.PathSeparator
    
    Dim separatorPosition As Long
    separatorPosition = InStrRev(filePath, separator)
    
    ' 3. Extracting file name
    If separatorPosition > 0 Then
        ' Return everything after the last slash
        sGetFileName = Mid$(filePath, separatorPosition + 1)
    Else
        ' If there is no separator, the file name itself was passed
        sGetFileName = filePath
    End If
    
    Exit Function

ErrorHandler:
    ' Re-throw error to the calling code
    Err.Raise Err.Number, "sGetFileName", Err.Description
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : sGetParentFolderName - retrieving the path to the parent folder
'* Created    : 12-01-2026 13:46
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                                     Description
'*
'* ByVal filePath As String                       : Full path to the file or folder
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function sGetParentFolderName(ByVal filePath As String) As String
    On Error GoTo ErrorHandler
    
    ' 1. Input validation
    If filePath = vbNullString Then Exit Function
    
    Dim separator As String
    separator = Application.PathSeparator
    
    Dim separatorPosition As Long
    separatorPosition = InStrRev(filePath, separator)
    
    ' 3. Extracting the path to the parent directory
    If separatorPosition > 1 Then
        ' Return path without trailing slash and file name
        sGetParentFolderName = Left$(filePath, separatorPosition - 1)
    Else
        ' If there is no separator or it is the first (disk root), return an empty string
        sGetParentFolderName = vbNullString
    End If
    
    Exit Function

ErrorHandler:
    ' Re-throw error to the calling code
    Err.Raise Err.Number, "sGetParentFolderName", Err.Description
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : MoveFile - moving or renaming a file
'* Created    : 12-01-2026 13:46
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                                     Description
'*
'* ByVal sourcePath As String                     : Full path to the source file
'* ByVal destinationPath As String                : Full destination path
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function MoveFile(ByVal sourcePath As String, ByVal destinationPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' 1. Input validation
    If sourcePath = vbNullString Or destinationPath = vbNullString Then
        Exit Function
    End If
    
    ' 2. Checking if the source file exists (vbNormal - regular files)
    If Dir$(sourcePath, vbNormal) = vbNullString Then
        Exit Function
    End If
    
    ' 3. Moving the file using built-in VBA statement
    ' The Name statement moves the file at the file system level (faster than FSO)
    Name sourcePath As destinationPath
    
    ' 4. Successful completion
    MoveFile = True
    Exit Function

ErrorHandler:
    ' In case of an error (no rights, file is locked, folder does not exist) return False
    MoveFile = False
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : WorkbookIsOpen - checking if an Excel workbook is open
'* Created    : 12-01-2026 13:46
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                                     Description
'*
'* ByVal workbookName As String                   : Workbook name (e.g., "Report.xlsx")
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function WorkbookIsOpen(ByVal workbookName As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' 1. Input validation
    If workbookName = vbNullString Then Exit Function

    
    ' 2. Iterating through open workbooks (safe check without On Error Resume Next)
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.Name, workbookName, vbTextCompare) = 0 Then
            WorkbookIsOpen = True
            Exit Function
        End If
    Next wb
    
    ' If the loop finished and the workbook was not found
    WorkbookIsOpen = False
    Exit Function

ErrorHandler:
    ' Re-throw error to the calling code
    Err.Raise Err.Number, "WorkbookIsOpen", Err.Description
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : IsTableExists - checking if a table exists on a worksheet
'* Created    : 12-01-2026 13:46
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                                     Description
'*
'* ByVal sh As Worksheet                          : Worksheet object to check
'* ByVal sTBName As String                        : Table name (ListObject)
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function IsTableExists(ByVal sh As Worksheet, ByVal sTBName As String) As Boolean
    On Error Resume Next
    IsTableExists = Not sh.ListObjects(sTBName) Is Nothing
    On Error GoTo 0
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Procedure  : base64ToFile - decoding a Base64 string to a file
'* Created    : 12-01-2026 13:46
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                                     Description
'*
'* ByVal sHashBase64 As String                    : Base64 formatted string
'* ByVal sFilePath As String                      : Full path to save the file
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub base64ToFile(ByVal sHashBase64 As String, ByVal sFilePath As String)
    Dim byteArr()   As Byte
    Dim oBase       As Object
    Set oBase = CreateObject("MSXML2.DOMDocument").createElement("b64")
    With oBase
        .DataType = "bin.base64"
        .text = sHashBase64
        byteArr = .nodeTypedValue
    End With
    Open sFilePath For Binary Access Write As #1
    Put #1, 1, byteArr
    Close #1
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function     :   addTabelFromArray - convert array to markdown table
'* Author       :   VBATools
'* Copyright    :   Apache License
'* Created      :   22-01-2026 10:24
'* Argument(s)  :               Description
'*
'* ByRef arr As Variant       :
'* ByVal sDelimetr As String  :
'* ByVal haveHeder As Boolean :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function addTabelFromArray(ByRef arr As Variant, ByVal sDelimetr As String, ByVal haveHeder As Boolean, ByVal iEndCol As Byte) As String
    Dim i           As Long
    Dim iCount      As Long
    Dim j           As Long
    Dim k           As Long
    Dim iMaxLen     As Long
    Dim Shift       As Byte
    Dim shift_2     As Byte

    iCount = UBound(arr, 1)
    If LBound(arr, 1) = 1 Then
        Shift = 1
    Else
        shift_2 = 1
    End If

    If haveHeder Then
        ReDim arrCol(0 To iCount + shift_2) As Variant
        ReDim arrLen(0 To iCount + Shift + shift_2) As Integer
    Else
        ReDim arrCol(0 To iCount - Shift) As Variant
        ReDim arrLen(0 To iCount + shift_2) As Integer
    End If

    If iEndCol = 0 Then iEndCol = UBound(arr, 2)
    iEndCol = iEndCol - shift_2

    For j = LBound(arr, 2) To iEndCol
        For i = LBound(arr, 1) To iCount
            k = k + 1
            If j > Shift Then arrCol(k - 1) = arrCol(k - 1) & VBA.String$(arrLen(iCount + shift_2) - arrLen(i - Shift), " ") & " " & sDelimetr & " "
            If j = Shift Then arrCol(k - 1) = " " & sDelimetr & " " & arrCol(k - 1)

            arrCol(k - 1) = arrCol(k - 1) & arr(i, j)
            arrLen(i - Shift) = VBA.Len(arr(i, j))
            If arrLen(i - Shift) > iMaxLen Then iMaxLen = arrLen(i - Shift)
            If haveHeder And i = Shift Then
                k = k + 1
                arrCol(k - 1) = arrCol(k - 1) & VBA.String$(arrLen(iCount + shift_2), "-") & " " & sDelimetr & " "
            End If
        Next i
        arrLen(iCount + shift_2) = iMaxLen
        iMaxLen = 0: k = 0
    Next j
    For i = LBound(arr, 1) To iCount
        k = k + 1
        arrCol(k - 1) = arrCol(k - 1) & VBA.String$(arrLen(iCount + shift_2) - arrLen(i - Shift), " ") & " " & sDelimetr & " "
        If haveHeder And i = Shift Then
            k = k + 1
            arrCol(k - 1) = arrCol(k - 1) & VBA.String$(arrLen(iCount + shift_2), "-") & " " & sDelimetr & " "
        End If
    Next i
    addTabelFromArray = VBA.Join(arrCol, vbNewLine)
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : getPath - calling the folder picker dialog
'* Created    : 12-01-2026 13:46
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                                     Description
'*
'* ByVal sTitle As String                         : Dialog window title
'* ByVal sPath As String                          : Initial path
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function getPath(ByVal sTitle As String, ByVal sPath As String) As String
    On Error GoTo OptB_UsePatch_Change_Err
    If sPath = vbNullString Or Not (Dir(sPath, vbDirectory) <> vbNullString) Then sPath = ThisWorkbook.Path
    Dim sGetFolder  As String
    With Application.FileDialog(msoFileDialogFolderPicker)        ' show dialog window
        .ButtonName = "Выбрать": .Title = sTitle: .InitialFileName = sPath
        If .Show <> -1 Then
            getPath = sPath
            Exit Function    ' if the user canceled the folder selection
        End If
        sGetFolder = .SelectedItems(1)
    End With
    getPath = sGetFolder & Application.PathSeparator
    Exit Function
OptB_UsePatch_Change_Err:
    MsgBox Err.Description, vbExclamation + vbOKOnly, "Error:"
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : fileDialogFun - calling the file picker dialog
'* Created    : 12-01-2026 13:46
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                                     Description
'*
'* ByVal sPath As String                          : Initial path
'* ByRef bMultiSelect As Boolean                  : Allow multi-select
'* Optional sExpansion As String = "*.xlsm;..."   : File extension filter
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function fileDialogFun(ByVal sPath As String, _
        ByRef bMultiSelect As Boolean, _
        Optional sExpansion As String = "*.xlsm;*.xlsb;*.xlsx") As String()

    If sPath = vbNullString Or Not (Dir(sPath, vbDirectory) <> vbNullString) Then sPath = ThisWorkbook.Path
    If sPath = vbNullString Then sPath = Environ("USERPROFILE") & Application.PathSeparator & "Desktop"

    Dim oFd         As FileDialog
    Set oFd = Application.FileDialog(msoFileDialogFilePicker)
    With oFd
        .AllowMultiSelect = bMultiSelect
        'dialog window title
        .Title = "Select Files:"
        'clear previously set file types
        .Filters.Clear
        'set the ability to select only Excel files
        .Filters.Add "Microsoft Excel Files", sExpansion, 1
        'assign the display folder and default file name
        .InitialFileName = sPath
        'dialog window view (9 variants available)
        .InitialView = msoFileDialogViewDetails
        If .Show = 0 Then
            Call MsgBox("No files selected!", vbCritical, "Select Files:")
            Exit Function
        End If
        Dim iCount  As Integer
        Dim i       As Integer
        iCount = .SelectedItems.Count
        'ReDim arr(1 To iCount) As String
        ReDim arr(1 To iCount, 1 To 1) As String
        For i = 1 To iCount
            'arr(i) = VBA.CStr(.SelectedItems.Item(i))
            arr(i, 1) = VBA.CStr(.SelectedItems.Item(i))
        Next
    End With
    fileDialogFun = arr

    'For procedure
    'If (Not (Not (v))) = 0 Then Exit Sub
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : GetFilesTable - retrieving a table of files from a folder
'* Created    : 12-01-2026 13:46
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                                     Description
'*
'* ByVal folderPath As String                     : Path to the target folder
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Function GetFilesTable(ByVal folderPath As String) As Variant
    Dim FSO         As Object
    Dim folder      As Object
    Dim fileCount   As Long
    Dim varResult   As Variant
    Dim rowIndex    As Long

    ' Create FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")

    ' 1. Check if folder exists
    If Not FSO.FolderExists(folderPath) Then
        MsgBox "Specified folder does not exist:" & folderPath, vbExclamation, "Error"
        GetFilesTable = Array()    ' Return empty array
        Exit Function
    End If

    Set folder = FSO.GetFolder(folderPath)

    ' 2. First pass: Count files
    fileCount = CountFilesRecursive(folder)

    ' If no files found
    If fileCount = 0 Then
        GetFilesTable = Array()
        Exit Function
    End If

    ' 3. Initialize two-dimensional array
    ' Rows: from 1 to fileCount
    ' Columns: from 1 to 4 (Path, Name, Size, Date)
    ReDim varResult(1 To fileCount, 1 To 4)

    ' 4. Second pass: Fill array
    rowIndex = 1    ' Start filling from the first row
    Call FillFilesTableRecursive(folder, varResult, rowIndex)

    ' Return filled table
    GetFilesTable = varResult

    ' Free memory
    Set folder = Nothing
    Set FSO = Nothing
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : CountFilesRecursive - recursive file count in a folder
'* Created    : 12-01-2026 13:46
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                                     Description
'*
'* ByVal currentFolder As Object                  : Folder object (FileSystemObject)
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function CountFilesRecursive(ByVal currentFolder As Object) As Long
    Dim subFolder   As Object
    Dim lCount      As Long

    lCount = currentFolder.Files.Count

    For Each subFolder In currentFolder.SubFolders
        lCount = lCount + CountFilesRecursive(subFolder)
    Next subFolder

    CountFilesRecursive = lCount
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Procedure  : FillFilesTableRecursive - recursive filling of an array with file data
'* Created    : 12-01-2026 13:46
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                                     Description
'*
'* ByVal currentFolder As Object                  : Folder object (FileSystemObject)
'* ByRef tableArray As Variant                    : Array to fill
'* ByRef currentRow As Long                       : Current row for writing
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub FillFilesTableRecursive(ByVal currentFolder As Object, ByRef tableArray As Variant, ByRef currentRow As Long)
    Dim subFolder   As Object
    Dim file        As Object

    ' Fill data for each file in the current folder
    For Each file In currentFolder.Files
        tableArray(currentRow, 1) = file.Name                ' File name
        tableArray(currentRow, 2) = file.Path                ' Full path
        tableArray(currentRow, 3) = file.Size                ' Size (bytes)
        tableArray(currentRow, 4) = file.DateLastModified    ' Modification date

        currentRow = currentRow + 1
    Next file

    ' Recursive call for subfolders
    For Each subFolder In currentFolder.SubFolders
        Call FillFilesTableRecursive(subFolder, tableArray, currentRow)
    Next subFolder
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : GetArrayFromDictionary - converting dictionary data to a 2D array
'* Created    : 12-01-2026 13:46
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                                     Description
'*
'* ByRef oDic As Dictionary                       : Dictionary object
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function GetArrayFromDictionary(ByRef oDic As Dictionary) As Variant
    If oDic Is Nothing Then Exit Function
    If oDic.Count = 0 Then Exit Function

    Dim lRowCount   As Long
    Dim lColCount   As Long
    Dim arrRes()    As String
    Dim arrItems    As Variant
    Dim arrItem     As Variant
    Dim i As Long, j As Long

    lRowCount = oDic.Count
    arrItems = oDic.Items

    On Error Resume Next
    lColCount = UBound(arrItems(0), 2)
    On Error GoTo 0

    If lColCount = 0 Then Exit Function
    ReDim arrRes(1 To lRowCount, 1 To lColCount)
    For i = 0 To lRowCount - 1
        arrItem = arrItems(i)

        For j = 1 To lColCount
            arrRes(i + 1, j) = arrItem(1, j)
        Next j
    Next i
    GetArrayFromDictionary = arrRes
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Procedure  : OutputResults - outputting an array of data to a workbook sheet
'* Created    : 12-01-2026 13:46
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                                     Description
'*
'* ByRef wb As Workbook                          : Target workbook
'* ByRef sSheetName As String                    : Name of the sheet for output
'* ByRef sNameVBAProj As String                  : Project name
'* ByRef arrOutput As Variant                    : Array of data to output
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub OutputResults(ByRef wb As Workbook, ByRef sSheetName As String, ByRef sNameVBAProj As String, ByRef arrOutput As Variant)
    If IsEmpty(arrOutput) Then Exit Sub
    Dim wsTarget    As Worksheet

    On Error Resume Next
    Set wsTarget = wb.Worksheets(sSheetName)
    On Error GoTo 0

    If wsTarget Is Nothing Then
        Set wsTarget = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        wsTarget.Name = sSheetName
    Else
        wsTarget.Cells.ClearContents
    End If

    Application.ScreenUpdating = False
    Dim iCol        As Integer
    Dim iRow        As Integer
    Dim i           As Integer

    On Error Resume Next
    iCol = UBound(arrOutput, 2)
    If LBound(arrOutput, 2) = 0 Then iCol = iCol + 1
    If Err.Number <> 0 Then iCol = 1
    On Error GoTo 0

    iRow = UBound(arrOutput, 1)
    If LBound(arrOutput, 1) = 0 Then iRow = iRow + 1
    With wsTarget
        If isArray(arrOutput) Then
            .Cells(3, 1).Resize(iRow, iCol).Value2 = arrOutput
            Rows("3:" & iRow + 3).RowHeight = 15
        End If
        For i = 1 To iCol
            With .Cells(1, i)
                .EntireColumn.AutoFit
                If .ColumnWidth > 50 Then .ColumnWidth = 50
            End With
            .Cells(2, i).Value2 = i
        Next i
        .Cells(1, 1).value = sNameVBAProj

        .Range(.Cells(2, 1), .Cells(2, iCol)).AutoFilter
        .Activate
    End With
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Procedure  : WriteErrorLog - recording and displaying error logs
'* Created    : 12-01-2026 13:46
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                                     Description
'*
'* ByVal sNameFunc As String                      : Name of the function that caused the error
'* ByVal bMsgBox As Boolean                       : Show MsgBox (True) or print to Debug
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub WriteErrorLog(ByVal sNameFunc As String, ByVal bMsgBox As Boolean)

    If bMsgBox Then
        Call MsgBox("Error! In " & sNameFunc & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "at line" & Erl, vbOKOnly + vbExclamation, "Critical Error:")
    Else
        Debug.Print ">> Error! In " & sNameFunc & vbCrLf & Err.Number & vbCrLf & Err.Description & vbCrLf & "at line" & Erl
    End If

    Dim clsLoger    As clsLogging
    Set clsLoger = New clsLogging
    Call clsLoger.LogError(sNameFunc)
    Set clsLoger = Nothing
End Sub

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : FilterArrayByText - filtering an array by text value
'* Created    : 12-01-2026 13:46
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                                     Description
'*
'* ByRef vSourceArray As Variant                  : Source data array
'* ByVal iColSerch As Byte                        : Index of the column to search
'* ByVal sSearch As String                        : Search string
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function FilterArrayByText(ByRef vSourceArray As Variant, ByVal iColSerch As Byte, ByVal sSearch As String) As Variant
    Dim i           As Long
    Dim lCount      As Long
    Dim lRowsStart  As Long
    Dim lRowsEnd    As Long
    Dim lColsStart  As Long
    Dim lColsEnd    As Long
    Dim arrResult() As Variant

    sSearch = VBA.UCase$(sSearch)
    lRowsStart = LBound(vSourceArray, 1)
    lRowsEnd = UBound(vSourceArray, 1)
    lColsStart = LBound(vSourceArray, 2)
    lColsEnd = UBound(vSourceArray, 2)

    ' 1. Count matches for accurate memory allocation
    For i = lRowsStart To lRowsEnd
        If VBA.InStr(1, VBA.UCase$(CStr(vSourceArray(i, iColSerch))), sSearch) > 0 Then
            lCount = lCount + 1
        End If
    Next i

    ' If nothing found, return Empty
    If lCount = 0 Then
        FilterArrayByText = Empty
        Exit Function
    End If

    ' 2. Fill result array
    If lRowsStart = 0 Then lCount = lCount - 1
    ReDim arrResult(lRowsStart To lCount, lColsStart To lColsEnd)
    lCount = 0

    Dim j           As Long
    lCount = lRowsStart
    For i = lRowsStart To lRowsEnd
        If VBA.InStr(1, VBA.UCase$(CStr(vSourceArray(i, iColSerch))), sSearch) > 0 Then
            For j = lColsStart To lColsEnd
                arrResult(lCount, j) = vSourceArray(i, j)
            Next j
            lCount = lCount + 1
        End If
    Next i

    FilterArrayByText = arrResult
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : GetTargetWorkbook - selecting a target workbook via a user form
'* Created    : 12-01-2026 13:46
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                                     Description
'*
'* ByRef wb As Workbook                          : Selected workbook object
'* ByVal sCaptionForm As String                  : Form title
'* ByVal sCaptionBtnOk As String                 : Confirmation button text
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function GetTargetWorkbook(ByRef wb As Workbook, ByVal sCaptionForm As String, ByVal sCaptionBtnOk As String) As Boolean
    With frmListWBOpen
        .Caption = sCaptionForm
        .lbOK.Caption = sCaptionBtnOk
        .Show

        If .cmbMain.value = vbNullString Or .lbRes.Caption = "0" Then Exit Function
        On Error Resume Next
        Set wb = Workbooks(.cmbMain.value)
        If Err.Number <> 0 Then Exit Function
        On Error GoTo 0
    End With
    GetTargetWorkbook = True
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : saveTextToFile - saving text to a file with a specified encoding
'* Created    : 12-01-2026 13:46
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                                     Description
'*
'* ByVal txt As String                           : Text to save
'* ByVal fileName As String                      : Full path to the file
'* Optional ByVal encoding As String = "win..."  : Encoding (utf-8, utf-16, ansi)
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function saveTextToFile(ByVal txt As String, ByVal fileName As String, Optional ByVal encoding As String = "windows-1251") As Boolean

    On Error Resume Next: Err.Clear
    Dim FSO         As Object
    Dim ts          As Object
    Select Case encoding
        Case "windows-1251", "", "ansi"
            Set FSO = CreateObject("scripting.filesystemobject")
            Set ts = FSO.CreateTextFile(fileName, True)
            ts.Write txt: ts.Close
            Set ts = Nothing: Set FSO = Nothing

        Case "utf-16", "utf-16LE"
            Set FSO = CreateObject("scripting.filesystemobject")
            Set ts = FSO.CreateTextFile(fileName, True, True)
            ts.Write txt: ts.Close
            Set ts = Nothing: Set FSO = Nothing

        Case "utf-8noBOM"
            Dim binaryStream As Object
            With CreateObject("ADODB.Stream")
                .Type = 2: .Charset = "utf-8": .Open
                .WriteText txt$

                Set binaryStream = CreateObject("ADODB.Stream")
                binaryStream.Type = 1: binaryStream.mode = 3: binaryStream.Open
                .Position = 3: .CopyTo binaryStream
                .flush: .Close
                binaryStream.SaveToFile fileName$, 2
                binaryStream.Close
            End With

        Case Else
            With CreateObject("ADODB.Stream")
                .Type = 2: .Charset = encoding$: .Open
                .WriteText txt$
                .SaveToFile fileName$, 2
                .Close
            End With
    End Select
    saveTextToFile = Err = 0: DoEvents
End Function

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : CollapseRepeats - removing repeating characters in a string
'* Created    : 12-01-2026 13:46
'* Author     : VBATools
'* Copyright  : Apache License
'* Argument(s):                                     Description
'*
'* ByVal strInput As String                       : Source string
'* ByVal strChar As String                        : Character to collapse
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function CollapseRepeats(strInput As String, strChar As String) As String
    If Len(strInput) = 0 Then Exit Function
    
    Dim i As Long
    Dim sDoubleChar As String
    sDoubleChar = strChar & strChar
    
    i = VBA.InStr(1, strInput, sDoubleChar)
    Do While i > 0
        strInput = VBA.Replace(strInput, sDoubleChar, strChar)
        i = VBA.InStr(1, strInput, sDoubleChar)
    Loop

    CollapseRepeats = strInput
End Function

