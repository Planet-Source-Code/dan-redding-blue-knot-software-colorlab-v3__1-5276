Attribute VB_Name = "modOpenSaveDialog"
Option Explicit

Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Const GWL_HINSTANCE = (-6)
Const SWP_NOSIZE = &H1
Const SWP_NOZORDER = &H4
Const SWP_NOACTIVATE = &H10
Const HCBT_ACTIVATE = 5
Const WH_CBT = 5

Dim hHook As Long

Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ENABLETEMPLATEHANDLE = &H80
Public Const OFN_EXPLORER = &H80000
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_LONGNAMES = &H200000
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_NODEREFERENCELINKS = &H100000
Public Const OFN_NOLONGNAMES = &H40000
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_READONLY = &H1
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_SHAREFALLTHROUGH = 2
Public Const OFN_SHAREWARN = 0
Public Const OFN_SHARENOWARN = 1
Public Const OFN_SHOWHELP = &H10
Public Const OFS_MAXPATHNAME = 256

Public Const LF_FACESIZE = 32

'OFS_FILE_OPEN_FLAGS and OFS_FILE_SAVE_FLAGS below
'are mine to save long statements; they're not
'a standard Win32 type.
Public Const OFS_FILE_OPEN_FLAGS = _
    OFN_EXPLORER Or OFN_LONGNAMES Or OFN_CREATEPROMPT Or OFN_NODEREFERENCELINKS Or OFN_HIDEREADONLY
Public Const OFS_FILE_SAVE_FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or _
    OFN_OVERWRITEPROMPT Or OFN_HIDEREADONLY

Public Type OPENFILENAME
    nStructSize As Long
    hwndOwner As Long
    hInstance As Long
    sFilter As String
    sCustomFilter As String
    nCustFilterSize As Long
    nFilterIndex As Long
    sFile As String
    nFileSize As Long
    sFileTitle As String
    nTitleSize As Long
    sInitDir As String
    sDlgTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExt As Integer
    sDefFileExt As String
    nCustDataSize As Long
    fnHook As Long
    sTemplateName As String
End Type

Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
End Type

Type OFNOTIFY
        hdr As NMHDR
        lpOFN As OPENFILENAME
        pszFile As String        '  May be NULL
End Type

Public Const LBSELCHSTRING = "commdlg_LBSelChangedNotify"
Public Const SHAREVISTRING = "commdlg_ShareViolation"
Public Const FILEOKSTRING = "commdlg_FileNameOK"
Public Const COLOROKSTRING = "commdlg_ColorOK"
Public Const SETRGBSTRING = "commdlg_SetRGBColor"
Public Const HELPMSGSTRING = "commdlg_help"
Public Const FINDMSGSTRING = "commdlg_FindReplace"


Public Type SelectedFile
    nFilesSelected As Integer
    sFiles() As String
    sLastDirectory As String
    bCanceled As Boolean
End Type


Public FileDialog As OPENFILENAME
Dim ParenthWnd As Long

Public Function ShowOpen(ByVal hwnd As Long, Optional ByVal centerForm As Boolean = True, Optional strFile As String = "", Optional strDir As String = "") As SelectedFile
Dim ret As Long
Dim Count As Integer
Dim fileNameHolder As String
Dim LastCharacter As Integer
Dim NewCharacter As Integer
Dim tempFiles(1 To 200) As String
Dim hInst As Long
Dim Thread As Long
    
    If strDir = "" Then strDir = App.Path
    ParenthWnd = hwnd
    FileDialog.nStructSize = Len(FileDialog)
    FileDialog.hwndOwner = hwnd
    FileDialog.sFileTitle = Space$(2048)
    FileDialog.nTitleSize = Len(FileDialog.sFileTitle)
    FileDialog.sFile = Space$(2047) & Chr$(0) 'FileDialog.sFile &
    Mid$(FileDialog.sFile, 1, Len(strFile)) = strFile
    FileDialog.sInitDir = strDir & Chr$(0)

    FileDialog.nFileSize = Len(FileDialog.sFile)
    
    'If FileDialog.flags = 0 Then
        FileDialog.flags = OFS_FILE_OPEN_FLAGS
    'End If
    
    'Set up the CBT hook
    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    If centerForm = True Then
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If
    
    ret = GetOpenFileName(FileDialog)

    If ret Then
        If Trim$(FileDialog.sFileTitle) = "" Then
            LastCharacter = 0
            Count = 0
            While ShowOpen.nFilesSelected = 0
                NewCharacter = InStr(LastCharacter + 1, FileDialog.sFile, Chr$(0), vbTextCompare)
                If Count > 0 Then
                    tempFiles(Count) = Mid(FileDialog.sFile, LastCharacter + 1, NewCharacter - LastCharacter - 1)
                Else
                    ShowOpen.sLastDirectory = Mid(FileDialog.sFile, LastCharacter + 1, NewCharacter - LastCharacter - 1)
                End If
                Count = Count + 1
                If InStr(NewCharacter + 1, FileDialog.sFile, Chr$(0), vbTextCompare) = InStr(NewCharacter + 1, FileDialog.sFile, Chr$(0) & Chr$(0), vbTextCompare) Then
                    tempFiles(Count) = Mid(FileDialog.sFile, NewCharacter + 1, InStr(NewCharacter + 1, FileDialog.sFile, Chr$(0) & Chr$(0), vbTextCompare) - NewCharacter - 1)
                    ShowOpen.nFilesSelected = Count
                End If
                LastCharacter = NewCharacter
            Wend
            ReDim ShowOpen.sFiles(1 To ShowOpen.nFilesSelected)
            For Count = 1 To ShowOpen.nFilesSelected
                ShowOpen.sFiles(Count) = tempFiles(Count)
            Next
        Else
            ReDim ShowOpen.sFiles(1 To 1)
            ShowOpen.sLastDirectory = left$(FileDialog.sFile, FileDialog.nFileOffset)
            ShowOpen.nFilesSelected = 1
            ShowOpen.sFiles(1) = Mid(FileDialog.sFile, FileDialog.nFileOffset + 1, InStr(1, FileDialog.sFile, Chr$(0), vbTextCompare) - FileDialog.nFileOffset - 1)
        End If
        ShowOpen.bCanceled = False
        Exit Function
    Else
        ShowOpen.sLastDirectory = ""
        ShowOpen.nFilesSelected = 0
        ShowOpen.bCanceled = True
        Erase ShowOpen.sFiles
        Exit Function
    End If
End Function

Public Function ShowSave(ByVal hwnd As Long, Optional ByVal centerForm As Boolean = True, Optional strFile As String = "", Optional strDir As String = "") As SelectedFile
Dim ret As Long
Dim hInst As Long
Dim Thread As Long
    
    ParenthWnd = hwnd
    FileDialog.nStructSize = Len(FileDialog)
    FileDialog.hwndOwner = hwnd
    FileDialog.sFileTitle = Space$(2048)
    FileDialog.nTitleSize = Len(FileDialog.sFileTitle)
    FileDialog.sFile = Space$(2047) & Chr$(0)
    Mid$(FileDialog.sFile, 1, Len(strFile)) = strFile
    FileDialog.sInitDir = strDir & Chr$(0)
    FileDialog.nFileSize = Len(FileDialog.sFile)
    
    If FileDialog.flags = 0 Then
        FileDialog.flags = OFS_FILE_SAVE_FLAGS
    End If
    
    'Set up the CBT hook
    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    If centerForm = True Then
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If
    
    ret = GetSaveFileName(FileDialog)
    ReDim Preserve ShowSave.sFiles(1)

    If ret Then
        ShowSave.sLastDirectory = left$(FileDialog.sFile, FileDialog.nFileOffset)
        ShowSave.nFilesSelected = 1
        ShowSave.sFiles(1) = Mid(FileDialog.sFile, FileDialog.nFileOffset + 1, InStr(1, FileDialog.sFile, Chr$(0), vbTextCompare) - FileDialog.nFileOffset - 1)
        ShowSave.bCanceled = False
        Exit Function
    Else
        ShowSave.sLastDirectory = ""
        ShowSave.nFilesSelected = 0
        ShowSave.bCanceled = True
        Erase ShowSave.sFiles
        Exit Function
    End If
End Function


Private Function WinProcCenterScreen(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim rectForm As RECT, rectMsg As RECT
    Dim x As Long, y As Long
    If lMsg = HCBT_ACTIVATE Then
        'Show the MsgBox at a fixed location (0,0)
        GetWindowRect wParam, rectMsg
        x = Screen.Width / Screen.TwipsPerPixelX / 2 - (rectMsg.Right - rectMsg.left) / 2
        y = Screen.Height / Screen.TwipsPerPixelY / 2 - (rectMsg.Bottom - rectMsg.top) / 2
        Debug.Print "Screen " & Screen.Height / 2
        Debug.Print "MsgBox " & (rectMsg.Right - rectMsg.left) / 2
        SetWindowPos wParam, 0, x, y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
        'Release the CBT hook
        UnhookWindowsHookEx hHook
    End If
    WinProcCenterScreen = False
End Function

Private Function WinProcCenterForm(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim rectForm As RECT, rectMsg As RECT
    Dim x As Long, y As Long
    'On HCBT_ACTIVATE, show the MsgBox centered over Form1
    If lMsg = HCBT_ACTIVATE Then
        'Get the coordinates of the form and the message box so that
        'you can determine where the center of the form is located
        GetWindowRect ParenthWnd, rectForm
        GetWindowRect wParam, rectMsg
        x = (rectForm.left + (rectForm.Right - rectForm.left) / 2) - ((rectMsg.Right - rectMsg.left) / 2)
        y = (rectForm.top + (rectForm.Bottom - rectForm.top) / 2) - ((rectMsg.Bottom - rectMsg.top) / 2)
        'Position the msgbox
        SetWindowPos wParam, 0, x, y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
        'Release the CBT hook
        UnhookWindowsHookEx hHook
     End If
     WinProcCenterForm = False
End Function
