VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CClipboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'=================================================================
'Class for automatic copying of a string to the Windows clipboard
'=================================================================

' From Access 2000 Developer's Handbook, Volume I
' by Getz, Litwin, and Gilbert. (Sybex)
' Copyright 1999. All Rights Reserved.
    
' Clipboard class

' Declare a clipboard object at the form level:
'   Private Clipboard As New CClipboard
'   Clipboard.SetText StringX                          - Save StringX to clipboard

' Check if the data format in the buffer is available:
Private Declare PtrSafe Function IsClipboardFormatAvailable _
    Lib "user32" _
    (ByVal uFormat As Long) As Long

' Open the clipboard:
Private Declare PtrSafe Function OpenClipboard _
    Lib "user32" _
    (ByVal hWnd As LongPtr) As Long

' Get data from the clipboard:
Private Declare PtrSafe Function GetClipboardData _
    Lib "user32" _
    (ByVal uFormat As Long) As LongPtr
    
' Get the size of global memory:
Private Declare PtrSafe Function GlobalSize _
    Lib "kernel32" _
    (ByVal hMem As LongPtr) As LongPtr
    
' Block global memory:
Private Declare PtrSafe Function GlobalLock _
    Lib "kernel32" _
    (ByVal hMem As LongPtr) As LongPtr
    
' Copying memory:
Private Declare PtrSafe Sub MoveMemory _
    Lib "kernel32" Alias "RtlMoveMemory" _
    (ByVal strDest As Any, _
    ByVal lpSource As Any, _
    ByVal Length As LongPtr)
    
' Unlock global memory:
Private Declare PtrSafe Function GlobalUnlock _
    Lib "kernel32" _
    (ByVal hMem As LongPtr) As Long
    
' Close the clipboard:
Private Declare PtrSafe Function CloseClipboard _
    Lib "user32" () As Long
    
' Allocate memory in the clipboard:
Private Declare PtrSafe Function GlobalAlloc _
    Lib "kernel32" _
    (ByVal uFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    
' Clearing the clipboard:
Private Declare PtrSafe Function EmptyClipboard _
    Lib "user32" () As Long
    
' Places data on the clipboard in the specified format:
Private Declare PtrSafe Function SetClipboardData _
    Lib "user32" _
    (ByVal uFormat As Long, ByVal hData As LongPtr) As LongPtr
    
' Frees a global memory object:
Private Declare PtrSafe Function GlobalFree _
    Lib "kernel32" _
    (ByVal hMem As LongPtr) As LongPtr

Private Const GMEM_MOVABLE = &H2&
Private Const GMEM_DDESHARE = &H2000&
Private Const CF_TEXT = 1

'Error return codes from Clipboard2Text
Private Const CLIPBOARDFORMATNOTAVAILABLE = 1
Private Const CANNOTOPENCLIPBOARD = 2
Private Const CANNOTGETCLIPBOARDDATA = 3
Private Const CANNOTGLOBALLOCK = 4
Private Const CANNOTCLOSECLIPBOARD = 5
Private Const CANNOTGLOBALALLOC = 6
Private Const CANNOTEMPTYCLIPBOARD = 7
Private Const CANNOTSETCLIPBOARDDATA = 8
Private Const CANNOTGLOBALFREE = 9

Function SetText(strText As String) As Variant
    Dim varRet As Variant
    Dim fSetClipboardData As Boolean
    Dim hMemory As LongPtr
    Dim lpMemory As LongPtr
    Dim lngSize As Long

    varRet = False
    fSetClipboardData = False

    ' Get the length, including one extra for a CHR$(0)
    ' at the end.
    lngSize = Len(strText) + 1
    hMemory = GlobalAlloc(GMEM_MOVABLE Or _
        GMEM_DDESHARE, lngSize)
    If Not CBool(hMemory) Then
        varRet = CVErr(CANNOTGLOBALALLOC)
        GoTo SetTextDone
    End If

    ' Lock the object into memory
    lpMemory = GlobalLock(hMemory)
    If Not CBool(lpMemory) Then
        varRet = CVErr(CANNOTGLOBALLOCK)
        GoTo SetTextGlobalFree
    End If

    ' Move the string into the memory we locked
    Call MoveMemory(lpMemory, strText, lngSize)

    ' Don't send clipboard locked memory.
    Call GlobalUnlock(hMemory)

    ' Open the clipboard
    If Not CBool(OpenClipboard(0&)) Then
        varRet = CVErr(CANNOTOPENCLIPBOARD)
        GoTo SetTextGlobalFree
    End If

    ' Remove the current contents of the clipboard
    If Not CBool(EmptyClipboard()) Then
        varRet = CVErr(CANNOTEMPTYCLIPBOARD)
        GoTo SetTextCloseClipboard
    End If

    ' Add our string to the clipboard as text
    If Not CBool(SetClipboardData(CF_TEXT, _
        hMemory)) Then
        varRet = CVErr(CANNOTSETCLIPBOARDDATA)
        GoTo SetTextCloseClipboard
    Else
        fSetClipboardData = True
    End If

SetTextCloseClipboard:
    ' Close the clipboard
    If Not CBool(CloseClipboard()) Then
        varRet = CVErr(CANNOTCLOSECLIPBOARD)
    End If

SetTextGlobalFree:
    If Not fSetClipboardData Then
        'If we have set the clipboard data, we no longer own
        ' the object--Windows does, so don't free it.
        If CBool(GlobalFree(hMemory)) Then
            varRet = CVErr(CANNOTGLOBALFREE)
        End If
    End If

SetTextDone:
    SetText = varRet
End Function

Public Function GetText() As Variant
    Dim hMemory As Long
    Dim lpMemory As Long
    Dim strText As String
    Dim lngSize As Long
    Dim varRet As Variant

    varRet = ""

    ' Is there text on the clipboard? If not, error out.
    If Not CBool(IsClipboardFormatAvailable _
        (CF_TEXT)) Then
        varRet = CVErr(CLIPBOARDFORMATNOTAVAILABLE)
        GoTo GetTextDone
    End If

    ' Open the clipboard
    If Not CBool(OpenClipboard(0&)) Then
        varRet = CVErr(CANNOTOPENCLIPBOARD)
        GoTo GetTextDone
    End If

    ' Get the handle to the clipboard data
    hMemory = GetClipboardData(CF_TEXT)
    If Not CBool(hMemory) Then
        varRet = CVErr(CANNOTGETCLIPBOARDDATA)
        GoTo GetTextCloseClipboard
    End If

    ' Find out how big it is and allocate enough space
    ' in a string
    lngSize = GlobalSize(hMemory)
    strText = Space$(lngSize)

    ' Lock the handle so we can use it
    lpMemory = GlobalLock(hMemory)
    If Not CBool(lpMemory) Then
        varRet = CVErr(CANNOTGLOBALLOCK)
        GoTo GetTextCloseClipboard
    End If

    ' Move the information from the clipboard memory
    ' into our string
    Call MoveMemory(strText, lpMemory, lngSize)

    ' Truncate it at the first Null character because
    ' the value reported by lngSize is erroneously large
    strText = Left$(strText, InStr(1, strText, Chr$(0)) - 1)

    ' Free the lock
    Call GlobalUnlock(hMemory)

GetTextCloseClipboard:
    ' Close the clipboard
    If Not CBool(CloseClipboard()) Then
        varRet = CVErr(CANNOTCLOSECLIPBOARD)
    End If

GetTextDone:
    If Not IsError(varRet) Then
        GetText = strText
    Else
        GetText = varRet
    End If
End Function

Public Function ErrorText(ByVal lngError As Long) As String
    ' Reports an error received from the clipboard
    '
    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert. (Sybex)
    ' Copyright 1999. All Rights Reserved.
    '
    
    Select Case CInt(lngError)
        Case CLIPBOARDFORMATNOTAVAILABLE
            ErrorText = "Clipboard format not available"
        Case CANNOTOPENCLIPBOARD
            ErrorText = "Cannot open clipboard"
        Case CANNOTGETCLIPBOARDDATA
            ErrorText = "Cannot get clipboard data"
        Case CANNOTGLOBALLOCK
            ErrorText = "Cannot global lock data"
        Case CANNOTCLOSECLIPBOARD
            ErrorText = "Cannot close clipboard"
        Case CANNOTGLOBALALLOC
            ErrorText = "Cannot global alloc"
        Case CANNOTEMPTYCLIPBOARD
            ErrorText = "Cannot empty clipboard"
        Case CANNOTSETCLIPBOARDDATA
            ErrorText = "Cannot set clipboard data"
        Case CANNOTGLOBALFREE
            ErrorText = "Cannot global free"
        Case Else
            ErrorText = "Unknown error"
    End Select
End Function

