Attribute VB_Name = "CheckAccess"
Option Compare Database
Option Explicit
' Declare a clipboard object at the form level:
Private Clipboard As New CClipboard

' Function #1     : CheckAccessBitness
' Purpose         : Checks whether the user is running a 32-bit or 64-bit version of Microsoft Access
'                   and provides a warning if 32-bit is detected.
' Behavior        :
'                  - Calls Is64BitAccess() to determine the Access bitness.
'                  - If 64-bit, the procedure exits without any warning.
'                  - If 32-bit, it displays a message box with steps to manually register the `mscomctl.ocx` file.
' External Calls  :
'                  - Is64BitAccess()    : Determines if Access is running in 64-bit mode.
'                  - GetMSCOMCTLPath()  : Retrieves the file path of `mscomctl.ocx`.
' Notes           :
'                  - This function is useful when working with ActiveX controls that may not be compatible with 32-bit Access.
'                  - The registration steps must be executed with Administrator privileges.
' --------------------------------------------------------------------------
Public Function CheckAccessBitness()
    Dim msg As String

    If Is64BitAccess() Then         ' - Check Access bitness
        Exit Function               ' - If 64-bit, no issues, continue
    Else
        ' If 32-bit, warn the user
        msg = "You are using the 32-bit version of Microsoft Access!" & vbCrLf & vbCrLf & _
              "ActiveX controls (`mscomctl.ocx`), such as TreeView, ListView, ImageList, ProgressBar, Slider (TrackBar), and StatusBar, may not work properly." & vbCrLf & _
              "To resolve this issue, follow these steps:" & vbCrLf & vbCrLf & _
              "1. Make sure the file `mscomctl.ocx` is located at `" & GetMSCOMCTLPath() & "`" & vbCrLf & _
              "2. Open the Command Prompt (CMD) as Administrator." & vbCrLf & _
              "3. Run the following command: `regsvr32 """ & GetMSCOMCTLPath() & """`" & vbCrLf & _
              "4. Restart Microsoft Access." & vbCrLf & vbCrLf & _
              "If the issue persists, consider installing the 64-bit version of Microsoft Office."

        MsgBox msg, vbExclamation, "Warning: 32-bit Access detected"
        
        Clipboard.SetText msg       ' - Save message to clipboard
        
        MsgBox "The message is written to the clipboard...", vbInformation + vbOKOnly, "Message in clipboard"
    End If
    
End Function

' --------------------------------------------------------------------------
' Function #2    : Is64BitAccess
' Purpose        : Determines whether the current instance of Microsoft Access is running in 64-bit mode.
' Returns        : Boolean - True if running in 64-bit, False otherwise.
' Behavior       :
'                  - Uses the `#If Win64 Then` compiler directive to check the architecture at compile time.
'                  - Returns True for 64-bit Access and False for 32-bit Access.
' Notes          :
'                  - This function is useful when working with ActiveX controls that require specific bitness compatibility.
' --------------------------------------------------------------------------
Public Function Is64BitAccess() As Boolean
    #If Win64 Then
        Is64BitAccess = True
    #Else
        Is64BitAccess = False
    #End If
End Function

' --------------------------------------------------------------------------
' Function #3   : GetMSCOMCTLPath
' Purpose        : Retrieves the expected file path of the `mscomctl.ocx` ActiveX control for 32-bit systems.
' Returns        : String - The full file path of `mscomctl.ocx`.
' Behavior       :
'                  - Uses the `Environ("SystemRoot")` function to get the system root path.
'                  - Appends `\SysWow64\mscomctl.ocx` for compatibility with 32-bit systems.
' Notes          :
'                  - This function assumes a standard system directory structure.
'                  - If the file is missing, manual verification may be required.
' --------------------------------------------------------------------------
Public Function GetMSCOMCTLPath() As String
    ' Standard path to 32-bit `mscomctl.ocx`
    GetMSCOMCTLPath = Environ("SystemRoot") & "\SysWow64\mscomctl.ocx"
End Function

