Attribute VB_Name = "Module2"
'This code checks if the app is running in the VB IDE, this module come's from Programmer's Corner

Private Declare Function EnumThreadWindows Lib "user32" (ByVal dwThreadId As Long, ByVal lpEnumFunc As Long, ByRef lParam As Long) As Long
Private Declare Function GetWindowClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nBufLen As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Private VBIDEs() As Long
Const m_0 As Long = 0
Const m_1 As Long = 1

' ****************************************************************
'
'   Use AlertErr.bas instead! Unless you already have error
'   handling sorted use AlertErr.bas: it is this module plus.
'
'   This module determines if it is running within an instance
'   of the VB Development Environment, or within a stand-alone
'   executable.
'
'   If running as a stand-alone executable the RunningInVbIDE
'   function returns zero.
'
'   If running within an instance of the VB IDE RunningInVbIDE
'   returns the window handle (hWnd) of the Main VB window.
'
' ****************************************************************

Option Explicit

Public Function RunningInVbIDE() As Long
    On Error GoTo ErrHandler
    Dim rc As Long, VBIDEsCount As Long

    ' Search all current thread windows for the VB IDE main window
    rc = EnumThreadWindows(App.ThreadID, AddressOf CallBackIDE, VBIDEsCount)

    ' If the IDE is running
    If (VBIDEsCount) Then
        Dim VBProcessID As Long, MeProcessID As Long
        Dim VBThreadID As Long, Idx As Long

        ' Get this components's Process ID
        MeProcessID = GetCurrentProcessId

        For Idx = m_1 To VBIDEsCount

            ' Get VB's Process ID
            VBThreadID = GetWindowThreadProcessId(VBIDEs(Idx), VBProcessID)

            ' If running in the same process
            If (VBProcessID = MeProcessID) Then
                RunningInVbIDE = VBIDEs(Idx) ' ©Rd
                Exit For
            End If

        Next Idx
    End If
ErrHandler:
End Function

Private Function CallBackIDE(ByVal hWnd As Long, ByRef lCount As Long) As Long
    On Error GoTo ErrHandler

    ' Default to Enum the next window
    CallBackIDE = m_1

    ' If it's a VB IDE instance
    If (GetClassName(hWnd) = "IDEOwner") Then

        lCount = lCount + m_1
        ReDim Preserve VBIDEs(m_1 To lCount) As Long

        ' Record the window handle
        VBIDEs(lCount) = hWnd

    End If
    Exit Function
ErrHandler:
    ' On error cancel callback
    CallBackIDE = m_0
End Function

Private Function GetClassName(ByVal hWnd As Long) As String
    GetClassName = "unknown"
    On Error GoTo ErrHandler
    Dim ClassName As String

    ' Allow ample length for the class name
    ClassName = String$(255, vbNullChar)

    If (GetWindowClassName(hWnd, ClassName, Len(ClassName))) Then
        GetClassName = Left$(ClassName, InStr(ClassName, vbNullChar) - m_1)
    End If
ErrHandler:
End Function

