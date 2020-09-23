Attribute VB_Name = "Module1"
'************************************************************************************
'Load Dialog From Resoucefile project written by Peter Hebels (http://www.phsoft.nl)*
'Please don't remove this message header when distributing this code in source form *
'                                                                                   *
'I've written this just for fun, don't expect anything special from it...           *
'                                                                                   *
'Don't forget you have to compile this project to an exe, otherwise it cannot read  *
'the resource file and thus cannot load the dialog from it.                         *
'                                                                                   *
'************************************************************************************

'I'll never code without it
Option Explicit

'Some API calls
Public Declare Function DialogBoxParam Lib "user32" Alias "DialogBoxParamA" (ByVal hInstance As Long, ByVal lpTemplate As Long, ByVal hWndParent As Long, ByVal lpDialogFunc As Long, ByVal dwInitParam As Long) As Long
Public Declare Function EndDialog Lib "user32" (ByVal hDlg As Long, ByVal nResult As Long) As Long

Public Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
Public Declare Function GetDlgItemText Lib "user32" Alias "GetDlgItemTextA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long

Public Declare Sub InitCommonControls Lib "comctl32" ()
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByRef lParam As WINDOWPOS) As Long

'Some constants
Public Const MB_OK = &H0&
Public Const MB_ICONINFORMATION = &H40&

Public Const WM_INITDIALOG = &H110
Public Const WM_COMMAND = &H111
Public Const BN_CLICKED = 0
Public Const WM_USER = &H400
Public Const TBM_SETRANGEMIN = (WM_USER + 7)
Public Const TBM_SETRANGEMAX = (WM_USER + 8)
Public Const TBM_GETPOS = (WM_USER)
Public Const GWL_WNDPROC = (-4)
Public Const WM_LBUTTONUP = &H202

Type WINDOWPOS
    hWnd As Long
    hWndInsertAfter As Long
    x As Long
    y As Long
    cx As Long
    cy As Long
    flags As Long
End Type

'Some variables
Dim hSlider As Long
Dim hTextBox As Long
Dim OldSliderProc As Long
Dim hDialog As Long

'Procedure for the dialog box loaded from the resource file
Public Function DialogBoxProc(ByVal hwndDlg As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim NotifyCode As Long
    Dim ItemID As Long
    Dim TmpStr As String * 255
    
    hDialog = hwndDlg
    
    'Command message filter
    If uMsg = WM_COMMAND Then
        NotifyCode = wParam \ 65536
        ItemID = wParam And 65535
        
        'OK button with resouce ID 1
        If ItemID = 1 And NotifyCode = BN_CLICKED Then
            'Show a messagebox to the user
            MessageBox hwndDlg, "You clicked the OK button", "Button clicked", MB_OK Or MB_ICONINFORMATION
            DialogBoxProc = 1
        'Cancel button with resouce ID 2
        ElseIf ItemID = 2 And NotifyCode = BN_CLICKED Then
            'Show a messagebox to the user
            MessageBox hwndDlg, "You clicked the Cancel button", "Button clicked", MB_OK Or MB_ICONINFORMATION
            DialogBoxProc = 1
        'Textbox with resouce ID 3
        ElseIf ItemID = 3 Then
            'Get the text from the textbox and copy it to Form1's textbox
            GetDlgItemText hwndDlg, str(3), TmpStr, 255
            Form1.Text1.Text = TrimZero(TmpStr)
            DialogBoxProc = 1
        'Close dialog button with resouce ID 5
        ElseIf ItemID = 5 And NotifyCode = BN_CLICKED Then
            'Close the dialog when button is clicked
            EndDialog hwndDlg, 0
            DialogBoxProc = 1
        End If
    End If
    
    'Initdialog message filter
    If uMsg = WM_INITDIALOG Then
        'Put some text into the textbox found on the dialog loaded from te resouce file
        hTextBox = GetDlgItem(hwndDlg, str(3))
        SetWindowText hTextBox, "Dialog loaded from resource file, click some buttons and move" & _
                                " the slider to see how messages are recieved by the main app."
        
        'Get the slider's hwnd
        hSlider = GetDlgItem(hwndDlg, str(4))
        
        'Set slider's max and min values
        SendMessageLong hSlider, TBM_SETRANGEMIN, False, 1
        SendMessageLong hSlider, TBM_SETRANGEMAX, False, 10
        
        'Get the text from the dialog's textbox and copy it to Form1's text box
        GetDlgItemText hwndDlg, str(3), TmpStr, 255
        Form1.Text1.Text = TrimZero(TmpStr)
        
        'Subclass the slider found on the dialog loaded from resource
        OldSliderProc = SetWindowLong(hSlider, GWL_WNDPROC, AddressOf SliderProc)
    End If
    
    DialogBoxProc = 0
End Function

'Slider procedure, recieves messages from the slider
Public Function SliderProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, lParam As WINDOWPOS) As Long
    Dim TmpStr As String * 255
    Dim SliderValue As Long
    
    'Check if left mouse button is up
    If Msg = WM_LBUTTONUP Then
        'Get the slider value and put it in the text box found on the resource dialog
        SliderValue = SendMessageLong(hSlider, TBM_GETPOS, 0, 0)
        SetWindowText hTextBox, "Slider Value is: " & SliderValue
    
        'Also put it in Form1's textbox
        GetDlgItemText hDialog, str(3), TmpStr, 255
        Form1.Text1.Text = TrimZero(TmpStr)
    End If
    
    ' Continue normal processing. VERY IMPORTANT!
    SliderProc = CallWindowProc(OldSliderProc, hSlider, Msg, wParam, lParam)

End Function

'LoadDialogRes function called when Command1 on Form1 is clicked
Public Function LoadDialogRes()
    Dim hInst As Long
    
    'Get het applications hInstance
    hInst = App.hInstance
    
    'Because we use some CommonControls on the dialog loaded from the
    'resource file, we have initialize them
    InitCommonControls
    
    'If we have a hInst value, we load the dialog (ID 101) from the resource file
    If hInst <> 0 Then
        DialogBoxParam hInst, str(101), Form1.hWnd, AddressOf DialogBoxProc, 0
    End If
End Function

'Strip zerro's from the end of a string
Private Function TrimZero(str As String) As String
    Dim lngPos As Long
    
    lngPos = InStr(str, Chr$(0))
    If lngPos > 0 Then
        TrimZero = Mid$(str, 1, lngPos - 1)
    Else
        TrimZero = str
    End If
End Function

